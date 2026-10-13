import { FastMCP } from "fastmcp";
import { z } from "zod";
import { execSync } from "child_process";
import { writeFileSync, unlinkSync } from "fs";
import { tmpdir } from "os";
import { join } from "path";
import { randomUUID } from "crypto";

const mcp = new FastMCP({
  name: "microsoft-planner-mcp",
  version: "1.1.0",
});

const GRAPH = "https://graph.microsoft.com/v1.0";

/**
 * Execute an az rest call. The JSON body is written to a temp file (UTF-8) and passed
 * via --body @file, so umlauts, quotes, ampersands etc. survive cmd.exe / sh quoting.
 */
function azRest(method: string, url: string, body?: object, headers: Record<string, string> = {}): string {
  const args = [`az rest --method ${method} --url "${url}"`];
  let tmpFile: string | undefined;
  const hdrs: Record<string, string> = { ...headers };
  if (body) {
    tmpFile = join(tmpdir(), `planner-mcp-${randomUUID()}.json`);
    writeFileSync(tmpFile, JSON.stringify(body), { encoding: "utf-8" });
    hdrs["Content-Type"] = "application/json";
    args.push(`--body "@${tmpFile}"`);
  }
  const hdrArgs = Object.entries(hdrs).map(([k, v]) => `"${k}=${v.replace(/"/g, '\\"')}"`);
  if (hdrArgs.length) args.push(`--headers ${hdrArgs.join(" ")}`);
  try {
    return execSync(args.join(" "), {
      encoding: "utf-8",
      maxBuffer: 20 * 1024 * 1024,
      env: { ...process.env, PYTHONIOENCODING: "utf-8" },
    });
  } catch (error: any) {
    throw new Error(`az rest failed: ${error.message}`);
  } finally {
    if (tmpFile) {
      try { unlinkSync(tmpFile); } catch { /* ignore */ }
    }
  }
}

function getJson(url: string): any {
  return JSON.parse(azRest("GET", url));
}

function etagOf(url: string): string {
  return getJson(url)["@odata.etag"];
}

/** PATCH with auto-fetched ETag. Planner returns 204 by default; ask for representation. */
function patchWithEtag(url: string, body: object): string {
  const etag = etagOf(url);
  const res = azRest("PATCH", url, body, { "If-Match": etag, Prefer: "return=representation" });
  return res || "OK";
}

const PRIORITY: Record<string, number> = { urgent: 1, important: 3, medium: 5, low: 9 };
const prioritySchema = z
  .union([z.enum(["urgent", "important", "medium", "low"]), z.number().int().min(0).max(10)])
  .optional()
  .describe("Priority: urgent|important|medium|low or 0-10 (1=urgent, 3=important, 5=medium, 9=low)");

function toPriority(p: unknown): number | undefined {
  if (p === undefined) return undefined;
  return typeof p === "number" ? p : PRIORITY[p as string];
}

function assignmentsFor(userIds: string[] | undefined, remove: string[] = []): Record<string, any> | undefined {
  if ((!userIds || userIds.length === 0) && remove.length === 0) return undefined;
  const a: Record<string, any> = {};
  for (const id of userIds ?? []) a[id] = { "@odata.type": "#microsoft.graph.plannerAssignment", orderHint: " !" };
  for (const id of remove) a[id] = null;
  return a;
}

function categoriesFor(add: string[] | undefined, remove: string[] = []): Record<string, boolean> | undefined {
  if ((!add || add.length === 0) && remove.length === 0) return undefined;
  const c: Record<string, boolean> = {};
  for (const k of add ?? []) c[k] = true;
  for (const k of remove) c[k] = false;
  return c;
}

const checklistItemSchema = z.object({
  title: z.string().describe("Checklist item text (max 100 chars, longer is truncated)"),
  isChecked: z.boolean().optional().describe("Checked/done"),
});

function checklistPatch(items: { title: string; isChecked?: boolean }[]): Record<string, any> {
  const out: Record<string, any> = {};
  // orderHint " !" places each new item... we build ascending hints so the given order is kept
  let prev = "";
  for (const it of items) {
    const hint = prev ? `${prev} !` : " !";
    out[randomUUID()] = {
      "@odata.type": "microsoft.graph.plannerChecklistItem",
      title: it.title.length > 100 ? it.title.slice(0, 97) + "..." : it.title,
      isChecked: it.isChecked ?? false,
      orderHint: hint,
    };
    prev = hint;
  }
  return out;
}

// ---------------------------------------------------------------- Plans

mcp.addTool({
  name: "list-plans",
  description: "List all Planner plans accessible to the current user",
  parameters: z.object({}),
  execute: async () => JSON.stringify(getJson(`${GRAPH}/me/planner/plans`).value, null, 2),
});

mcp.addTool({
  name: "get-plan-details",
  description: "Get plan details incl. label (category) names",
  parameters: z.object({ planId: z.string() }),
  execute: async ({ planId }) => azRest("GET", `${GRAPH}/planner/plans/${planId}/details`),
});

mcp.addTool({
  name: "update-plan-categories",
  description: "Set label names of a plan, e.g. {category1: 'Review', category2: 'Stuck'}",
  parameters: z.object({
    planId: z.string(),
    categoryDescriptions: z.record(z.string(), z.string()).describe("Map category1..category25 -> label name"),
  }),
  execute: async ({ planId, categoryDescriptions }) =>
    patchWithEtag(`${GRAPH}/planner/plans/${planId}/details`, { categoryDescriptions }),
});

// ---------------------------------------------------------------- Buckets

mcp.addTool({
  name: "list-buckets",
  description: "List all buckets in a Planner plan",
  parameters: z.object({ planId: z.string().describe("The Planner plan ID") }),
  execute: async ({ planId }) => JSON.stringify(getJson(`${GRAPH}/planner/plans/${planId}/buckets`).value, null, 2),
});

mcp.addTool({
  name: "create-bucket",
  description: "Create a bucket in a plan. New buckets are placed first unless orderHint is given.",
  parameters: z.object({
    planId: z.string(),
    name: z.string(),
    orderHint: z.string().optional().describe("Optional order hint (' !' = default)"),
  }),
  execute: async ({ planId, name, orderHint }) =>
    azRest("POST", `${GRAPH}/planner/buckets`, { planId, name, orderHint: orderHint ?? " !" }),
});

mcp.addTool({
  name: "update-bucket",
  description: "Rename / reorder a bucket. Auto-fetches ETag.",
  parameters: z.object({ bucketId: z.string(), name: z.string().optional(), orderHint: z.string().optional() }),
  execute: async ({ bucketId, name, orderHint }) => {
    const body: Record<string, any> = {};
    if (name !== undefined) body.name = name;
    if (orderHint !== undefined) body.orderHint = orderHint;
    return patchWithEtag(`${GRAPH}/planner/buckets/${bucketId}`, body);
  },
});

mcp.addTool({
  name: "delete-bucket",
  description: "Delete a bucket (and its tasks!). Auto-fetches ETag.",
  parameters: z.object({ bucketId: z.string() }),
  execute: async ({ bucketId }) => {
    const url = `${GRAPH}/planner/buckets/${bucketId}`;
    azRest("DELETE", url, undefined, { "If-Match": etagOf(url) });
    return "Bucket deleted";
  },
});

// ---------------------------------------------------------------- Users

mcp.addTool({
  name: "find-user",
  description: "Find Entra ID users by display name start or e-mail (returns id, displayName, mail)",
  parameters: z.object({ query: z.string().describe("Display name prefix or e-mail") }),
  execute: async ({ query }) => {
    const q = query.replace(/'/g, "''");
    const filter = query.includes("@")
      ? `mail eq '${q}' or userPrincipalName eq '${q}'`
      : `startswith(displayName,'${q}')`;
    const url = `${GRAPH}/users?$select=id,displayName,mail,userPrincipalName&$filter=${encodeURIComponent(filter)}`;
    return JSON.stringify(getJson(url).value, null, 2);
  },
});

// ---------------------------------------------------------------- Tasks

mcp.addTool({
  name: "list-tasks",
  description: "List all tasks in a Planner plan",
  parameters: z.object({ planId: z.string().describe("The Planner plan ID") }),
  execute: async ({ planId }) => JSON.stringify(getJson(`${GRAPH}/planner/plans/${planId}/tasks`).value, null, 2),
});

mcp.addTool({
  name: "get-task",
  description: "Get details of a specific Planner task",
  parameters: z.object({ taskId: z.string().describe("The task ID") }),
  execute: async ({ taskId }) => azRest("GET", `${GRAPH}/planner/tasks/${taskId}`),
});

mcp.addTool({
  name: "get-task-details",
  description: "Get extended task details including description and checklist",
  parameters: z.object({ taskId: z.string().describe("The task ID") }),
  execute: async ({ taskId }) => azRest("GET", `${GRAPH}/planner/tasks/${taskId}/details`),
});

const taskFields = {
  title: z.string().optional().describe("Task title"),
  percentComplete: z.number().min(0).max(100).optional().describe("0 = not started, 50 = in progress, 100 = completed"),
  priority: prioritySchema,
  startDateTime: z.string().optional().describe("ISO date/time, e.g. 2026-03-09T00:00:00Z"),
  dueDateTime: z.string().optional().describe("ISO date/time"),
  assignUserIds: z.array(z.string()).optional().describe("Entra user IDs to assign"),
  categories: z.array(z.string()).optional().describe("Labels to apply (category1..category25)"),
};

mcp.addTool({
  name: "create-task",
  description:
    "Create a task in a plan/bucket, optionally with progress, priority, dates, assignees, labels, " +
    "description (notes) and checklist in one call. Returns the created task incl. id.",
  parameters: z.object({
    planId: z.string().describe("The plan ID"),
    bucketId: z.string().describe("The bucket ID"),
    ...taskFields,
    title: z.string().describe("Task title"),
    description: z.string().optional().describe("Notes / description (plain text)"),
    checklist: z.array(checklistItemSchema).max(20).optional().describe("Checklist items (Planner max 20)"),
    showChecklistOnCard: z.boolean().optional().describe("Show checklist on card (default true if checklist given)"),
  }),
  execute: async (a) => {
    const body: Record<string, any> = { planId: a.planId, bucketId: a.bucketId, title: a.title };
    if (a.percentComplete !== undefined) body.percentComplete = a.percentComplete;
    const prio = toPriority(a.priority);
    if (prio !== undefined) body.priority = prio;
    if (a.startDateTime) body.startDateTime = a.startDateTime;
    if (a.dueDateTime) body.dueDateTime = a.dueDateTime;
    const asg = assignmentsFor(a.assignUserIds);
    if (asg) body.assignments = asg;
    const cats = categoriesFor(a.categories);
    if (cats) body.appliedCategories = cats;

    const task = JSON.parse(azRest("POST", `${GRAPH}/planner/tasks`, body));

    if (a.description || a.checklist?.length) {
      const details: Record<string, any> = {};
      if (a.description) details.description = a.description;
      if (a.checklist?.length) details.checklist = checklistPatch(a.checklist);
      if (a.checklist?.length && a.showChecklistOnCard !== false) details.previewType = "checklist";
      // details object is created asynchronously; retry briefly on 404/412
      let lastErr: any;
      for (let i = 0; i < 5; i++) {
        try {
          patchWithEtag(`${GRAPH}/planner/tasks/${task.id}/details`, details);
          lastErr = undefined;
          break;
        } catch (e) {
          lastErr = e;
          execSync(process.platform === "win32" ? "ping -n 3 127.0.0.1 >NUL" : "sleep 2");
        }
      }
      if (lastErr) return JSON.stringify({ id: task.id, title: task.title, warning: `Task created, details failed: ${lastErr.message}` });
    }
    return JSON.stringify({ id: task.id, title: task.title, bucketId: task.bucketId, percentComplete: task.percentComplete, priority: task.priority });
  },
});

mcp.addTool({
  name: "update-task",
  description: "Update task properties (title, progress, priority, dates, assignees, labels, bucket). Auto-fetches ETag.",
  parameters: z.object({
    taskId: z.string().describe("The task ID"),
    ...taskFields,
    bucketId: z.string().optional().describe("Move to bucket"),
    assignUserId: z.string().optional().describe("(legacy) single user ID to assign"),
    category: z.string().optional().describe("(legacy) single label category1-category25"),
    unassignUserIds: z.array(z.string()).optional(),
    removeCategories: z.array(z.string()).optional(),
  }),
  execute: async (a) => {
    const body: Record<string, any> = {};
    if (a.title !== undefined) body.title = a.title;
    if (a.percentComplete !== undefined) body.percentComplete = a.percentComplete;
    const prio = toPriority(a.priority);
    if (prio !== undefined) body.priority = prio;
    if (a.startDateTime !== undefined) body.startDateTime = a.startDateTime;
    if (a.dueDateTime !== undefined) body.dueDateTime = a.dueDateTime;
    if (a.bucketId) body.bucketId = a.bucketId;
    const users = [...(a.assignUserIds ?? []), ...(a.assignUserId ? [a.assignUserId] : [])];
    const asg = assignmentsFor(users, a.unassignUserIds ?? []);
    if (asg) body.assignments = asg;
    const cats = categoriesFor([...(a.categories ?? []), ...(a.category ? [a.category] : [])], a.removeCategories ?? []);
    if (cats) body.appliedCategories = cats;
    return patchWithEtag(`${GRAPH}/planner/tasks/${a.taskId}`, body);
  },
});

mcp.addTool({
  name: "update-task-details",
  description: "Update task description (notes) and/or append checklist items. Auto-fetches ETag.",
  parameters: z.object({
    taskId: z.string().describe("The task ID"),
    description: z.string().optional().describe("Task description / notes"),
    addChecklistItems: z.array(checklistItemSchema).optional().describe("Checklist items to append"),
    showChecklistOnCard: z.boolean().optional(),
  }),
  execute: async ({ taskId, description, addChecklistItems, showChecklistOnCard }) => {
    const body: Record<string, any> = {};
    if (description !== undefined) body.description = description;
    if (addChecklistItems?.length) body.checklist = checklistPatch(addChecklistItems);
    if (showChecklistOnCard !== undefined) body.previewType = showChecklistOnCard ? "checklist" : "automatic";
    return patchWithEtag(`${GRAPH}/planner/tasks/${taskId}/details`, body);
  },
});

mcp.addTool({
  name: "set-checklist-item",
  description: "Check/uncheck, rename or delete a checklist item by its id. Auto-fetches ETag.",
  parameters: z.object({
    taskId: z.string(),
    itemId: z.string(),
    isChecked: z.boolean().optional(),
    title: z.string().optional(),
    delete: z.boolean().optional(),
  }),
  execute: async ({ taskId, itemId, isChecked, title, delete: del }) => {
    const item: any = del ? null : { "@odata.type": "microsoft.graph.plannerChecklistItem" };
    if (!del) {
      if (isChecked !== undefined) item.isChecked = isChecked;
      if (title !== undefined) item.title = title;
    }
    return patchWithEtag(`${GRAPH}/planner/tasks/${taskId}/details`, { checklist: { [itemId]: item } });
  },
});

mcp.addTool({
  name: "delete-task",
  description: "Delete a Planner task. Auto-fetches ETag.",
  parameters: z.object({ taskId: z.string().describe("The task ID to delete") }),
  execute: async ({ taskId }) => {
    const url = `${GRAPH}/planner/tasks/${taskId}`;
    azRest("DELETE", url, undefined, { "If-Match": etagOf(url) });
    return "Task deleted successfully";
  },
});

mcp.start({ transportType: "stdio" });
