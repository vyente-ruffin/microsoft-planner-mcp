# microsoft-planner-mcp

Lightweight MCP (Model Context Protocol) server for Microsoft Planner that uses `az rest` for authentication instead of complex OAuth flows.

## What This Is

A minimal MCP server that enables Claude Code to interact with Microsoft Planner tasks directly. It leverages Azure CLI's existing authentication (`az login`) to make Graph API calls, eliminating the need for app registrations or token management.

## Prerequisites

- **Node.js** (v18 or later)
- **Azure CLI** installed and authenticated:
  ```bash
  az login
  ```

## Installation

```bash
git clone https://github.com/vyente-ruffin/microsoft-planner-mcp.git
cd microsoft-planner-mcp
npm install
npm run build
```

## Usage with Claude Code

Add the MCP server to Claude Code:

```bash
claude mcp add microsoft-planner-mcp node /path/to/microsoft-planner-mcp/dist/index.js
```

Or add directly to your MCP settings file:

```json
{
  "mcpServers": {
    "microsoft-planner-mcp": {
      "command": "node",
      "args": ["/path/to/microsoft-planner-mcp/dist/index.js"]
    }
  }
}
```

## Available Tools

### Plans

| Tool | Description |
|------|-------------|
| `list-plans` | List all Planner plans accessible to the current user |
| `get-plan-details` | Get plan details, including the label (category) names |
| `update-plan-categories` | Set label names of a plan, e.g. `{ "category1": "Review", "category2": "Stuck" }` |

### Buckets

| Tool | Description |
|------|-------------|
| `list-buckets` | List all buckets in a plan |
| `create-bucket` | Create a bucket in a plan (optional `orderHint`; placed first by default) |
| `update-bucket` | Rename or reorder a bucket |
| `delete-bucket` | Delete a bucket **including its tasks** |

### Tasks

| Tool | Description |
|------|-------------|
| `list-tasks` | List all tasks in a Planner plan |
| `get-task` | Get details of a specific task |
| `get-task-details` | Get extended task details (description, checklist, references) |
| `create-task` | Create a task in one call, optionally with progress, priority, start/due dates, assignees, labels, description and checklist (max. 20 items) |
| `update-task` | Update title, progress, priority, start/due dates, bucket (move), assignees (add/remove) and labels (add/remove) |
| `update-task-details` | Update the description and/or append checklist items; toggle checklist preview on the card |
| `set-checklist-item` | Check/uncheck, rename or delete a single checklist item by its ID |
| `delete-task` | Delete a Planner task |

### Users

| Tool | Description |
|------|-------------|
| `find-user` | Find Entra ID users by display name prefix or e-mail (returns `id`, `displayName`, `mail`) – use the `id` for assignments |

### Parameter notes

- **Priority**: `urgent`, `important`, `medium`, `low` or a number 0–10 (1 = urgent, 3 = important, 5 = medium, 9 = low).
- **Dates**: ISO 8601, e.g. `2026-03-09T00:00:00Z`.
- **Labels**: `category1` … `category25`; use `get-plan-details` to see their names.
- **Progress** (`percentComplete`): `0` = not started, `50` = in progress, `100` = completed.

## Claude Code Agent (Optional)

For enhanced automation, create a Claude Code agent at `~/.claude/agents/microsoft-planner.md` with:

- **Auto-assignment**: Automatically assign tasks to a default user
- **Auto-labeling**: Apply default categories/labels to tasks
- **Status Intelligence**: Infer task bucket and progress from conversation context
- **GitHub Integration**: Include repository links in task descriptions

See the [agent template](https://github.com/vyente-ruffin/microsoft-planner-mcp/wiki/Agent-Configuration) for configuration details.

## Finding Your IDs

To use the MCP tools, you'll need your Planner Plan ID and Bucket IDs:

1. Use `list-plans` to discover all accessible plans and their IDs
2. Use `list-buckets` with your Plan ID to get bucket IDs
3. Or find your Plan ID from the Planner web URL: `https://tasks.office.com/...planId=YOUR_PLAN_ID`

## Example Usage

Once configured, use natural language with Claude Code:

```
"List all my planner plans"
"List all tasks in the planner"
"Create a task called 'Review PR #123' in the backlog bucket"
"Mark task XYZ as complete"
"Add a description with the GitHub PR link to the task"
"Create a bucket 'In Review' and move task XYZ into it"
"Create an urgent task due Friday, assigned to Jane Doe, with a 3-item checklist"
"Tick off the first checklist item of task XYZ"
```

## How It Works

This server uses `az rest` to make Microsoft Graph API calls. The Azure CLI handles all authentication, so as long as you're logged in with `az login`, the server can access Planner data your account has permissions for.

All Planner operations that require ETags (update, delete) automatically fetch the current ETag before making changes.

Request bodies are written to a temporary UTF-8 file and passed to `az rest` via `--body @file`, so umlauts, quotes, ampersands and other special characters survive shell quoting on Windows (`cmd.exe`) and Unix shells alike.

## License

ISC
