# mcp-microsoft-todo

Microsoft To Do (Microsoft 365) MCP Pack

Part of [Pipeworx](https://pipeworx.io) — an MCP gateway connecting AI agents to 1481+ live data sources.

## Tools

| Tool | Description |
|------|-------------|
| `list_task_lists` | List the Microsoft To Do task lists for the signed-in user, with display names and whether each is a well-known list (e.g. the default "Tasks" or flagged-email list). Use to discover a to-do list and its ID before listing tasks or reminders. |
| `list_tasks` | List the tasks in a Microsoft To Do list. Returns compact task summaries (title, status, importance, due date, created/completed times, and a body preview). Optionally filter by status. Use to browse a user's to-do list and see what tasks or reminders it contains. |
| `get_task` | Get the full details of a single Microsoft To Do task by its list ID and task ID, including title, status, importance, full body/notes, due date, reminder time, created/completed times, and any checklist (sub-task) items. Use after list_tasks or find_due_tasks to read a to-do item in full. |
| `find_due_tasks` | Find upcoming, incomplete Microsoft To Do tasks across all of the user's to-do lists that are due within the next N days. Returns a flat list of due reminders with their list name, title, due date, and importance. Use to answer "what tasks are due soon" or "what are my upcoming reminders". |

## Quick Start

Add to your MCP client (Claude Desktop, Cursor, Windsurf, etc.):

```json
{
  "mcpServers": {
    "microsoft_todo": {
      "url": "https://gateway.pipeworx.io/microsoft_todo/mcp"
    }
  }
}
```

### What this endpoint actually serves

`tools/list` at `https://gateway.pipeworx.io/microsoft_todo/mcp` returns the tools in the table
above **plus the shared Pipeworx meta-tools** — `ask_pipeworx`,
`discover_tools`, `search_within`, `remember`/`recall` and the rest of the
gateway-wide set. So the tool count you see is larger than this table: a
single-pack endpoint currently lists roughly 30 shared tools alongside the
pack's own. The connection's `initialize` response states its exact scope, and
is the authoritative answer for a given day.

This is deliberate, not multiplexing by accident. The meta-tools are what let a
scoped connection answer a question this pack does not cover — via
`ask_pipeworx`, which routes across the whole catalog — without you adding a
second MCP server. There is currently no way to mount a pack endpoint without
them; if the extra schemas cost you more context than the routing is worth,
connect to the full gateway once rather than to several pack endpoints.

Or connect to the full Pipeworx gateway to get every pack's tools listed
directly, instead of just this one's:

```json
{
  "mcpServers": {
    "pipeworx": {
      "url": "https://gateway.pipeworx.io/mcp"
    }
  }
}
```

Both URLs reach the same gateway and the same 1481+ data sources. The
only difference is which pack's tools are listed **directly**; `ask_pipeworx`
reaches all of them from either one.

## Using with ask_pipeworx

Instead of calling tools directly, you can ask questions in plain English —
this works on the pack endpoint above as well as on the full gateway:

```
ask_pipeworx({ question: "your question about Microsoft Todo data" })
```

The gateway picks the right tool and fills the arguments automatically.

## More

- [Docs and guides](https://pipeworx.io/docs)
- [pipeworx.io](https://pipeworx.io)

## License

MIT
