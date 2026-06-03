interface McpToolDefinition {
  name: string;
  description: string;
  inputSchema: {
    type: 'object';
    properties: Record<string, unknown>;
    required?: string[];
  };
}

interface McpToolExport {
  tools: McpToolDefinition[];
  callTool: (name: string, args: Record<string, unknown>) => Promise<unknown>;
  meter?: { credits: number };
  cost?: Record<string, unknown>;
  provider?: string;
}

/**
 * Microsoft To Do (Microsoft 365) MCP Pack
 *
 * Requires OAuth connection — gateway injects credentials via _context.microsoft.
 * Read-only access to Microsoft To Do tasks via Microsoft Graph v1.0.
 * Tools: list task lists, list tasks, get task, find due tasks.
 */


interface MicrosoftTodoContext {
  microsoft?: { accessToken: string };
}

const API = 'https://graph.microsoft.com/v1.0';

/**
 * Fetch helper for Microsoft Graph.
 * - Returns { error: 'connection_required' } when no OAuth token is present.
 * - Returns { error: <status>, message: <body text> } on non-2xx responses.
 * - Otherwise returns the parsed JSON body.
 */
async function gFetch(
  ctx: MicrosoftTodoContext,
  url: string,
  options: RequestInit = {},
): Promise<unknown> {
  if (!ctx.microsoft) {
    return {
      error: 'connection_required',
      message: 'Connect your Microsoft 365 account at https://pipeworx.io/account',
    };
  }
  const res = await fetch(url, {
    ...options,
    headers: {
      Authorization: `Bearer ${ctx.microsoft.accessToken}`,
      'Content-Type': 'application/json',
      ...(options.headers ?? {}),
    },
  });
  if (!res.ok) {
    const text = await res.text();
    return { error: res.status, message: text };
  }
  return res.json();
}

const tools: McpToolExport['tools'] = [
  {
    name: 'list_task_lists',
    description:
      'List the Microsoft To Do task lists for the signed-in user, with display names and whether each is a well-known list (e.g. the default "Tasks" or flagged-email list). Use to discover a to-do list and its ID before listing tasks or reminders.',
    inputSchema: {
      type: 'object' as const,
      properties: {},
      required: [],
    },
  },
  {
    name: 'list_tasks',
    description:
      'List the tasks in a Microsoft To Do list. Returns compact task summaries (title, status, importance, due date, created/completed times, and a body preview). Optionally filter by status. Use to browse a user\'s to-do list and see what tasks or reminders it contains.',
    inputSchema: {
      type: 'object' as const,
      properties: {
        list_id: {
          type: 'string',
          description: 'The ID of the Microsoft To Do task list to read (from list_task_lists).',
        },
        top: {
          type: 'number',
          description: 'Maximum number of tasks to return (default 50, max 100).',
        },
        status: {
          type: 'string',
          enum: ['notStarted', 'inProgress', 'completed'],
          description: 'Optional filter restricting results to tasks with this status.',
        },
      },
      required: ['list_id'],
    },
  },
  {
    name: 'get_task',
    description:
      'Get the full details of a single Microsoft To Do task by its list ID and task ID, including title, status, importance, full body/notes, due date, reminder time, created/completed times, and any checklist (sub-task) items. Use after list_tasks or find_due_tasks to read a to-do item in full.',
    inputSchema: {
      type: 'object' as const,
      properties: {
        list_id: {
          type: 'string',
          description: 'The ID of the Microsoft To Do task list containing the task.',
        },
        task_id: {
          type: 'string',
          description: 'The ID of the task to retrieve (from a list or find result).',
        },
      },
      required: ['list_id', 'task_id'],
    },
  },
  {
    name: 'find_due_tasks',
    description:
      'Find upcoming, incomplete Microsoft To Do tasks across all of the user\'s to-do lists that are due within the next N days. Returns a flat list of due reminders with their list name, title, due date, and importance. Use to answer "what tasks are due soon" or "what are my upcoming reminders".',
    inputSchema: {
      type: 'object' as const,
      properties: {
        days_ahead: {
          type: 'number',
          description: 'How many days ahead to look for due tasks (default 7).',
        },
      },
      required: [],
    },
  },
];

interface GraphDateTime {
  dateTime?: string;
  timeZone?: string;
}

interface GraphTask {
  id?: string;
  title?: string;
  status?: string;
  importance?: string;
  isReminderOn?: boolean;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  dueDateTime?: GraphDateTime;
  reminderDateTime?: GraphDateTime;
  completedDateTime?: GraphDateTime;
  body?: { contentType?: string; content?: string };
  checklistItems?: unknown;
}

interface GraphList {
  id?: string;
  displayName?: string;
  wellknownListName?: string;
  isOwner?: boolean;
}

function compactTask(t: GraphTask): Record<string, unknown> {
  return {
    id: t.id,
    title: t.title,
    status: t.status,
    importance: t.importance,
    due: t.dueDateTime?.dateTime,
    created: t.createdDateTime,
    completed: t.completedDateTime?.dateTime,
    preview: t.body?.content?.slice(0, 200),
  };
}

async function callTool(name: string, args: Record<string, unknown>): Promise<unknown> {
  const context = (args._context ?? {}) as MicrosoftTodoContext;
  delete args._context;

  switch (name) {
    case 'list_task_lists': {
      const params = new URLSearchParams({
        $select: 'id,displayName,wellknownListName,isOwner',
      });
      const result = await gFetch(context, `${API}/me/todo/lists?${params}`);
      const value = (result as { value?: unknown[] }).value;
      if (Array.isArray(value)) return value;
      return result;
    }
    case 'list_tasks': {
      const listId = args.list_id as string;
      const top = Math.min(100, Math.max(1, (args.top as number) ?? 50));
      const status = args.status as string | undefined;

      const params = new URLSearchParams({
        $top: String(top),
        $select:
          'id,title,status,importance,isReminderOn,createdDateTime,lastModifiedDateTime,dueDateTime,completedDateTime,body',
      });
      if (status) {
        params.set('$filter', `status eq '${status}'`);
      }

      const result = await gFetch(
        context,
        `${API}/me/todo/lists/${encodeURIComponent(listId)}/tasks?${params}`,
      );
      const value = (result as { value?: GraphTask[] }).value;
      if (Array.isArray(value)) return value.map(compactTask);
      return result;
    }
    case 'get_task': {
      const listId = args.list_id as string;
      const taskId = args.task_id as string;
      const params = new URLSearchParams({ $expand: 'checklistItems' });
      const result = await gFetch(
        context,
        `${API}/me/todo/lists/${encodeURIComponent(listId)}/tasks/${encodeURIComponent(taskId)}?${params}`,
      );
      const t = result as GraphTask;
      if (t && typeof t === 'object' && 'id' in t && !('error' in t)) {
        return {
          id: t.id,
          title: t.title,
          status: t.status,
          importance: t.importance,
          body: t.body?.content,
          dueDateTime: t.dueDateTime,
          reminderDateTime: t.reminderDateTime,
          createdDateTime: t.createdDateTime,
          completedDateTime: t.completedDateTime,
          checklistItems: t.checklistItems,
        };
      }
      return result;
    }
    case 'find_due_tasks': {
      const daysAhead = (args.days_ahead as number) ?? 7;
      const now = new Date();
      const windowEnd = new Date(now.getTime() + daysAhead * 24 * 60 * 60 * 1000);

      const listsResult = await gFetch(context, `${API}/me/todo/lists`);
      const lists = (listsResult as { value?: GraphList[] }).value;
      if (!Array.isArray(lists)) return listsResult;

      const due: Array<Record<string, unknown>> = [];
      const params = new URLSearchParams({ $filter: "status ne 'completed'" });

      for (const list of lists.slice(0, 20)) {
        if (!list.id) continue;
        const tasksResult = await gFetch(
          context,
          `${API}/me/todo/lists/${encodeURIComponent(list.id)}/tasks?${params}`,
        );
        const tasks = (tasksResult as { value?: GraphTask[] }).value;
        if (!Array.isArray(tasks)) continue;
        for (const t of tasks) {
          const dueStr = t.dueDateTime?.dateTime;
          if (!dueStr) continue;
          const dueDate = new Date(dueStr);
          if (dueDate >= now && dueDate <= windowEnd) {
            due.push({
              list: list.displayName,
              title: t.title,
              due: dueStr,
              importance: t.importance,
            });
          }
        }
      }
      return due;
    }
    default:
      throw new Error(`Unknown tool: ${name}`);
  }
}

export default { tools, callTool, meter: { credits: 1 }, provider: 'microsoft' } satisfies McpToolExport;
