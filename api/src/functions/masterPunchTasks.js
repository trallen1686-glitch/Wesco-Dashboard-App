const { app } = require("@azure/functions");
const { graphFetch } = require("../graphClient");

const DEFAULT_OWNER = "tallen@wesconc.com";
const DEFAULT_LIST_ID =
  "AQMkADYzOTRjM2FiLTNmOGMtNDE4ZS1hOTJiLWNhN2U5ZGEwZTg5YgAuAAADug1uys4PrEuQXED6qOMjzQEApFTVRWPTN0OAHHwFRbnXIAABDWwlRwAAAA==";

app.http("masterPunchTasks", {
  methods: ["GET"],
  authLevel: "anonymous",
  route: "master-punch-tasks",
  handler: async (_request, context) => {
    try {
      const owner = process.env.TODO_OWNER_USER || DEFAULT_OWNER;
      const listId = process.env.TODO_LIST_ID || DEFAULT_LIST_ID;
      const fields = "id,title,status,importance,dueDateTime,lastModifiedDateTime";
      const path =
        `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(owner)}` +
        `/todo/lists/${encodeURIComponent(listId)}/tasks` +
        `?$top=100&$select=${fields}`;
      const response = await graphFetch(path);
      const tasks = (response.value || []).map((task) => ({
        id: task.id,
        title: task.title || "Untitled task",
        status: task.status || "notStarted",
        importance: task.importance || "normal",
        dueDateTime: task.dueDateTime || null,
        lastModifiedDateTime: task.lastModifiedDateTime || null,
      }));

      return {
        headers: {
          "Cache-Control": "public, max-age=60, stale-while-revalidate=300",
          "Content-Type": "application/json; charset=utf-8",
        },
        jsonBody: { tasks, refreshedAt: new Date().toISOString() },
      };
    } catch (err) {
      context.error(err);
      return {
        status: err.status || 500,
        headers: { "Cache-Control": "no-store" },
        jsonBody: { error: "The Master Punch List could not be loaded." },
      };
    }
  },
});
