const { app } = require("@azure/functions");
const { graphFetch } = require("../graphClient");

const SITE_HOSTNAME = "wesconc.sharepoint.com";
const SITE_PATH = "/sites/Wesco";
const ALLOWED_ROOTS = ["We App/Projects", "We App/Internal Staff"];

let siteIdPromise = null;

function cleanPath(value) {
  const path = String(value || "")
    .replace(/\\/g, "/")
    .replace(/^\/+|\/+$/g, "")
    .replace(/\/+/g, "/");
  if (!path || path.split("/").some((part) => part === "." || part === "..")) {
    throw Object.assign(new Error("A valid Wesco folder path is required."), {
      status: 400,
    });
  }
  const lower = path.toLowerCase();
  const allowed = ALLOWED_ROOTS.some(
    (root) =>
      lower === root.toLowerCase() ||
      lower.startsWith(`${root.toLowerCase()}/`)
  );
  if (!allowed) {
    throw Object.assign(new Error("That folder is outside the Wesco Hub."), {
      status: 403,
    });
  }
  return path;
}

function cleanName(value, label) {
  const name = String(value || "").trim();
  if (
    !name ||
    name.length > 180 ||
    /[\/\\]/.test(name) ||
    name === "." ||
    name === ".."
  ) {
    throw Object.assign(new Error(`A valid ${label} is required.`), {
      status: 400,
    });
  }
  return name;
}

function encodeGraphPath(path) {
  return path.split("/").map(encodeURIComponent).join("/");
}

async function getSiteId() {
  if (!siteIdPromise) {
    siteIdPromise = graphFetch(
      `https://graph.microsoft.com/v1.0/sites/${SITE_HOSTNAME}:${SITE_PATH}`
    ).then((site) => site.id);
  }
  return siteIdPromise;
}

async function listChildren(path) {
  const siteId = await getSiteId();
  let url =
    `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/root:/` +
    `${encodeGraphPath(path)}:/children` +
    "?$select=id,name,webUrl,folder,file,lastModifiedDateTime&$top=400";
  const items = [];
  while (url) {
    const data = await graphFetch(url);
    items.push(...(data.value || []));
    url = data["@odata.nextLink"] || null;
  }
  return items.sort((a, b) => {
    if (Boolean(a.folder) !== Boolean(b.folder)) return a.folder ? -1 : 1;
    return a.name.localeCompare(b.name);
  });
}

app.http("hubFiles", {
  methods: ["GET"],
  authLevel: "anonymous",
  route: "hub-files",
  handler: async (request, context) => {
    try {
      const url = new URL(request.url);
      const path = cleanPath(url.searchParams.get("path"));
      return {
        headers: { "Cache-Control": "no-store" },
        jsonBody: { items: await listChildren(path) },
      };
    } catch (err) {
      context.error(err);
      return {
        status: err.status || 500,
        jsonBody: { error: err.status ? err.message : "Wesco files are temporarily unavailable." },
      };
    }
  },
});

app.http("hubFilesCreateFolder", {
  methods: ["POST"],
  authLevel: "anonymous",
  route: "hub-files/folder",
  handler: async (request, context) => {
    try {
      const body = await request.json();
      const path = cleanPath(body.path);
      const name = cleanName(body.name, "folder name");
      const siteId = await getSiteId();
      const created = await graphFetch(
        `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/root:/` +
          `${encodeGraphPath(path)}:/children`,
        {
          method: "POST",
          body: JSON.stringify({
            name,
            folder: {},
            "@microsoft.graph.conflictBehavior": "fail",
          }),
        }
      );
      return { status: 201, jsonBody: { item: created } };
    } catch (err) {
      context.error(err);
      return {
        status: err.status || 500,
        jsonBody: { error: err.status ? err.message : "The folder could not be created." },
      };
    }
  },
});

app.http("hubFilesUploadSession", {
  methods: ["POST"],
  authLevel: "anonymous",
  route: "hub-files/upload-session",
  handler: async (request, context) => {
    try {
      const body = await request.json();
      const path = cleanPath(body.path);
      const fileName = cleanName(body.fileName, "file name");
      const siteId = await getSiteId();
      const target = encodeGraphPath(`${path}/${fileName}`);
      const session = await graphFetch(
        `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/root:/` +
          `${target}:/createUploadSession`,
        {
          method: "POST",
          body: JSON.stringify({
            item: {
              "@microsoft.graph.conflictBehavior": "rename",
              name: fileName,
            },
          }),
        }
      );
      return { jsonBody: { uploadUrl: session.uploadUrl } };
    } catch (err) {
      context.error(err);
      return {
        status: err.status || 500,
        jsonBody: { error: err.status ? err.message : "The upload could not be started." },
      };
    }
  },
});
