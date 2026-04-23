const http = require("http");
const fs = require("fs");
const path = require("path");
const https = require("https");
const crypto = require("crypto");

const projectRoot = path.resolve(__dirname, "..", "..");
const port = Number(process.env.PORT || 4173);
const host = process.env.HOST || "0.0.0.0";
const publicDir = path.join(projectRoot, "public");
const clientDir = path.join(projectRoot, "src", "client");
const dataDir = path.join(projectRoot, "data");
const dashboardStorePath = path.join(dataDir, "dashboard-state.json");
const clients = new Set();
const fileTypes = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".svg": "image/svg+xml",
  ".png": "image/png",
  ".jpg": "image/jpeg",
  ".jpeg": "image/jpeg"
};

const liveReloadSnippet = `
<script>
  (() => {
    const events = new EventSource('/__live_reload');
    events.addEventListener('reload', () => window.location.reload());
  })();
</script>
`;
const microsoftAuthStates = new Set();
let microsoftSession = {
  connected: false,
  email: null,
  displayName: null,
  connectedAt: null
};
const mockNetSuiteRevenue = {
  generatedAt: new Date().toISOString(),
  byMask: {
    MIPOWR: {
      ytdRevenue: 428500,
      projectedYearEnd: 1214000,
      priorYearYtd: 391200,
      outstandingInvoices: 18200,
      monthly: [
        { label: "Q1", value: 306000 },
        { label: "Q2", value: 122500 },
        { label: "Q3", value: 392000 },
        { label: "Q4", value: 393500 }
      ]
    },
    "846164": {
      ytdRevenue: 219300,
      projectedYearEnd: 640000,
      priorYearYtd: 203900,
      outstandingInvoices: 11850,
      monthly: [
        { label: "Q1", value: 158400 },
        { label: "Q2", value: 60900 },
        { label: "Q3", value: 203600 },
        { label: "Q4", value: 217100 }
      ]
    },
    "273187": {
      ytdRevenue: 186700,
      projectedYearEnd: 521000,
      priorYearYtd: 171400,
      outstandingInvoices: 9400,
      monthly: [
        { label: "Q1", value: 132800 },
        { label: "Q2", value: 53900 },
        { label: "Q3", value: 166500 },
        { label: "Q4", value: 167800 }
      ]
    },
    KM: {
      ytdRevenue: 148900,
      projectedYearEnd: 455000,
      priorYearYtd: 136200,
      outstandingInvoices: 7300,
      monthly: [
        { label: "Q1", value: 105400 },
        { label: "Q2", value: 43500 },
        { label: "Q3", value: 150900 },
        { label: "Q4", value: 155200 }
      ]
    },
    USLUBE: {
      ytdRevenue: 203400,
      projectedYearEnd: 592000,
      priorYearYtd: 188100,
      outstandingInvoices: 12600,
      monthly: [
        { label: "Q1", value: 145800 },
        { label: "Q2", value: 57600 },
        { label: "Q3", value: 191200 },
        { label: "Q4", value: 197400 }
      ]
    },
    OILANA: {
      ytdRevenue: 131600,
      projectedYearEnd: 384000,
      priorYearYtd: 120800,
      outstandingInvoices: 6100,
      monthly: [
        { label: "Q1", value: 94700 },
        { label: "Q2", value: 36900 },
        { label: "Q3", value: 121300 },
        { label: "Q4", value: 131100 }
      ]
    },
    JGLUBR: {
      ytdRevenue: 116200,
      projectedYearEnd: 338000,
      priorYearYtd: 108900,
      outstandingInvoices: 5400,
      monthly: [
        { label: "Q1", value: 82400 },
        { label: "Q2", value: 33800 },
        { label: "Q3", value: 109600 },
        { label: "Q4", value: 112200 }
      ]
    },
    LUB: {
      ytdRevenue: 174900,
      projectedYearEnd: 498000,
      priorYearYtd: 162100,
      outstandingInvoices: 8700,
      monthly: [
        { label: "Q1", value: 123300 },
        { label: "Q2", value: 51600 },
        { label: "Q3", value: 163900 },
        { label: "Q4", value: 159200 }
      ]
    },
    KUBONA: {
      ytdRevenue: 194500,
      projectedYearEnd: 558000,
      priorYearYtd: 181000,
      outstandingInvoices: 9800,
      monthly: [
        { label: "Q1", value: 139600 },
        { label: "Q2", value: 54900 },
        { label: "Q3", value: 184400 },
        { label: "Q4", value: 179100 }
      ]
    }
  }
};
const mockGenesysSignals = {
  generatedAt: new Date().toISOString(),
  orgName: "POLARIS CX Cloud",
  byRegion: {
    NA: {
      hotOpportunities: 5,
      openCallbacks: 12,
      abandonedInteractions: 4,
      queueAlerts: ["Pricing follow-up requests rising in Midwest", "Two onboarding handoff calls waiting on AM response"]
    },
    EMEA: {
      hotOpportunities: 2,
      openCallbacks: 5,
      abandonedInteractions: 2,
      queueAlerts: ["Lubrication monitoring demo requests need qualification"]
    },
    APAC: {
      hotOpportunities: 3,
      openCallbacks: 7,
      abandonedInteractions: 1,
      queueAlerts: ["Training-heavy opportunities trending upward"]
    }
  }
};

function ensureDataDir() {
  if (!fs.existsSync(dataDir)) {
    fs.mkdirSync(dataDir, { recursive: true });
  }
}

function readDashboardStore() {
  try {
    if (!fs.existsSync(dashboardStorePath)) {
      return null;
    }

    const raw = fs.readFileSync(dashboardStorePath, "utf8");
    return raw ? JSON.parse(raw) : null;
  } catch {
    return null;
  }
}

function writeDashboardStore(payload) {
  ensureDataDir();
  const stored = {
    version: 1,
    updatedAt: new Date().toISOString(),
    data: payload
  };
  fs.writeFileSync(dashboardStorePath, JSON.stringify(stored, null, 2), "utf8");
  return stored;
}

function readJsonBody(request) {
  return new Promise((resolve, reject) => {
    let body = "";
    request.on("data", (chunk) => {
      body += chunk;
      if (body.length > 1024 * 1024) {
        reject(new Error("Payload too large"));
      }
    });
    request.on("end", () => {
      try {
        resolve(body ? JSON.parse(body) : {});
      } catch (error) {
        reject(error);
      }
    });
    request.on("error", reject);
  });
}

function toAdf(text) {
  const content = String(text || "")
    .split(/\r?\n/)
    .filter((line, index, all) => line.trim() || index < all.length - 1)
    .map((line) => ({
      type: "paragraph",
      content: line
        ? [
            {
              type: "text",
              text: line
            }
          ]
        : []
    }));

  return {
    type: "doc",
    version: 1,
    content: content.length ? content : [{ type: "paragraph", content: [] }]
  };
}

function writeJson(response, statusCode, payload) {
  response.writeHead(statusCode, {
    "Content-Type": "application/json; charset=utf-8",
    "Cache-Control": "no-store"
  });
  response.end(JSON.stringify(payload));
}

function writeHtml(response, statusCode, html) {
  response.writeHead(statusCode, {
    "Content-Type": "text/html; charset=utf-8",
    "Cache-Control": "no-store"
  });
  response.end(html);
}

function postForm(url, body) {
  return new Promise((resolve, reject) => {
    const payload = new URLSearchParams(body).toString();
    const req = https.request(
      url,
      {
        method: "POST",
        headers: {
          "Content-Type": "application/x-www-form-urlencoded",
          "Content-Length": Buffer.byteLength(payload)
        }
      },
      (res) => {
        let responseBody = "";
        res.on("data", (chunk) => {
          responseBody += chunk;
        });
        res.on("end", () => {
          try {
            resolve({
              statusCode: res.statusCode || 500,
              data: responseBody ? JSON.parse(responseBody) : {}
            });
          } catch (error) {
            reject(error);
          }
        });
      }
    );

    req.on("error", reject);
    req.write(payload);
    req.end();
  });
}

function getJson(url, accessToken) {
  return new Promise((resolve, reject) => {
    const req = https.request(
      url,
      {
        method: "GET",
        headers: {
          Accept: "application/json",
          Authorization: `Bearer ${accessToken}`
        }
      },
      (res) => {
        let responseBody = "";
        res.on("data", (chunk) => {
          responseBody += chunk;
        });
        res.on("end", () => {
          try {
            resolve({
              statusCode: res.statusCode || 500,
              data: responseBody ? JSON.parse(responseBody) : {}
            });
          } catch (error) {
            reject(error);
          }
        });
      }
    );

    req.on("error", reject);
    req.end();
  });
}

function createJiraIssue(payload) {
  return new Promise((resolve, reject) => {
    const baseUrl = process.env.JIRA_BASE_URL;
    const email = process.env.JIRA_EMAIL;
    const token = process.env.JIRA_API_TOKEN;
    const projectKey = process.env.JIRA_PROJECT_KEY;
    const issueType = process.env.JIRA_ISSUE_TYPE || "Task";

    if (!baseUrl || !email || !token || !projectKey) {
      resolve({
        configured: false
      });
      return;
    }

    const issuePayload = {
      fields: {
        project: { key: projectKey },
        issuetype: { name: issueType },
        summary: payload.summary,
        description: toAdf(payload.description),
        priority: payload.priority ? { name: payload.priority } : undefined
      }
    };

    const requestUrl = new URL("/rest/api/3/issue", baseUrl);
    const auth = Buffer.from(`${email}:${token}`).toString("base64");
    const req = https.request(
      requestUrl,
      {
        method: "POST",
        headers: {
          Accept: "application/json",
          "Content-Type": "application/json",
          Authorization: `Basic ${auth}`
        }
      },
      (jiraResponse) => {
        let responseBody = "";
        jiraResponse.on("data", (chunk) => {
          responseBody += chunk;
        });
        jiraResponse.on("end", () => {
          const parsed = responseBody ? JSON.parse(responseBody) : {};
          if (jiraResponse.statusCode && jiraResponse.statusCode >= 200 && jiraResponse.statusCode < 300) {
            resolve({
              configured: true,
              key: parsed.key,
              id: parsed.id,
              self: parsed.self,
              browseUrl: new URL(`/browse/${parsed.key}`, baseUrl).toString()
            });
            return;
          }

          resolve({
            configured: true,
            error: parsed.errorMessages?.join(" ") || parsed.errors ? JSON.stringify(parsed.errors) : "Jira create failed"
          });
        });
      }
    );

    req.on("error", reject);
    req.write(JSON.stringify(issuePayload));
    req.end();
  });
}

function safePath(urlPath) {
  const requestedPath = urlPath === "/" ? "/index.html" : urlPath;
  const staticMap = {
    "/index.html": path.join(publicDir, "index.html"),
    "/styles.css": path.join(clientDir, "styles.css"),
    "/app.js": path.join(clientDir, "app.js"),
    "/equipment-list-template.csv": path.join(publicDir, "equipment-list-template.csv"),
    "/data/complaints-snapshot.js": path.join(clientDir, "data", "complaints-snapshot.js"),
    "/data/jira-snapshot.js": path.join(clientDir, "data", "jira-snapshot.js")
  };

  const mappedPath = staticMap[requestedPath];
  if (mappedPath) {
    return mappedPath;
  }

  const filePath = path.normalize(path.join(publicDir, requestedPath.replace(/^\//, "")));
  if (!filePath.startsWith(publicDir)) {
    return null;
  }
  return filePath;
}

function serveFile(filePath, response) {
  fs.readFile(filePath, (error, content) => {
    if (error) {
      response.writeHead(404, { "Content-Type": "text/plain; charset=utf-8" });
      response.end("Not found");
      return;
    }

    const extension = path.extname(filePath).toLowerCase();
    const type = fileTypes[extension] || "application/octet-stream";
    const body =
      extension === ".html"
        ? content.toString("utf8").replace("</body>", `${liveReloadSnippet}</body>`)
        : content;

    response.writeHead(200, { "Content-Type": type, "Cache-Control": "no-store" });
    response.end(body);
  });
}

const server = http.createServer((request, response) => {
  if (!request.url) {
    response.writeHead(400);
    response.end();
    return;
  }

  if (request.url === "/__live_reload") {
    response.writeHead(200, {
      "Content-Type": "text/event-stream",
      "Cache-Control": "no-cache",
      Connection: "keep-alive"
    });
    response.write("\n");
    clients.add(response);
    request.on("close", () => clients.delete(response));
    return;
  }

  if (request.method === "GET" && request.url === "/api/jira/config") {
    writeJson(response, 200, {
      configured: Boolean(
        process.env.JIRA_BASE_URL &&
          process.env.JIRA_EMAIL &&
          process.env.JIRA_API_TOKEN &&
          process.env.JIRA_PROJECT_KEY
      ),
      baseUrl: process.env.JIRA_BASE_URL || null,
      projectKey: process.env.JIRA_PROJECT_KEY || null,
      issueType: process.env.JIRA_ISSUE_TYPE || "Task"
    });
    return;
  }

  if (request.method === "GET" && request.url === "/api/dashboard-state") {
    const stored = readDashboardStore();
    writeJson(response, 200, {
      ok: true,
      storeAvailable: Boolean(stored),
      version: stored?.version || 1,
      updatedAt: stored?.updatedAt || null,
      data: stored?.data || null
    });
    return;
  }

  if (request.method === "PUT" && request.url === "/api/dashboard-state") {
    readJsonBody(request)
      .then((payload) => {
        if (!payload || typeof payload !== "object" || Array.isArray(payload)) {
          writeJson(response, 400, {
            ok: false,
            message: "Dashboard state payload must be a JSON object."
          });
          return;
        }

        const stored = writeDashboardStore(payload);
        writeJson(response, 200, {
          ok: true,
          version: stored.version,
          updatedAt: stored.updatedAt
        });
      })
      .catch((error) => {
        writeJson(response, 400, {
          ok: false,
          message: error.message || "Could not save dashboard state."
        });
      });
    return;
  }

  if (request.method === "GET" && request.url === "/api/microsoft/config") {
    const scopes = (process.env.MS_GRAPH_SCOPES || "User.Read Calendars.ReadWrite OnlineMeetings.ReadWrite Files.Read")
      .split(/[,\s]+/)
      .filter(Boolean);
    writeJson(response, 200, {
      configured: Boolean(
        process.env.MS_TENANT_ID &&
          process.env.MS_CLIENT_ID &&
          process.env.MS_CLIENT_SECRET &&
          process.env.MS_REDIRECT_URI
      ),
      tenantId: process.env.MS_TENANT_ID || null,
      clientId: process.env.MS_CLIENT_ID || null,
      redirectUri: process.env.MS_REDIRECT_URI || null,
      scopes,
      connected: microsoftSession.connected,
      email: microsoftSession.email,
      displayName: microsoftSession.displayName,
      connectedAt: microsoftSession.connectedAt,
      teamsUrl: "https://teams.microsoft.com",
      officeUrl: "https://www.office.com",
      authUrl: "/auth/microsoft/start"
    });
    return;
  }

  if (request.method === "GET" && request.url === "/api/netsuite/config") {
    writeJson(response, 200, {
      configured: true,
      mode: "mock",
      generatedAt: mockNetSuiteRevenue.generatedAt,
      masks: Object.keys(mockNetSuiteRevenue.byMask)
    });
    return;
  }

  if (request.method === "GET" && request.url === "/api/netsuite/revenue") {
    const overall = Object.values(mockNetSuiteRevenue.byMask).reduce(
      (totals, item) => {
        totals.ytdRevenue += item.ytdRevenue || 0;
        totals.projectedYearEnd += item.projectedYearEnd || 0;
        totals.priorYearYtd += item.priorYearYtd || 0;
        totals.outstandingInvoices += item.outstandingInvoices || 0;
        return totals;
      },
      { ytdRevenue: 0, projectedYearEnd: 0, priorYearYtd: 0, outstandingInvoices: 0 }
    );

    writeJson(response, 200, {
      configured: true,
      mode: "mock",
      generatedAt: mockNetSuiteRevenue.generatedAt,
      overall,
      byMask: mockNetSuiteRevenue.byMask
    });
    return;
  }

  if (request.method === "GET" && request.url === "/api/genesys/config") {
    const region = process.env.GENESYS_CLOUD_REGION || "mypurecloud.com";
    const scopes = (process.env.GENESYS_SCOPES || "analytics:readonly routing:readonly conversations:readonly users:readonly")
      .split(/[,\s]+/)
      .filter(Boolean);
    writeJson(response, 200, {
      configured: Boolean(process.env.GENESYS_CLIENT_ID && process.env.GENESYS_CLIENT_SECRET),
      mode: process.env.GENESYS_CLIENT_ID && process.env.GENESYS_CLIENT_SECRET ? "oauth-scaffold" : "mock",
      region,
      orgName: process.env.GENESYS_ORG_NAME || mockGenesysSignals.orgName,
      generatedAt: mockGenesysSignals.generatedAt,
      portalUrl: process.env.GENESYS_PORTAL_URL || "",
      developerUrl: "https://developer.genesys.cloud",
      scopes,
      byRegion: mockGenesysSignals.byRegion
    });
    return;
  }

  if (request.method === "GET" && request.url.startsWith("/auth/microsoft/start")) {
    if (
      !process.env.MS_TENANT_ID ||
      !process.env.MS_CLIENT_ID ||
      !process.env.MS_CLIENT_SECRET ||
      !process.env.MS_REDIRECT_URI
    ) {
      writeHtml(
        response,
        501,
        `<html><body><h1>Microsoft not configured</h1><p>Set the Microsoft Graph environment variables first.</p></body></html>`
      );
      return;
    }

    const requestUrl = new URL(request.url, `http://127.0.0.1:${port}`);
    const loginHint = requestUrl.searchParams.get("login_hint") || "";
    const state = crypto.randomBytes(16).toString("hex");
    microsoftAuthStates.add(state);
    const scopes = (process.env.MS_GRAPH_SCOPES || "User.Read Calendars.ReadWrite OnlineMeetings.ReadWrite Files.Read offline_access openid profile")
      .split(/[,\s]+/)
      .filter(Boolean)
      .join(" ");
    const authorizeUrl = new URL(`https://login.microsoftonline.com/${process.env.MS_TENANT_ID}/oauth2/v2.0/authorize`);
    authorizeUrl.searchParams.set("client_id", process.env.MS_CLIENT_ID);
    authorizeUrl.searchParams.set("response_type", "code");
    authorizeUrl.searchParams.set("redirect_uri", process.env.MS_REDIRECT_URI);
    authorizeUrl.searchParams.set("response_mode", "query");
    authorizeUrl.searchParams.set("scope", scopes);
    authorizeUrl.searchParams.set("state", state);
    authorizeUrl.searchParams.set("prompt", "select_account");
    if (loginHint) {
      authorizeUrl.searchParams.set("login_hint", loginHint);
    }

    response.writeHead(302, { Location: authorizeUrl.toString() });
    response.end();
    return;
  }

  if (request.method === "GET" && request.url.startsWith("/auth/microsoft/callback")) {
    const callbackUrl = new URL(request.url, `http://127.0.0.1:${port}`);
    const code = callbackUrl.searchParams.get("code");
    const state = callbackUrl.searchParams.get("state");
    const error = callbackUrl.searchParams.get("error");

    if (error) {
      writeHtml(response, 400, `<html><body><h1>Microsoft sign-in failed</h1><p>${error}</p></body></html>`);
      return;
    }

    if (!code || !state || !microsoftAuthStates.has(state)) {
      writeHtml(response, 400, `<html><body><h1>Invalid Microsoft callback</h1><p>The sign-in state could not be verified.</p></body></html>`);
      return;
    }

    microsoftAuthStates.delete(state);

    const tokenUrl = new URL(`https://login.microsoftonline.com/${process.env.MS_TENANT_ID}/oauth2/v2.0/token`);
    postForm(tokenUrl, {
      client_id: process.env.MS_CLIENT_ID,
      client_secret: process.env.MS_CLIENT_SECRET,
      grant_type: "authorization_code",
      code,
      redirect_uri: process.env.MS_REDIRECT_URI,
      scope: (process.env.MS_GRAPH_SCOPES || "User.Read Calendars.ReadWrite OnlineMeetings.ReadWrite Files.Read offline_access openid profile")
        .split(/[,\s]+/)
        .filter(Boolean)
        .join(" ")
    })
      .then(async (tokenResponse) => {
        if (tokenResponse.statusCode < 200 || tokenResponse.statusCode >= 300 || !tokenResponse.data.access_token) {
          throw new Error(tokenResponse.data.error_description || "Token exchange failed");
        }

        const meResponse = await getJson(new URL("https://graph.microsoft.com/v1.0/me"), tokenResponse.data.access_token);
        const profile = meResponse.data || {};
        microsoftSession = {
          connected: true,
          email: profile.mail || profile.userPrincipalName || null,
          displayName: profile.displayName || null,
          connectedAt: new Date().toISOString()
        };

        writeHtml(
          response,
          200,
          `<html><body><h1>Microsoft connected</h1><p>${microsoftSession.displayName || microsoftSession.email || "Account connected"} is now available in the sandbox.</p><script>setTimeout(() => window.close(), 1200);</script></body></html>`
        );
      })
      .catch((authError) => {
        writeHtml(response, 500, `<html><body><h1>Microsoft sign-in error</h1><p>${authError.message}</p></body></html>`);
      });
    return;
  }

  if (request.method === "POST" && request.url === "/api/jira/create") {
    readJsonBody(request)
      .then((payload) => createJiraIssue(payload))
      .then((result) => {
        if (!result.configured) {
          writeJson(response, 501, {
            ok: false,
            code: "jira_not_configured",
            message: "Jira credentials are not configured in the local preview server."
          });
          return;
        }

        if (result.error) {
          writeJson(response, 502, {
            ok: false,
            code: "jira_create_failed",
            message: result.error
          });
          return;
        }

        writeJson(response, 201, {
          ok: true,
          key: result.key,
          id: result.id,
          self: result.self,
          browseUrl: result.browseUrl
        });
      })
      .catch((error) => {
        writeJson(response, 500, {
          ok: false,
          code: "server_error",
          message: error.message
        });
      });
    return;
  }

  const target = safePath(request.url.split("?")[0]);
  if (!target) {
    response.writeHead(403, { "Content-Type": "text/plain; charset=utf-8" });
    response.end("Forbidden");
    return;
  }

  serveFile(target, response);
});

let reloadTimer;
fs.watch(projectRoot, { recursive: true }, (_, filename) => {
  if (!filename || filename.startsWith(".git")) {
    return;
  }

  clearTimeout(reloadTimer);
  reloadTimer = setTimeout(() => {
    for (const client of clients) {
      client.write("event: reload\ndata: now\n\n");
    }
  }, 80);
});

server.listen(port, host, () => {
  console.log(`Sandbox preview running at http://${host === "0.0.0.0" ? "127.0.0.1" : host}:${port}`);
});
