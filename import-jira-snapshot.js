const fs = require("fs");
const path = require("path");

const sourcePath = process.argv[2];
const outputPath = process.argv[3] || path.resolve(process.cwd(), "jira-snapshot.js");

if (!sourcePath) {
  console.error("Usage: node scripts/import-jira-snapshot.js <csv-path> [output-path]");
  process.exit(1);
}

const knownAccounts = [
  { id: "acct-mi", name: "Motion Industries", mask: "MIPOWR" },
  { id: "acct-voltagrid", name: "Voltagrid", mask: "846164" },
  { id: "acct-meta", name: "Meta", mask: "273187" },
  { id: "acct-komatsu", name: "Komatsu", mask: "KM" },
  { id: "acct-uslubricants", name: "US Lubricants", mask: "USLUBE" },
  { id: "acct-chevron", name: "Chevron", mask: "LUB" },
  { id: "acct-kubota", name: "Kubota", mask: "KUBONA" }
];

function parseJiraDate(value) {
  if (!value) {
    return null;
  }

  const match = value.trim().match(/^([A-Za-z]{3})\/(\d{1,2})\/(\d{2}) (\d{1,2}):(\d{2}) (AM|PM)$/);
  if (!match) {
    return null;
  }

  const monthMap = {
    Jan: 0,
    Feb: 1,
    Mar: 2,
    Apr: 3,
    May: 4,
    Jun: 5,
    Jul: 6,
    Aug: 7,
    Sep: 8,
    Oct: 9,
    Nov: 10,
    Dec: 11
  };

  let hour = Number(match[4]);
  const minute = Number(match[5]);
  const meridiem = match[6];
  if (meridiem === "PM" && hour !== 12) {
    hour += 12;
  }
  if (meridiem === "AM" && hour === 12) {
    hour = 0;
  }

  return new Date(2000 + Number(match[3]), monthMap[match[1]], Number(match[2]), hour, minute);
}

function normalizeStatus(status) {
  const value = String(status || "").trim().toLowerCase();
  if (value.includes("progress")) {
    return "In progress";
  }
  if (value.includes("resolve") || value.includes("done") || value.includes("closed")) {
    return "Resolved";
  }
  return "Open";
}

function normalizeMask(value) {
  return String(value || "").toUpperCase().replace(/[^A-Z0-9]/g, "");
}

function rawTokens(value) {
  return String(value || "")
    .toUpperCase()
    .split(/[^A-Z0-9]+/)
    .filter(Boolean);
}

function mapAccountByMask(accountNumber) {
  const compact = normalizeMask(accountNumber);
  const tokens = rawTokens(accountNumber);

  return knownAccounts.find((account) => {
    const mask = account.mask.toUpperCase();
    if (mask.length <= 3) {
      return compact === mask || tokens.includes(mask);
    }
    return compact === mask || compact.startsWith(mask);
  }) || null;
}

function updateCounter(counter, key) {
  counter[key] = (counter[key] || 0) + 1;
}

function createBucket() {
  return {
    total: 0,
    open: 0,
    inProgress: 0,
    resolved: 0,
    recent: [],
    componentTypes: {}
  };
}

function addTicketToBucket(bucket, ticket) {
  bucket.total += 1;
  if (ticket.statusGroup === "Open") {
    bucket.open += 1;
  } else if (ticket.statusGroup === "In progress") {
    bucket.inProgress += 1;
  } else {
    bucket.resolved += 1;
  }
  bucket.recent.push(ticket);
  if (ticket.componentType) {
    updateCounter(bucket.componentTypes, ticket.componentType);
  }
}

const cutoff = new Date("2026-03-19T00:00:00");
const stream = fs.createReadStream(sourcePath, { encoding: "utf8" });

let headers = null;
let row = [];
let field = "";
let inQuotes = false;
let prevChar = "";
let processedRows = 0;

const snapshot = {
  generatedAt: new Date().toISOString(),
  sourceFile: sourcePath,
  windowDays: 30,
  processedRows: 0,
  importedRows: 0,
  overall: createBucket(),
  unmapped: createBucket(),
  byAccount: {},
  byManager: {},
  recent: []
};

function buildIndex(headersRow) {
  const findNth = (name, occurrence = 1) => {
    let seen = 0;
    for (let i = 0; i < headersRow.length; i += 1) {
      if (headersRow[i] === name) {
        seen += 1;
        if (seen === occurrence) {
          return i;
        }
      }
    }
    return -1;
  };

  return {
    summary: findNth("Summary"),
    issueKey: findNth("Issue key"),
    status: findNth("Status"),
    priority: findNth("Priority"),
    assignee: findNth("Assignee"),
    created: findNth("Created"),
    componentType: findNth("Components", 1),
    accountManager: findNth("Custom field (Account Manager)"),
    accountNumber: findNth("Custom field (Cust. Account Number)"),
    companyName: findNth("Custom field (Company Name)", 1),
    companyNameAlt: findNth("Custom field (Company Name)", 2)
  };
}

let indexMap = null;

function processRow(fields) {
  if (!headers) {
    headers = fields;
    indexMap = buildIndex(headers);
    return;
  }

  processedRows += 1;
  const createdAt = parseJiraDate(fields[indexMap.created]);
  if (!createdAt || createdAt < cutoff) {
    return;
  }

  const companyName =
    String(fields[indexMap.companyName] || "").trim() ||
    String(fields[indexMap.companyNameAlt] || "").trim();
  const accountNumber = String(fields[indexMap.accountNumber] || "").trim();
  const mappedAccount = mapAccountByMask(accountNumber);

  if (!mappedAccount) {
    return;
  }

  const ticket = {
    key: String(fields[indexMap.issueKey] || "").trim(),
    summary: String(fields[indexMap.summary] || "").trim(),
    status: String(fields[indexMap.status] || "").trim(),
    statusGroup: normalizeStatus(fields[indexMap.status]),
    priority: String(fields[indexMap.priority] || "").trim() || "Medium",
    assignee: String(fields[indexMap.assignee] || "").trim() || "Unassigned",
    created: createdAt.toISOString(),
    createdLabel: String(fields[indexMap.created] || "").trim(),
    componentType: String(fields[indexMap.componentType] || "").trim(),
    accountManager: String(fields[indexMap.accountManager] || "").trim(),
    companyName,
    accountNumber,
    matchedMask: mappedAccount.mask
  };

  snapshot.importedRows += 1;
  addTicketToBucket(snapshot.overall, ticket);
  snapshot.recent.push(ticket);

  if (!snapshot.byAccount[mappedAccount.id]) {
    snapshot.byAccount[mappedAccount.id] = createBucket();
    snapshot.byAccount[mappedAccount.id].accountName = mappedAccount.name;
    snapshot.byAccount[mappedAccount.id].mask = mappedAccount.mask;
  } else {
    snapshot.byAccount[mappedAccount.id].mask = mappedAccount.mask;
  }
  addTicketToBucket(snapshot.byAccount[mappedAccount.id], ticket);

  if (ticket.accountManager) {
    if (!snapshot.byManager[ticket.accountManager]) {
      snapshot.byManager[ticket.accountManager] = createBucket();
    }
    addTicketToBucket(snapshot.byManager[ticket.accountManager], ticket);
  }
}

function finalizeRow() {
  row.push(field);
  processRow(row);
  row = [];
  field = "";
}

stream.on("data", (chunk) => {
  for (let i = 0; i < chunk.length; i += 1) {
    const char = chunk[i];

    if (char === '"' && prevChar !== "\\") {
      if (inQuotes && chunk[i + 1] === '"') {
        field += '"';
        i += 1;
      } else {
        inQuotes = !inQuotes;
      }
    } else if (char === "," && !inQuotes) {
      row.push(field);
      field = "";
    } else if ((char === "\n" || char === "\r") && !inQuotes) {
      if (char === "\r" && chunk[i + 1] === "\n") {
        i += 1;
      }
      finalizeRow();
    } else {
      field += char;
    }

    prevChar = char;
  }
});

stream.on("end", () => {
  if (field.length || row.length) {
    finalizeRow();
  }

  snapshot.processedRows = processedRows;
  snapshot.recent = snapshot.recent
    .sort((a, b) => new Date(b.created) - new Date(a.created))
    .slice(0, 40);

  Object.values(snapshot.byAccount).forEach((bucket) => {
    bucket.recent = bucket.recent
      .sort((a, b) => new Date(b.created) - new Date(a.created))
      .slice(0, 12);
  });

  Object.values(snapshot.byManager).forEach((bucket) => {
    bucket.recent = bucket.recent
      .sort((a, b) => new Date(b.created) - new Date(a.created))
      .slice(0, 12);
  });

  snapshot.unmapped.recent = snapshot.unmapped.recent
    .sort((a, b) => new Date(b.created) - new Date(a.created))
    .slice(0, 12);

  const output = `export const jiraSnapshot = ${JSON.stringify(snapshot, null, 2)};\n`;
  fs.writeFileSync(outputPath, output, "utf8");
  console.log(`Wrote snapshot to ${outputPath}`);
  console.log(`Imported ${snapshot.importedRows} Jira rows from the last ${snapshot.windowDays} days.`);
});

stream.on("error", (error) => {
  console.error(error.message);
  process.exit(1);
});
