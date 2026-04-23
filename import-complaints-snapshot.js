const fs = require("fs");
const path = require("path");

const sourcePaths = process.argv.slice(2).filter((value) => value && value !== "--");
const outputPath = path.resolve(process.cwd(), "complaints-snapshot.js");

if (sourcePaths.length === 0) {
  console.error("Usage: node scripts/import-complaints-snapshot.js <csv-path> [more-csv-paths...]");
  process.exit(1);
}

const knownAccounts = [
  { id: "acct-mi", name: "Motion Industries", mask: "MIPOWR" },
  { id: "acct-voltagrid", name: "Voltagrid", mask: "846164" },
  { id: "acct-meta", name: "Meta", mask: "273187" },
  { id: "acct-komatsu", name: "Komatsu", mask: "KM" },
  { id: "acct-uslubricants", name: "US Lubricants", mask: "USLUBE" },
  { id: "acct-oilanalyzers", name: "Oil Analyzers", mask: "OILANA" },
  { id: "acct-jglubricants", name: "JG Lubricants", mask: "JGLUBR" },
  { id: "acct-chevron", name: "Chevron", mask: "LUB" },
  { id: "acct-kubota", name: "Kubota", mask: "KUBONA" }
];

const cutoff = new Date("2026-03-19T00:00:00");

function parseCsv(text) {
  const rows = [];
  let row = [];
  let field = "";
  let inQuotes = false;

  for (let index = 0; index < text.length; index += 1) {
    const char = text[index];

    if (char === '"') {
      if (inQuotes && text[index + 1] === '"') {
        field += '"';
        index += 1;
      } else {
        inQuotes = !inQuotes;
      }
    } else if (char === "," && !inQuotes) {
      row.push(field);
      field = "";
    } else if ((char === "\n" || char === "\r") && !inQuotes) {
      if (char === "\r" && text[index + 1] === "\n") {
        index += 1;
      }
      row.push(field);
      rows.push(row);
      row = [];
      field = "";
    } else {
      field += char;
    }
  }

  if (field.length || row.length) {
    row.push(field);
    rows.push(row);
  }

  return rows;
}

function toObjects(rows) {
  if (!rows.length) {
    return [];
  }

  const [headers, ...dataRows] = rows;
  return dataRows
    .filter((row) => row.some((value) => String(value || "").trim()))
    .map((row) =>
      headers.reduce((record, header, index) => {
        record[header] = row[index] || "";
        return record;
      }, {})
    );
}

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

function normalizeMask(value) {
  return String(value || "").toUpperCase().replace(/[^A-Z0-9]/g, "");
}

function tokenize(value) {
  return String(value || "")
    .toUpperCase()
    .split(/[^A-Z0-9]+/)
    .filter(Boolean);
}

function mapAccountFromSummary(summary) {
  const normalized = normalizeMask(summary);
  const tokens = tokenize(summary);

  for (const account of knownAccounts) {
    const normalizedMask = normalizeMask(account.mask);
    const normalizedName = normalizeMask(account.name);

    if (normalizedName && normalized.includes(normalizedName)) {
      return account;
    }

    if (normalizedMask.length <= 3) {
      if (tokens.some((token) => token === normalizedMask || token.startsWith(normalizedMask))) {
        return account;
      }
      continue;
    }

    if (normalized.includes(normalizedMask) || tokens.includes(normalizedMask)) {
      return account;
    }
  }

  return null;
}

function normalizeStatus(status) {
  const value = String(status || "").trim();
  if (!value) {
    return "Open";
  }
  return value;
}

function updateCounter(counter, key) {
  const label = String(key || "").trim() || "Unknown";
  counter[label] = (counter[label] || 0) + 1;
}

function createBucket() {
  return {
    total: 0,
    statuses: {},
    issueTypes: {},
    items: []
  };
}

function addItemToBucket(bucket, item) {
  bucket.total += 1;
  updateCounter(bucket.statuses, item.status);
  updateCounter(bucket.issueTypes, item.issueType);
  bucket.items.push(item);
}

const snapshot = {
  generatedAt: new Date().toISOString(),
  sourceFiles: sourcePaths,
  windowDays: 30,
  processedRows: 0,
  importedRows: 0,
  mappedRows: 0,
  unmappedRows: 0,
  overall: createBucket(),
  byAccount: {},
  unmatched: createBucket()
};

const seenKeys = new Set();

for (const sourcePath of sourcePaths) {
  const fileText = fs.readFileSync(sourcePath, "utf8");
  const rows = toObjects(parseCsv(fileText));

  rows.forEach((row) => {
    snapshot.processedRows += 1;

    const issueKey = String(row["Issue key"] || "").trim();
    if (!issueKey || seenKeys.has(issueKey)) {
      return;
    }

    const createdAt = parseJiraDate(row.Created);
    if (!createdAt || createdAt < cutoff) {
      return;
    }

    seenKeys.add(issueKey);
    snapshot.importedRows += 1;

    const item = {
      key: issueKey,
      issueId: String(row["Issue id"] || "").trim(),
      issueType: String(row["Issue Type"] || "").trim() || "Complaint",
      summary: String(row.Summary || "").trim(),
      status: normalizeStatus(row.Status),
      created: createdAt.toISOString(),
      createdLabel: String(row.Created || "").trim(),
      updatedLabel: String(row.Updated || "").trim(),
      resolvedLabel: String(row.Resolved || "").trim()
    };

    addItemToBucket(snapshot.overall, item);

    const mappedAccount = mapAccountFromSummary(item.summary);
    if (!mappedAccount) {
      snapshot.unmappedRows += 1;
      addItemToBucket(snapshot.unmatched, item);
      return;
    }

    snapshot.mappedRows += 1;
    item.matchedMask = mappedAccount.mask;
    item.accountId = mappedAccount.id;
    item.accountName = mappedAccount.name;

    if (!snapshot.byAccount[mappedAccount.id]) {
      snapshot.byAccount[mappedAccount.id] = createBucket();
      snapshot.byAccount[mappedAccount.id].mask = mappedAccount.mask;
      snapshot.byAccount[mappedAccount.id].accountName = mappedAccount.name;
    }

    addItemToBucket(snapshot.byAccount[mappedAccount.id], item);
  });
}

snapshot.overall.items = snapshot.overall.items
  .sort((left, right) => new Date(right.created) - new Date(left.created))
  .slice(0, 80);

snapshot.unmatched.items = snapshot.unmatched.items
  .sort((left, right) => new Date(right.created) - new Date(left.created))
  .slice(0, 40);

Object.values(snapshot.byAccount).forEach((bucket) => {
  bucket.items = bucket.items.sort((left, right) => new Date(right.created) - new Date(left.created)).slice(0, 20);
});

fs.writeFileSync(outputPath, `export const complaintsSnapshot = ${JSON.stringify(snapshot, null, 2)};\n`, "utf8");

console.log(`Wrote complaints snapshot to ${outputPath}`);
console.log(`Imported ${snapshot.importedRows} unique complaint rows from the last ${snapshot.windowDays} days.`);
console.log(`Mapped ${snapshot.mappedRows} complaints to known account masks.`);
