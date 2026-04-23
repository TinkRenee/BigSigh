import { complaintsSnapshot } from "/data/complaints-snapshot.js";

const PANEL_LAYOUT_STORAGE_KEY = "am-sandbox-panel-layouts-v1";
const DEFAULT_PANEL_ORDER = [
  "onboarding-workspace",
  "calendar",
  "workspace-tools",
  "revenue",
  "shared-workspace",
  "pipelines",
  "summary-table",
  "manager-workspace",
  "profile",
  "microsoft",
  "security"
];

function loadSavedPanelLayouts() {
  try {
    const raw = window.localStorage.getItem(PANEL_LAYOUT_STORAGE_KEY);
    return raw ? JSON.parse(raw) : {};
  } catch {
    return {};
  }
}

function getImportedComplaintBucket(accountId) {
  return complaintsSnapshot.byAccount?.[accountId] || null;
}

function buildAiComplaintDescription(item) {
  const summary = String(item?.summary || "").trim();
  if (!summary) {
    return "Imported complaint tied to this account.";
  }

  return `AI summary: ${summary.replace(/\s+/g, " ").trim()}.`;
}

function createComplaintData(accountId, mask, manager) {
  const importedBucket = getImportedComplaintBucket(accountId);
  if (!importedBucket) {
    return {
      total: 0,
      entries: [{ label: `${maskOrFallback(mask)} / Recorded complaints`, count: 0 }],
      records: []
    };
  }

  return {
    total: importedBucket.total || 0,
    entries: Object.entries(importedBucket.issueTypes || {})
      .sort((left, right) => right[1] - left[1])
      .map(([label, count]) => ({
        label,
        count
      })),
    records: (importedBucket.items || []).map((item) => ({
      id: `${accountId}-${item.key}`,
      title: item.key || item.issueType || "Complaint",
      summary: item.summary || "",
      issueType: item.issueType || "Complaint",
      status: item.status || "Open",
      owner: manager,
      date: item.createdLabel || "",
      detail: [item.issueType, buildAiComplaintDescription(item)].filter(Boolean).join(" - "),
      key: item.key || "",
      externalUrl: getJiraTicketUrl(item.key, item.externalUrl)
    }))
  };
}

const ONBOARDING_SECTION_TITLE = "Onboarding Space";
const ACCOUNT_ONBOARDING_TITLE = "Account onboarding";

function getTodayIsoDate() {
  return new Date().toISOString().slice(0, 10);
}

function addDaysToIsoDate(value, days) {
  const date = new Date(`${value || getTodayIsoDate()}T00:00:00`);
  date.setDate(date.getDate() + days);
  return date.toISOString().slice(0, 10);
}

function formatTimelineDate(value) {
  if (!value) {
    return "Not set";
  }

  const date = new Date(`${value}T00:00:00`);
  return date.toLocaleDateString("en-US", {
    month: "short",
    day: "numeric",
    year: "numeric"
  });
}

function getOnboardingSection(account) {
  return (
    account?.managerSections?.find(
      (section) => section.title === ONBOARDING_SECTION_TITLE || section.title === ACCOUNT_ONBOARDING_TITLE
    ) || null
  );
}

function buildSeedFromText(value) {
  return Array.from(String(value || "")).reduce((sum, char) => sum + char.charCodeAt(0), 0);
}

const EVE_NAME = "Eve";

const SIMULATED_JIRA_COMPONENTS = [
  "Sample Inquiry",
  "Orders for Product",
  "Phone Call",
  "Account Setup",
  "Training Request",
  "Billing Follow-up",
  "Portal Access"
];

const SIMULATED_JIRA_SUMMARIES = [
  "Customer asked for a status update on recent samples.",
  "AM needs CX follow-up on onboarding progress.",
  "Customer requested product order clarification.",
  "CX needs to confirm account setup details.",
  "Training users need access and scheduling support.",
  "Customer requested billing review and invoice explanation.",
  "Portal access needs troubleshooting before next review."
];

function createSampleDataFromMask(mask, projectedVolume, company) {
  const seed = buildSeedFromText(mask || company || "ACCOUNT");
  const projectedBase = Number.parseInt(String(projectedVolume || "").replace(/[^0-9]/g, ""), 10) || 0;
  const received = Math.max(6, (seed % 18) + Math.max(0, Math.floor(projectedBase / 1000)));
  const loggedOnline = received + (seed % 7) + 2;
  const inTesting = Math.max(2, Math.floor(received * 0.35));
  const awaitingReview = Math.max(1, Math.floor(received * 0.2));
  const onHold = seed % 3;
  const reported = Math.max(1, received - inTesting - awaitingReview - onHold);

  return {
    loggedOnline,
    received,
    autoSummary: `Auto-summary for ${maskOrFallback(mask)}: ${received} samples received and ${loggedOnline} logged online based on the created account mask.`,
    statuses: [
      { label: "Received", count: received },
      { label: "In testing", count: inTesting },
      { label: "Awaiting review", count: awaitingReview },
      { label: "On Hold", count: onHold },
      { label: "Reported", count: reported }
    ]
  };
}

function getSimulatedJiraBucket(account) {
  if (!account || account.id === "all-accounts") {
    return null;
  }

  const seed = buildSeedFromText(`${account.id}-${account.email}-${account.owner}`);
  const total = 4 + (seed % 5);
  const items = Array.from({ length: total }, (_, index) => {
    const componentType = SIMULATED_JIRA_COMPONENTS[(seed + index) % SIMULATED_JIRA_COMPONENTS.length];
    const summary = SIMULATED_JIRA_SUMMARIES[(seed + index * 2) % SIMULATED_JIRA_SUMMARIES.length];
    const statusGroup =
      index % 4 === 0 ? "Resolved" : index % 3 === 0 ? "In progress" : "Open";
    const status = statusGroup === "Resolved" ? "Done" : statusGroup === "In progress" ? "In Progress" : "Open";
    const dayOffset = (seed + index * 3) % 27;
    const date = new Date();
    date.setDate(date.getDate() - dayOffset);
    const resolvedAt = new Date(date);
    const turnaroundDays = 1 + ((seed + index * 5) % 8);
    resolvedAt.setDate(resolvedAt.getDate() + turnaroundDays);
    if (resolvedAt > new Date()) {
      resolvedAt.setTime(Date.now());
    }
    const keyNumber = 1000 + ((seed + index * 17) % 9000);

    return {
      id: `${account.id}-sim-jira-${index + 1}`,
      accountId: account.id,
      key: `CS-${keyNumber}`,
      summary,
      componentType,
      componentName: componentType,
      status,
      statusGroup,
      priority: index % 5 === 0 ? "High" : index % 2 === 0 ? "Medium" : "Low",
      createdAt: date.toISOString(),
      resolvedAt: statusGroup === "Resolved" ? resolvedAt.toISOString() : "",
      createdLabel: date.toLocaleDateString("en-US", { month: "short", day: "numeric" }),
      matchedMask: maskOrFallback(account.email)
    };
  });

  return {
    accountId: account.id,
    mask: maskOrFallback(account.email),
    total: items.length,
    open: items.filter((item) => item.statusGroup === "Open").length,
    inProgress: items.filter((item) => item.statusGroup === "In progress").length,
    resolved: items.filter((item) => item.statusGroup === "Resolved").length,
    componentTypes: items.reduce((map, item) => {
      map[item.componentType] = (map[item.componentType] || 0) + 1;
      return map;
    }, {}),
    recent: items,
    items
  };
}

function getJiraTicketAgeDays(ticket) {
  const createdAt = ticket?.createdAt || ticket?.createdDate || ticket?.createdIso || null;
  if (!createdAt) {
    return null;
  }

  const createdDate = new Date(createdAt);
  if (Number.isNaN(createdDate.getTime())) {
    return null;
  }

  const today = new Date();
  today.setHours(0, 0, 0, 0);
  createdDate.setHours(0, 0, 0, 0);
  return Math.max(0, Math.floor((today.getTime() - createdDate.getTime()) / 86400000));
}

function formatJiraAgeLabel(ticket) {
  const ageDays = getJiraTicketAgeDays(ticket);
  if (ageDays === null) {
    return "Age unavailable";
  }

  if (ticket?.statusGroup === "Resolved" || ticket?.status === "Resolved" || ticket?.status === "Done") {
    return `Resolved after ${ageDays}d`;
  }

  return `${ageDays}d open`;
}

function getJiraTurnaroundDays(ticket) {
  const createdAt = ticket?.createdAt || ticket?.createdDate || ticket?.createdIso || null;
  if (!createdAt) {
    return null;
  }

  const createdDate = new Date(createdAt);
  if (Number.isNaN(createdDate.getTime())) {
    return null;
  }

  const resolvedAt = ticket?.resolvedAt || ticket?.resolvedDate || null;
  const isResolved = ticket?.statusGroup === "Resolved" || ticket?.status === "Resolved" || ticket?.status === "Done";
  if (!resolvedAt && !isResolved) {
    return null;
  }

  const resolvedDate = resolvedAt ? new Date(resolvedAt) : new Date();
  if (Number.isNaN(resolvedDate.getTime())) {
    return null;
  }

  createdDate.setHours(0, 0, 0, 0);
  resolvedDate.setHours(0, 0, 0, 0);
  return Math.max(0, Math.floor((resolvedDate.getTime() - createdDate.getTime()) / 86400000));
}

function getAverageJiraTurnaroundDays(items) {
  const values = (items || [])
    .map((ticket) => getJiraTurnaroundDays(ticket))
    .filter((value) => Number.isFinite(value));

  if (!values.length) {
    return null;
  }

  return Math.round(values.reduce((sum, value) => sum + value, 0) / values.length);
}

function getVisibleSimulatedJiraAccounts() {
  return getLoginAccessibleAccounts().map((account) => getSimulatedJiraBucket(account)).filter(Boolean);
}

function createDefaultOnboardingIntake() {
  return {
    onboardingCompanyName: "",
    onboardingContactName: "",
    onboardingContactEmail: "",
    returnedAccountMask: "",
    onboardingStartDate: getTodayIsoDate(),
    projectedVolume: "",
    quoteNumber: "",
    accountStructureSetup: "",
    billingInfo: "",
    accountingRequirements: "",
    accountingHandoffSummary: "",
    netSuiteHandoffPreparedAt: "",
    accountCreationTicketAt: "",
    accountCreationTicketKey: "",
    equipmentList: "",
    equipmentTemplateName: "equipment-list-template.csv",
    equipmentUploadTicketAt: "",
    equipmentUploadTicketKey: "",
    horizonTrainingStatus: "",
    horizonTrainingSchedule: "",
    horizonTrainingUsers: "",
    pricingReviewPlan: "If projected volume matches the estimate, keep pricing. If not, schedule a customer conversation.",
    weeklyCheckInNotes: "",
    weeklyCheckInsCompletedAt: "",
    day45ReviewNotes: "",
    day45ReviewCompletedAt: "",
    day70ReviewNotes: "",
    day70ReviewCompletedAt: "",
    onboardingStartedAt: "",
    checklistStates: {},
    automationEnabled: true,
    automationLastRunAt: "",
    automationFeed: []
  };
}

function getPrivateOnboardingDraft() {
  if (!state.privateOnboardingDrafts[state.selectedLogin]) {
    state.privateOnboardingDrafts[state.selectedLogin] = createDefaultOnboardingIntake();
  }
  return state.privateOnboardingDrafts[state.selectedLogin];
}

function getEditableOnboardingIntake() {
  if (state.viewMode === "shared") {
    return null;
  }

  if (state.selectedId === "all-accounts") {
    return getPrivateOnboardingDraft();
  }

  return getOnboardingSection(getSelectedAccount())?.intake || null;
}

function getOnboardingContextAccount() {
  if (state.viewMode === "shared") {
    return null;
  }

  if (state.selectedId === "all-accounts") {
    const intake = getPrivateOnboardingDraft();
    return {
      id: "draft-onboarding",
      owner: intake.onboardingCompanyName?.trim() || "New account",
      email: intake.returnedAccountMask?.trim() || "",
      contactName: intake.onboardingContactName?.trim() || "Primary Contact",
      contactEmail: intake.onboardingContactEmail?.trim() || getCurrentLogin()?.microsoftEmail || "",
      team: state.selectedLogin
    };
  }

  return getSelectedAccount();
}

function normalizeAccountIdFragment(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .slice(0, 40) || `acct-${Math.floor(Math.random() * 900 + 100)}`;
}

function addAccessibleAccount(loginName, accountId) {
  const login = state.loginDirectory.find((person) => person.name === loginName);
  if (!login) {
    return;
  }

  if (!Array.isArray(login.accessibleAccountIds)) {
    login.accessibleAccountIds = [];
  }

  if (!login.accessibleAccountIds.includes(accountId)) {
    login.accessibleAccountIds.push(accountId);
  }
}

function convertOnboardingDraftToAccount(intake) {
  const company = intake.onboardingCompanyName?.trim();
  const mask = intake.returnedAccountMask?.trim();
  const contactName = intake.onboardingContactName?.trim();
  const contactEmail = intake.onboardingContactEmail?.trim();

  const missing = [];
  if (!company) {
    missing.push("Company name");
  }
  if (!mask) {
    missing.push("Returned account number");
  }
  if (!contactName) {
    missing.push("Contact person");
  }
  if (!contactEmail) {
    missing.push("Contact email");
  }

  if (missing.length) {
    return { created: false, missing };
  }

  const nextIdBase = normalizeAccountIdFragment(company);
  let nextId = nextIdBase;
  let counter = 2;
  while (state.accounts.some((account) => account.id === nextId)) {
    nextId = `${nextIdBase}-${counter}`;
    counter += 1;
  }

  const manager = state.selectedLogin;
  const newAccount = createAccount(nextId, company, mask, manager, {
    contactName,
    contactEmail
  });

  const sampleSection = newAccount.managerSections.find((section) => section.title === "Sample Tracker");
  if (sampleSection?.sampleData) {
    sampleSection.sampleData = createSampleDataFromMask(mask, intake.projectedVolume, company);
  }

  const onboardingSection = getOnboardingSection(newAccount);
  if (onboardingSection?.intake) {
    onboardingSection.title = ACCOUNT_ONBOARDING_TITLE;
    onboardingSection.detail = "Track follow-through after the account number is created, including revenue handoff, equipment, training, and milestone reviews.";
    onboardingSection.intake = {
      ...createDefaultOnboardingIntake(),
      ...intake,
      onboardingCompanyName: company,
      onboardingContactName: contactName,
      onboardingContactEmail: contactEmail,
      returnedAccountMask: mask
    };
  }

  state.accounts.unshift(newAccount);
  [state.selectedLogin, "D Williams", "Kelsie Miller", "Jordan Dickey"].forEach((name) =>
    addAccessibleAccount(name, newAccount.id)
  );

  state.privateOnboardingDrafts[state.selectedLogin] = createDefaultOnboardingIntake();
  state.selectedId = newAccount.id;

  return { created: true, account: newAccount };
}

function getOnboardingChecklistItems(account, intake) {
  const checklistStates = intake.checklistStates || {};
  const resolveComplete = (key, derived) =>
    typeof checklistStates[key] === "boolean" ? checklistStates[key] : derived;

  const projectedVolumeReady = Boolean(intake.projectedVolume?.trim());
  const accountStructureReady = Boolean(intake.accountStructureSetup?.trim());
  const accountCreationReady = Boolean(intake.accountCreationTicketAt);
  const netSuiteReady = Boolean(intake.netSuiteHandoffPreparedAt);
  const equipmentReady = Boolean(intake.equipmentUploadTicketAt);
  const horizonReady = Boolean(intake.horizonTrainingSchedule?.trim());
  const day45Ready = Boolean(intake.day45ReviewCompletedAt);
  const weeklyReady = Boolean(intake.weeklyCheckInsCompletedAt);
  const day70Ready = Boolean(intake.day70ReviewCompletedAt);

  return [
    {
      key: "projectedVolume",
      label: "Projected volume",
      detail: projectedVolumeReady ? intake.projectedVolume : "Add the expected sample or account volume.",
      complete: resolveComplete("projectedVolume", projectedVolumeReady)
    },
    {
      key: "accountStructureSetup",
      label: "Account structure set up",
      detail: accountStructureReady ? intake.accountStructureSetup : "Document the account structure and setup approach.",
      complete: resolveComplete("accountStructureSetup", accountStructureReady)
    },
    {
      key: "accountCreationTicket",
      label: "Create CX Jira ticket for account creation",
      detail: intake.accountCreationTicketAt
        ? `Created ${intake.accountCreationTicketAt}${intake.accountCreationTicketKey ? ` as ${intake.accountCreationTicketKey}` : ""}.`
        : "Create a CX request so the account can be built.",
      complete: resolveComplete("accountCreationTicket", accountCreationReady)
    },
    {
      key: "netSuiteHandoff",
      label: "Connect to NetSuite and send accounting handoff",
      detail: intake.netSuiteHandoffPreparedAt
        ? `Accounting handoff prepared ${intake.netSuiteHandoffPreparedAt}.`
        : "Send the required info to accounting once requirements are confirmed.",
      complete: resolveComplete("netSuiteHandoff", netSuiteReady)
    },
    {
      key: "equipmentUploadTicket",
      label: "Create Jira ticket to upload equipment",
      detail: intake.equipmentUploadTicketAt
        ? `Created ${intake.equipmentUploadTicketAt}${intake.equipmentUploadTicketKey ? ` as ${intake.equipmentUploadTicketKey}` : ""}.`
        : "Submit the equipment upload request to CX.",
      complete: resolveComplete("equipmentUploadTicket", equipmentReady)
    },
    {
      key: "horizonTraining",
      label: "Schedule HORIZON training",
      detail: horizonReady
        ? `${intake.horizonTrainingSchedule}${intake.horizonTrainingUsers?.trim() ? ` · Users: ${intake.horizonTrainingUsers}` : ""}`
        : "Set the training date and attendees.",
      complete: resolveComplete("horizonTraining", horizonReady)
    },
    {
      key: "day45Review",
      label: "45-day review",
      detail: day45Ready
        ? `Completed ${intake.day45ReviewCompletedAt}.`
        : "Check projected volume, sample volume, and complaints.",
      complete: resolveComplete("day45Review", day45Ready)
    },
    {
      key: "weeklyCheckIns",
      label: "Weekly check-ins",
      detail: weeklyReady
        ? `Marked complete ${intake.weeklyCheckInsCompletedAt}.`
        : "Run weekly check-ins during onboarding.",
      complete: resolveComplete("weeklyCheckIns", weeklyReady)
    },
    {
      key: "day70Review",
      label: "70-day pricing check",
      detail: day70Ready
        ? `Completed ${intake.day70ReviewCompletedAt}.`
        : "Review sample volume and schedule pricing discussion if off track.",
      complete: resolveComplete("day70Review", day70Ready)
    }
  ];
}

function getOnboardingCompletionPercent(account, intake) {
  const items = getOnboardingChecklistItems(account, intake);
  const completed = items.filter((item) => item.complete).length;
  return {
    completed,
    total: items.length,
    percent: items.length ? Math.round((completed / items.length) * 100) : 0,
    items
  };
}

function getOnboardingTimeline(intake) {
  const startDate = intake.onboardingStartDate || getTodayIsoDate();
  return [
    {
      label: "Start",
      date: formatTimelineDate(startDate),
      detail: "Open onboarding, gather requirements, and begin account setup."
    },
    {
      label: "Day 45",
      date: formatTimelineDate(addDaysToIsoDate(startDate, 45)),
      detail: "Check projected volume, sample volume, complaints, and complete weekly check-ins."
    },
    {
      label: "Day 70",
      date: formatTimelineDate(addDaysToIsoDate(startDate, 70)),
      detail: "Review sample volume and schedule a pricing meeting if volume is off track."
    },
    {
      label: "Day 90",
      date: formatTimelineDate(addDaysToIsoDate(startDate, 90)),
      detail: "Finish the 90-day onboarding review and confirm final pricing direction."
    }
  ];
}

function getOnboardingAutomationFeed(account, intake) {
  const timeline = getOnboardingTimeline(intake);
  return [
    {
      label: "CX account creation automation",
      status: intake.accountCreationTicketAt ? "Complete" : "Watching",
      detail: intake.accountCreationTicketAt
        ? `Ticket created ${intake.accountCreationTicketAt}${intake.accountCreationTicketKey ? ` as ${intake.accountCreationTicketKey}` : ""}.`
        : "Will flag when onboarding details are ready for CX account creation."
    },
    {
      label: "Accounting handoff automation",
      status: intake.netSuiteHandoffPreparedAt ? "Prepared" : "Watching",
      detail: intake.netSuiteHandoffPreparedAt
        ? `Prepared ${intake.netSuiteHandoffPreparedAt} for ${account?.owner || "new account"}.`
        : "Will remind accounting handoff once requirements and billing info are entered."
    },
    {
      label: "45-day checkpoint",
      status: intake.day45ReviewCompletedAt ? "Complete" : "Scheduled",
      detail: intake.day45ReviewCompletedAt
        ? `Completed ${intake.day45ReviewCompletedAt}.`
        : `Scheduled for ${timeline.find((item) => item.label === "Day 45")?.date || "Day 45"}.`
    },
    {
      label: "Weekly check-in cadence",
      status: intake.weeklyCheckInsCompletedAt ? "Complete" : "Active",
      detail: intake.weeklyCheckInsCompletedAt
        ? `Completed ${intake.weeklyCheckInsCompletedAt}.`
        : "Automation keeps weekly check-ins visible during the onboarding window."
    },
    {
      label: "70-day pricing checkpoint",
      status: intake.day70ReviewCompletedAt ? "Complete" : "Scheduled",
      detail: intake.day70ReviewCompletedAt
        ? `Completed ${intake.day70ReviewCompletedAt}.`
        : `Scheduled for ${timeline.find((item) => item.label === "Day 70")?.date || "Day 70"}.`
    },
    {
      label: "90-day pricing review",
      status: "Scheduled",
      detail: `Scheduled for ${timeline.find((item) => item.label === "Day 90")?.date || "Day 90"}.`
    }
  ];
}

function runOnboardingAutomationSweep(account, intake) {
  intake.automationFeed = getOnboardingAutomationFeed(account, intake);
  intake.automationLastRunAt = new Date().toLocaleString("en-US");
}

const state = {
  selectedLogin: "D Williams",
  viewMode: "private",
  selectedId: "acct-mi",
  loginOverlayOpen: false,
  accountFilterQuery: "",
  accountListScroll: 0,
  accounts: [
    createAccount("acct-mi", "Motion Industries", "MIPOWR", "Marian Kiley"),
    createAccount("acct-voltagrid", "Voltagrid", "846164", "Mel Prieve"),
    createAccount("acct-meta", "Meta", "273187", "Mel Prieve"),
    createAccount("acct-komatsu", "Komatsu", "KM", "Sade Thomas"),
    createAccount("acct-uslubricants", "US Lubricants", "USLUBE", "Sasha Dennison"),
    createAccount("acct-oilanalyzers", "Oil Analyzers", "OILANA", "Sasha Dennison", {
      contactName: "Ryan Lawry; Zach Austin"
    }),
    createAccount("acct-jglubricants", "JG Lubricants", "JGLUBR", "Sasha Dennison", {
      contactName: "Tom Johnson"
    }),
    createAccount("acct-chevron", "Chevron", "LUB", "Sasha Dennison"),
    createAccount("acct-kubota", "Kubota", "KUBONA", "Prad Marur")
  ],
  loginDirectory: [
    {
      name: "Marian Kiley",
      role: "Account Manager",
      primaryRegion: "NA",
      microsoftEmail: "marian.kiley@northstarops.test",
      avatar: "MK",
      calendarConnected: true,
      aiMode: "Account insights and AM prep",
      sharedAccess: true,
      accessibleAccountIds: ["acct-mi"]
    },
    {
      name: "Curt Watson",
      role: "Account Manager",
      primaryRegion: "NA",
      microsoftEmail: "curt.watson@northstarops.test",
      avatar: "CW",
      calendarConnected: true,
      aiMode: "Whitespace review and follow-up suggestions",
      sharedAccess: true,
      accessibleAccountIds: []
    },
    {
      name: "Mel Prieve",
      role: "Account Manager",
      primaryRegion: "NA",
      microsoftEmail: "mel.prieve@northstarops.test",
      avatar: "MP",
      calendarConnected: true,
      aiMode: "Portfolio recap and spend monitoring",
      sharedAccess: true,
      accessibleAccountIds: ["acct-voltagrid", "acct-meta"]
    },
    {
      name: "Sasha Dennison",
      role: "Account Manager",
      primaryRegion: "NA",
      microsoftEmail: "sasha.dennison@northstarops.test",
      avatar: "SD",
      calendarConnected: true,
      aiMode: "Renewal planning and customer notes",
      sharedAccess: true,
      accessibleAccountIds: ["acct-uslubricants", "acct-oilanalyzers", "acct-jglubricants", "acct-chevron"]
    },
    {
      name: "Prad Marur",
      role: "Account Manager",
      primaryRegion: "NA",
      microsoftEmail: "prad.marur@northstarops.test",
      avatar: "PM",
      calendarConnected: true,
      aiMode: "Opportunity tracking and account planning",
      sharedAccess: true,
      accessibleAccountIds: ["acct-kubota"]
    },
    {
      name: "Kelsie Miller",
      role: "Customer Engagement Manager",
      primaryRegion: "NA",
      microsoftEmail: "kelsie.miller@northstarops.test",
      avatar: "KM",
      calendarConnected: true,
      aiMode: "Engagement planning and meeting preparation",
      sharedAccess: true,
      accessibleAccountIds: ["acct-mi", "acct-voltagrid", "acct-meta", "acct-komatsu", "acct-uslubricants", "acct-oilanalyzers", "acct-jglubricants", "acct-chevron", "acct-kubota"]
    },
    {
      name: "Sade Thomas",
      role: "Customer Engagement Specialist",
      primaryRegion: "NA",
      microsoftEmail: "sade.thomas@northstarops.test",
      avatar: "ST",
      calendarConnected: true,
      aiMode: "Engagement follow-up and action items",
      sharedAccess: true,
      accessibleAccountIds: ["acct-komatsu"]
    },
    {
      name: "Rob Dixon",
      role: "Sales/Marketing",
      primaryRegion: "NA",
      microsoftEmail: "rob.dixon@northstarops.test",
      avatar: "RD",
      calendarConnected: true,
      aiMode: "Campaign coordination and account context",
      sharedAccess: true,
      accessibleAccountIds: []
    },
    {
      name: "Sarah DeWolf",
      role: "Sales/Marketing",
      primaryRegion: "NA",
      microsoftEmail: "sarah.dewolf@northstarops.test",
      avatar: "SD",
      calendarConnected: true,
      aiMode: "Pipeline messaging and launch planning",
      sharedAccess: true,
      accessibleAccountIds: []
    },
    {
      name: "Jordan Dickey",
      role: "Salesforce Admin",
      primaryRegion: "NA",
      microsoftEmail: "jordan.dickey@northstarops.test",
      avatar: "JD",
      calendarConnected: false,
      aiMode: "CRM admin assistant and workflow diagnostics",
      sharedAccess: true,
      accessibleAccountIds: ["acct-mi", "acct-voltagrid", "acct-meta", "acct-komatsu", "acct-uslubricants", "acct-oilanalyzers", "acct-jglubricants", "acct-chevron", "acct-kubota"]
    },
    {
      name: "D Williams",
      role: "Customer Experience Tech 2",
      primaryRegion: "NA",
      microsoftEmail: "dwilliams@polarislabs.com",
      avatar: "DW",
      calendarConnected: true,
      aiMode: "Customer experience support, dashboard triage, and calendar context",
      sharedAccess: true,
      accessibleAccountIds: ["acct-mi", "acct-voltagrid", "acct-meta", "acct-komatsu", "acct-uslubricants", "acct-oilanalyzers", "acct-jglubricants", "acct-chevron", "acct-kubota"]
    }
  ],
  sharedItems: [
    {
      title: "Revenue",
      detail: "$482,300 total shared pipeline value across the sandbox workspace."
    },
    {
      title: "Samples submitted online",
      detail: "184 total sample submissions have been created through the online intake flow."
    },
    {
      title: "Samples received",
      detail: "161 submitted samples have been received and logged in the shared workspace."
    },
    {
      title: "Outstanding invoices",
      detail: "$73,450 is currently outstanding across all visible sandbox accounts."
    },
    {
      title: "Everybody calendar",
      detail: "Team calendar consolidates Microsoft availability for account managers, engagement, marketing, and admin support."
    }
  ],
  teamCalendar: [
    { title: "Weekly revenue review", owner: "Marian Kiley", time: "Monday 9:00 AM" },
    { title: "Customer engagement standup", owner: "Kelsie Miller", time: "Tuesday 10:30 AM" },
    { title: "Sales and marketing sync", owner: "Rob Dixon and Sarah DeWolf", time: "Wednesday 1:00 PM" },
    { title: "Salesforce routing check-in", owner: "Jordan Dickey", time: "Thursday 2:30 PM" },
    { title: "Executive dashboard recap", owner: "D Williams", time: "Friday 11:00 AM" }
  ],
  pipelineStages: ["Discovery", "Qualification", "Evaluation", "Proposal", "Negotiation", "Closed Won"],
  pipelineLeads: [
    { id: "lead-na-1", company: "Atlas Mining", region: "NA", stage: "Discovery", fitScore: 92, signal: "High sample growth in heavy equipment segment", owner: "Rob Dixon", nextStep: "Coordinate AM intro and sample strategy review." },
    { id: "lead-na-2", company: "Delta Fleet Services", region: "NA", stage: "Proposal", fitScore: 86, signal: "Needs consolidated oil analysis coverage across sites", owner: "Sarah DeWolf", nextStep: "Prepare pricing and onboarding plan." },
    { id: "lead-na-3", company: "Summit Aggregates", region: "NA", stage: "Evaluation", fitScore: 79, signal: "Early product interest with recurring sample demand", owner: "Rob Dixon", nextStep: "Surface to regional AM for follow-up." },
    { id: "lead-emea-1", company: "Nordic Transport Group", region: "EMEA", stage: "Qualification", fitScore: 84, signal: "Requested multi-site lubrication monitoring", owner: "Sarah DeWolf", nextStep: "Confirm regional coverage and quote path." },
    { id: "lead-apac-1", company: "Pacific Marine Works", region: "APAC", stage: "Discovery", fitScore: 81, signal: "Asking for training plus reporting workflow support", owner: "Rob Dixon", nextStep: "Review onboarding requirements and service fit." }
  ],
  workspaceSearchQuery: "",
  workspaceNoteDraft: "",
  pipelineFilterRegion: "ALL",
  pipelineFilterStage: "ALL",
  pipelineStageDraft: "",
  eveDraft: "",
  activeOperationsTab: "Sample Tracker",
  eveConversations: {
    "D Williams": [
      {
        id: "eve-msg-1",
        sender: EVE_NAME,
        text: "I’m watching account trends, CX load, and regional pipeline movement. Ask me for urgent issues, account risk, or new opportunities in your region."
      }
    ]
  },
  jiraBaseUrl: "",
  microsoftIntegration: {
    configured: false,
    connected: false,
    email: "",
    displayName: "",
    connectedAt: "",
    scopes: [],
    teamsUrl: "https://teams.microsoft.com",
    officeUrl: "https://www.office.com",
    authUrl: "/auth/microsoft/start"
  },
  genesysIntegration: {
    configured: false,
    mode: "mock",
    region: "",
    orgName: "",
    generatedAt: "",
    portalUrl: "",
    developerUrl: "https://developer.genesys.cloud",
    scopes: [],
    byRegion: {}
  },
  privateOnboardingDrafts: {
    "D Williams": createDefaultOnboardingIntake()
  },
  netSuiteIntegration: {
    configured: false,
    mode: "mock",
    generatedAt: "",
    overall: {
      ytdRevenue: 0,
      projectedYearEnd: 0,
      priorYearYtd: 0,
      outstandingInvoices: 0
    }
  },
  jiraModalAccountId: null,
  sharedWorkspaceNotes: [
    { id: "shared-note-1", scope: "Shared workspace", author: "D Williams", text: "Shared dashboard is ready for revenue, sample, invoice, and calendar review." }
  ],
  privateWorkspaceNotes: {
    "D Williams": [
      { id: "private-note-1", scope: "All accounts", author: "D Williams", text: "Use the All view for a combined portfolio-level look before drilling into a single account." }
    ]
  },
  privateWorkspaceJira: {
    "D Williams": []
  },
  selectedImportedJiraAccountId: null,
  selectedImportedJiraComponent: "All",
  selectedSummaryBreakdown: null,
  panelLayouts: loadSavedPanelLayouts(),
  draggedPanelId: null
};

let dashboardStateSyncTimer = null;
let dashboardBootstrapComplete = false;

function buildPersistableState() {
  return {
    accounts: state.accounts,
    loginDirectory: state.loginDirectory,
    sharedItems: state.sharedItems,
    teamCalendar: state.teamCalendar,
    pipelineStages: state.pipelineStages,
    pipelineLeads: state.pipelineLeads,
    privateOnboardingDrafts: state.privateOnboardingDrafts,
    sharedWorkspaceNotes: state.sharedWorkspaceNotes,
    privateWorkspaceNotes: state.privateWorkspaceNotes,
    privateWorkspaceJira: state.privateWorkspaceJira,
    eveConversations: state.eveConversations,
    panelLayouts: state.panelLayouts
  };
}

function applyPersistedDashboardState(payload) {
  if (!payload || typeof payload !== "object") {
    return;
  }

  if (Array.isArray(payload.accounts)) {
    state.accounts = payload.accounts;
  }

  if (Array.isArray(payload.loginDirectory)) {
    state.loginDirectory = payload.loginDirectory;
  }

  if (Array.isArray(payload.sharedItems)) {
    state.sharedItems = payload.sharedItems;
  }

  if (Array.isArray(payload.teamCalendar)) {
    state.teamCalendar = payload.teamCalendar;
  }

  if (Array.isArray(payload.pipelineStages)) {
    state.pipelineStages = payload.pipelineStages;
  }

  if (Array.isArray(payload.pipelineLeads)) {
    state.pipelineLeads = payload.pipelineLeads;
  }

  if (payload.privateOnboardingDrafts && typeof payload.privateOnboardingDrafts === "object") {
    state.privateOnboardingDrafts = payload.privateOnboardingDrafts;
  }

  if (Array.isArray(payload.sharedWorkspaceNotes)) {
    state.sharedWorkspaceNotes = payload.sharedWorkspaceNotes;
  }

  if (payload.privateWorkspaceNotes && typeof payload.privateWorkspaceNotes === "object") {
    state.privateWorkspaceNotes = payload.privateWorkspaceNotes;
  }

  if (payload.privateWorkspaceJira && typeof payload.privateWorkspaceJira === "object") {
    state.privateWorkspaceJira = payload.privateWorkspaceJira;
  }

  if (payload.eveConversations && typeof payload.eveConversations === "object") {
    state.eveConversations = payload.eveConversations;
  }

  if (payload.panelLayouts && typeof payload.panelLayouts === "object") {
    state.panelLayouts = payload.panelLayouts;
  }
}

async function syncDashboardState() {
  if (!dashboardBootstrapComplete) {
    return;
  }

  try {
    await fetch("/api/dashboard-state", {
      method: "PUT",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify(buildPersistableState())
    });
  } catch {
    // Keep the local UI responsive even if the persistence layer is unavailable.
  }
}

function queueDashboardStateSync() {
  if (!dashboardBootstrapComplete) {
    return;
  }

  window.clearTimeout(dashboardStateSyncTimer);
  dashboardStateSyncTimer = window.setTimeout(() => {
    syncDashboardState();
  }, 400);
}

async function loadDashboardState() {
  try {
    const response = await fetch("/api/dashboard-state");
    if (!response.ok) {
      dashboardBootstrapComplete = true;
      return;
    }

    const payload = await response.json();
    if (payload?.data) {
      applyPersistedDashboardState(payload.data);
      dashboardBootstrapComplete = true;
      return;
    }

    dashboardBootstrapComplete = true;
    await syncDashboardState();
  } catch {
    // Fall back to seeded local state if the persistence layer is unavailable.
    dashboardBootstrapComplete = true;
  }
}

function applyImportedComplaintData(account) {
  const complaintSection = account.managerSections.find((section) => section.title === "Complaints Trackers");
  if (complaintSection) {
    complaintSection.complaintData = createComplaintData(account.id, account.email, account.team);
  }
}

state.accounts.forEach((account) => {
  applyImportedComplaintData(account);
});

const accountList = document.querySelector("#account-list");
const accountFilterInput = document.querySelector("#account-filter-input");
const loginSelect = document.querySelector("#login-select");
const viewOptionSelect = document.querySelector("#view-option-select");
const loginSummary = document.querySelector("#login-summary");
const loginOverlay = document.querySelector("#login-overlay");
const loginCardList = document.querySelector("#login-card-list");
const activeUserCard = document.querySelector("#active-user-card");
const workspaceSummary = document.querySelector("#workspace-summary");
const privateViewButton = document.querySelector("#private-view-button");
const sharedViewButton = document.querySelector("#shared-view-button");
const homepagePanel = document.querySelector("#homepage-panel");
const accountName = document.querySelector("#account-name");
const heroLabel = document.querySelector("#hero-label");
const heroPlan = document.querySelector("#hero-plan");
const heroSummary = document.querySelector("#hero-summary");
const metricMrr = document.querySelector("#metric-mrr");
const metricHealth = document.querySelector("#metric-health");
const metricSpend = document.querySelector("#metric-spend");
const ownerInput = document.querySelector("#owner-input");
const emailInput = document.querySelector("#email-input");
const teamInput = document.querySelector("#team-input");
const regionInput = document.querySelector("#region-input");
const spendCurrent = document.querySelector("#spend-current");
const spendProjected = document.querySelector("#spend-projected");
const spendDelta = document.querySelector("#spend-delta");
const spendBars = document.querySelector("#spend-bars");
const calendarPanelContent = document.querySelector("#calendar-panel-content");
const securityBadge = document.querySelector("#security-badge");
const msBadge = document.querySelector("#ms-badge");
const msAccess = document.querySelector("#ms-access");
const sharedPanel = document.querySelector("#shared-panel");
const sharedList = document.querySelector("#shared-list");
const pipelinesPanel = document.querySelector("#pipelines-panel");
const pipelinesPanelContent = document.querySelector("#pipelines-panel-content");
const dashboardGrid = document.querySelector("#dashboard-grid");
const onboardingPanel = document.querySelector("#onboarding-panel");
const onboardingWorkspace = document.querySelector("#onboarding-workspace");
const summaryTablePanel = document.querySelector("#summary-table-panel");
const summaryTableMeta = document.querySelector("#summary-table-meta");
const summaryTableBody = document.querySelector("#summary-table-body");
const summaryBreakdown = document.querySelector("#summary-breakdown");
const profilePanel = document.querySelector("#profile-panel");
const workspaceSearch = document.querySelector("#workspace-search");
const workspaceNoteInput = document.querySelector("#workspace-note-input");
const addWorkspaceNoteButton = document.querySelector("#add-workspace-note-button");
const workspaceNotes = document.querySelector("#workspace-notes");
const workspaceSearchResults = document.querySelector("#workspace-search-results");
const aiPanel = document.querySelector("#ai-panel");
const managerWorkspacePanel = document.querySelector("#manager-workspace-panel");
const managerWorkspace = document.querySelector("#manager-workspace");
const resetLayoutButton = document.querySelector("#reset-layout-button");
const jiraModal = document.querySelector("#jira-modal");
const jiraModalForm = document.querySelector("#jira-modal-form");
const jiraModalAccountSelect = document.querySelector("#jira-modal-account-select");
const jiraModalCompany = document.querySelector("#jira-modal-company");
const jiraModalMask = document.querySelector("#jira-modal-mask");
const jiraModalContactName = document.querySelector("#jira-modal-contact-name");
const jiraModalContactEmail = document.querySelector("#jira-modal-contact-email");
const jiraModalCc = document.querySelector("#jira-modal-cc");
const jiraModalTitleInput = document.querySelector("#jira-modal-title-input");
const jiraModalDescription = document.querySelector("#jira-modal-description");
const jiraModalPriority = document.querySelector("#jira-modal-priority");
const jiraModalAttachments = document.querySelector("#jira-modal-attachments");
const jiraModalAttachmentList = document.querySelector("#jira-modal-attachment-list");
const saveButton = document.querySelector("#save-button");
const toggleStatusButton = document.querySelector("#toggle-status-button");
const newAccountButton = document.querySelector("#new-account-button");
const toast = document.querySelector("#toast");

function createAccount(id, company, mask, manager, options = {}) {
  const contactName = options.contactName || createContactName(company);
  const contactEmail = options.contactEmail || createContactEmail(company);
  return {
    id,
    owner: company,
    email: mask,
    team: manager,
    contactName,
    contactEmail,
    region: "NA",
    status: "active",
    plan: "Custom",
    summary: `Assigned to ${manager}.`,
    mrr: 0,
    loginHealth: 0,
    seats: 0,
    renewal: "",
    billingStatus: "",
    spend: {
      currentMonth: 0,
      projected: 0,
      previousMonth: 0,
      monthly: [
        { label: "Q1", value: 0 },
        { label: "Q2", value: 0 },
        { label: "Q3", value: 0 },
        { label: "Q4", value: 0 }
      ],
      charges: []
    },
    notes: [
      {
        id: `${id}-note-1`,
        scope: company || "Account",
        author: manager,
        text: `Keep this workspace focused on ${company || "account"} follow-up, sample movement, and next actions.`
      }
    ],
    security: {
      mfa: false,
      sso: true,
      audit: false
    },
    activity: [],
    managerSections: [
      {
        title: "Sample Tracker",
        detail: "Track submitted, received, and pending samples tied to this account.",
        sampleData: {
          loggedOnline: 18,
          received: 14,
          statuses: [
            { label: "Received", count: 14 },
            { label: "In testing", count: 6 },
            { label: "Awaiting review", count: 4 },
            { label: "On Hold", count: 2 },
            { label: "Reported", count: 4 }
          ]
        }
      },
      {
        title: "Call Log",
        detail: "Track total calls by account mask and call type.",
        callLogData: {
          total: 21,
          entries: [
            { label: `${maskOrFallback(mask)} / Intro`, count: 5 },
            { label: `${maskOrFallback(mask)} / Follow-up`, count: 9 },
            { label: `${maskOrFallback(mask)} / Escalation`, count: 3 },
            { label: `${maskOrFallback(mask)} / Quarterly review`, count: 4 }
          ]
        }
      },
      {
        title: "Complaints Trackers",
        detail: "Track the number of recorded complaints by account mask.",
        complaintData: {
          total: 4,
          entries: [
            { label: `${maskOrFallback(mask)} / Recorded complaints`, count: 4 }
          ],
          records: [
            {
              id: `${id}-complaint-1`,
              title: "Turnaround time concern",
              status: "Open",
              owner: manager,
              date: "Apr 03",
              detail: `${maskOrFallback(mask)} · Customer asked for an update on delayed reporting.`
            },
            {
              id: `${id}-complaint-2`,
              title: "Sample status visibility",
              status: "In review",
              owner: manager,
              date: "Apr 07",
              detail: `${maskOrFallback(mask)} · Customer reported limited visibility into current sample status.`
            },
            {
              id: `${id}-complaint-3`,
              title: "Billing follow-up",
              status: "Open",
              owner: manager,
              date: "Apr 11",
              detail: `${maskOrFallback(mask)} · Customer requested clarification on an invoice line item.`
            },
            {
              id: `${id}-complaint-4`,
              title: "Portal login friction",
              status: "Resolved",
              owner: manager,
              date: "Apr 15",
              detail: `${maskOrFallback(mask)} · Customer had trouble logging in and needed reset assistance.`
            }
          ]
        }
      },
      {
        title: "JIRA tickets",
        detail: "Track all requests AMs make to CX for visibility.",
        jiraData: {
          total: 7,
          entries: [
            { label: "Open CX visibility requests", count: 3 },
            { label: "In progress with CX", count: 2 },
            { label: "Resolved for AM follow-up", count: 2 }
          ],
          requests: [
            {
              id: `${id}-jira-1`,
              title: "Clarify CX visibility on account setup",
              priority: "High",
              status: "Open",
              owner: manager
            }
          ]
        }
      },
      {
        title: ONBOARDING_SECTION_TITLE,
        detail: "Manage the first 90 days of onboarding with milestone tracking, handoffs, and checklist progress.",
        intake: {
          onboardingStartDate: getTodayIsoDate(),
          projectedVolume: "",
          accountStructureSetup: "",
          billingInfo: "",
          accountingRequirements: "",
          accountingHandoffSummary: "",
          netSuiteHandoffPreparedAt: "",
          accountCreationTicketAt: "",
          accountCreationTicketKey: "",
          equipmentList: "",
          equipmentTemplateName: "equipment-list-template.csv",
          equipmentUploadTicketAt: "",
          equipmentUploadTicketKey: "",
          horizonTrainingStatus: "",
          horizonTrainingSchedule: "",
          horizonTrainingUsers: "",
          pricingReviewPlan: "If projected volume matches the estimate, keep pricing. If not, schedule a customer conversation.",
          weeklyCheckInNotes: "",
          weeklyCheckInsCompletedAt: "",
          day45ReviewNotes: "",
          day45ReviewCompletedAt: "",
          day70ReviewNotes: "",
          day70ReviewCompletedAt: "",
          onboardingStartedAt: ""
        }
      },
      {
        title: "Personal Calendar",
        detail: "Show Outlook-style meetings and new emails relevant to this account.",
        outlookData: {
          meetings: [
            { title: `${company || "Account"} weekly touch base`, time: "Today 10:00 AM", attendees: `${manager}, CX, Customer contact` },
            { title: `${maskOrFallback(mask)} sample review`, time: "Tomorrow 1:30 PM", attendees: `${manager}, Lab ops` }
          ],
          emails: [
            { subject: `${company || "Account"} pricing follow-up`, sender: "customer@company.com", time: "8:14 AM" },
            { subject: `${maskOrFallback(mask)} sample status update`, sender: "labupdates@polarislabs.com", time: "Yesterday" }
          ]
        }
      }
    ]
  };
}

function maskOrFallback(mask) {
  return mask || "No mask";
}

function createContactName(company) {
  if (!company) {
    return "Primary Contact";
  }

  const cleaned = company.replace(/[^A-Za-z0-9 ]/g, "").trim();
  const parts = cleaned.split(/\s+/).filter(Boolean);
  const first = parts[0] || "Primary";
  const second = parts[1] || "Contact";
  return `${first} ${second}`;
}

function createContactEmail(company) {
  if (!company) {
    return "primary.contact@customer.test";
  }

  const slug = company.toLowerCase().replace(/[^a-z0-9]+/g, ".").replace(/^\.+|\.+$/g, "");
  return `contact@${slug || "customer"}.test`;
}

function getCurrentLogin() {
  return state.loginDirectory.find((person) => person.name === state.selectedLogin);
}

function syncSelectedLogin() {
  if (!Array.isArray(state.loginDirectory) || state.loginDirectory.length === 0) {
    return null;
  }

  const current = getCurrentLogin();
  if (current) {
    return current;
  }

  state.selectedLogin = state.loginDirectory[0].name;
  return state.loginDirectory[0];
}

function getAccountById(accountId) {
  return state.accounts.find((account) => account.id === accountId) || null;
}

function isCreatorView() {
  return state.selectedLogin === "D Williams";
}

function getLoginAccessibleAccounts() {
  const login = getCurrentLogin();
  if (!login) {
    return [];
  }

  if (Array.isArray(login.accessibleAccountIds)) {
    const allowedIds = new Set(login.accessibleAccountIds);
    return state.accounts.filter((account) => allowedIds.has(account.id));
  }

  return [];
}

function buildAggregateAccount(accounts) {
  const monthlyLabels = ["Q1", "Q2", "Q3", "Q4"];
  const monthly = monthlyLabels.map((label) => ({
    label,
    value: accounts.reduce((sum, account) => {
      const match = account.spend.monthly.find((entry) => entry.label === label);
      return sum + (match?.value || 0);
    }, 0)
  }));

  const receivedTotal = accounts.reduce(
    (sum, account) => sum + (account.managerSections.find((section) => section.title === "Sample Tracker")?.sampleData?.received || 0),
    0
  );
  const loggedOnlineTotal = accounts.reduce(
    (sum, account) => sum + (account.managerSections.find((section) => section.title === "Sample Tracker")?.sampleData?.loggedOnline || 0),
    0
  );
  const sampleStatuses = ["Received", "In testing", "Awaiting review", "On Hold", "Reported"].map((label) => ({
    label,
    count: accounts.reduce((sum, account) => {
      const section = account.managerSections.find((item) => item.title === "Sample Tracker");
      const status = section?.sampleData?.statuses?.find((entry) => entry.label === label);
      return sum + (status?.count || 0);
    }, 0)
  }));

  const callEntries = ["Intro", "Follow-up", "Escalation", "Quarterly review"].map((type) => ({
    label: type,
    count: accounts.reduce((sum, account) => {
      const section = account.managerSections.find((item) => item.title === "Call Log");
      const entry = section?.callLogData?.entries?.find((item) => item.label.includes(` / ${type}`));
      return sum + (entry?.count || 0);
    }, 0)
  }));

  const complaintEntries = accounts.map((account) => ({
    label: `${account.email || "No mask"} / Recorded complaints`,
    count:
      account.managerSections.find((item) => item.title === "Complaints Trackers")?.complaintData?.total || 0
  }));
  const complaintRecords = accounts.flatMap((account) =>
    (account.managerSections.find((item) => item.title === "Complaints Trackers")?.complaintData?.records || []).map((record) => ({
      ...record,
      accountMask: account.email || "No mask",
      accountName: account.owner
    }))
  );

  const jiraEntries = [
    { label: "Open CX visibility requests", key: "Open CX visibility requests" },
    { label: "In progress with CX", key: "In progress with CX" },
    { label: "Resolved for AM follow-up", key: "Resolved for AM follow-up" }
  ].map((item) => ({
    label: item.label,
    count: accounts.reduce((sum, account) => {
      const section = account.managerSections.find((entry) => entry.title === "JIRA tickets");
      const match = section?.jiraData?.entries?.find((entry) => entry.label === item.key);
      return sum + (match?.count || 0);
    }, 0)
  }));

  const onboardingDraft = getPrivateOnboardingDraft();

  const privateJiraRequests = state.privateWorkspaceJira[state.selectedLogin] || [];
  const combinedRequests = [
    ...accounts.flatMap(
      (account) => account.managerSections.find((entry) => entry.title === "JIRA tickets")?.jiraData?.requests?.map((request) => ({
        ...request,
        account: account.owner
      })) || []
    ),
    ...privateJiraRequests.map((request) => ({
      ...request,
      account: "All accounts"
    }))
  ];

  const combinedNotes = [
    ...accounts.flatMap((account) =>
      (account.notes || []).map((note) => ({
        ...note,
        scope: account.owner
      }))
    ),
    ...((state.privateWorkspaceNotes[state.selectedLogin] || []).map((note) => ({
      ...note,
      scope: "All accounts"
    })))
  ];

  return {
    id: "all-accounts",
    owner: "All accounts",
    email: `${accounts.length} private accounts`,
    team: getCurrentLogin()?.name || "Private workspace",
    region: "NA",
    status: "active",
    plan: "Portfolio",
    summary: "Combined private workspace view across every account available to this login.",
    mrr: accounts.reduce((sum, account) => sum + account.mrr, 0),
    loginHealth:
      accounts.length > 0
        ? Math.round(accounts.reduce((sum, account) => sum + account.loginHealth, 0) / accounts.length)
        : 0,
    spend: {
      currentMonth: accounts.reduce((sum, account) => sum + account.spend.currentMonth, 0),
      projected: accounts.reduce((sum, account) => sum + account.spend.projected, 0),
      previousMonth: accounts.reduce((sum, account) => sum + account.spend.previousMonth, 0),
      monthly
    },
    notes: combinedNotes,
    security: { mfa: false, sso: true, audit: false },
    managerSections: [
      {
        title: "Sample Tracker",
        detail: "Combined sample activity across all private accounts for this login.",
        sampleData: {
          loggedOnline: loggedOnlineTotal,
          received: receivedTotal,
          statuses: sampleStatuses
        }
      },
      {
        title: "Call Log",
        detail: "Combined call totals across all private account masks.",
        callLogData: {
          total: callEntries.reduce((sum, entry) => sum + entry.count, 0),
          entries: callEntries
        }
      },
      {
        title: "Complaints Trackers",
        detail: "Recorded complaints across all private account masks.",
        complaintData: {
          total: complaintEntries.reduce((sum, entry) => sum + entry.count, 0),
          entries: complaintEntries,
          records: complaintRecords
        }
      },
      {
        title: "JIRA tickets",
        detail: "All AM-to-CX visibility requests across the private account set.",
        jiraData: {
          total: jiraEntries.reduce((sum, entry) => sum + entry.count, 0),
          entries: jiraEntries,
          requests: combinedRequests
        }
      },
      {
        title: ONBOARDING_SECTION_TITLE,
        detail: "Create and manage onboarding for a brand-new account before CX returns the final account number.",
        intake: onboardingDraft
      },
      {
        title: "Personal Calendar",
        detail: "Use Microsoft access and shared calendar coordination alongside this combined account view.",
        outlookData: {
          meetings: [
            { title: "Private account portfolio review", time: "Today 9:00 AM", attendees: `${getCurrentLogin()?.name || "User"}, CX leadership` },
            { title: "Customer follow-up block", time: "Tomorrow 2:00 PM", attendees: "Account owners and support partners" }
          ],
          emails: [
            { subject: "Weekly AM portfolio digest", sender: "outlook@polarislabs.com", time: "7:42 AM" },
            { subject: "New customer follow-up requests", sender: "cx-team@polarislabs.com", time: "Yesterday" }
          ]
        }
      }
    ]
  };
}

function getVisibleAccounts() {
  if (state.viewMode === "shared") {
    return [];
  }

  return getLoginAccessibleAccounts();
}

function getFilteredVisibleAccounts() {
  const query = state.accountFilterQuery.trim().toLowerCase();
  const visibleAccounts = getVisibleAccounts();
  if (!query) {
    return visibleAccounts;
  }

  return visibleAccounts.filter((account) =>
    [account.owner, account.email, account.team].some((value) =>
      String(value || "").toLowerCase().includes(query)
    )
  );
}

function getCurrentPanelLayout() {
  const saved = state.panelLayouts[state.selectedLogin];
  return {
    order: Array.isArray(saved?.order) ? saved.order : [...DEFAULT_PANEL_ORDER],
    collapsed: saved?.collapsed && typeof saved.collapsed === "object" ? saved.collapsed : {},
    viewOption: typeof saved?.viewOption === "string" ? saved.viewOption : "comfortable"
  };
}

function saveCurrentPanelLayout(layout) {
  state.panelLayouts[state.selectedLogin] = layout;
  try {
    window.localStorage.setItem(PANEL_LAYOUT_STORAGE_KEY, JSON.stringify(state.panelLayouts));
  } catch {
    // Ignore localStorage failures in the sandbox preview.
  }
  queueDashboardStateSync();
}

function ensurePanelChrome() {
  dashboardGrid.querySelectorAll(".dashboard-panel").forEach((panel) => {
    const heading = panel.querySelector(".section-heading");
    if (!heading || heading.querySelector(".panel-controls")) {
      return;
    }

    const controls = document.createElement("div");
    controls.className = "panel-controls";
    controls.innerHTML = `
      <button class="ghost-button panel-control-button panel-handle" type="button" draggable="true" data-panel-handle="${panel.dataset.panelId}" aria-label="Move panel">Move</button>
      <button class="ghost-button panel-control-button" type="button" data-panel-collapse="${panel.dataset.panelId}" aria-label="Collapse panel">Toggle</button>
    `;
    heading.append(controls);
  });
}

function renderDashboardLayout() {
  ensurePanelChrome();
  const layout = getCurrentPanelLayout();
  viewOptionSelect.value = layout.viewOption;
  dashboardGrid.dataset.viewOption = layout.viewOption;
  const panels = Array.from(dashboardGrid.querySelectorAll(".dashboard-panel"));
  const panelMap = new Map(panels.map((panel) => [panel.dataset.panelId, panel]));
  const orderedIds = [
    ...layout.order.filter((id) => panelMap.has(id)),
    ...DEFAULT_PANEL_ORDER.filter((id) => panelMap.has(id) && !layout.order.includes(id))
  ];

  orderedIds.forEach((id) => {
    dashboardGrid.append(panelMap.get(id));
  });

  panels.forEach((panel) => {
    panel.classList.toggle("panel-collapsed", Boolean(layout.collapsed[panel.dataset.panelId]));
  });
}

function getSelectedAccount() {
  const visibleAccounts = getFilteredVisibleAccounts();
  if (state.selectedId === "all-accounts") {
    return buildAggregateAccount(visibleAccounts);
  }
  return visibleAccounts.find((account) => account.id === state.selectedId);
}

function getJiraModalAccounts() {
  const selectedAccount = getSelectedAccount();
  if (selectedAccount && selectedAccount.id !== "all-accounts") {
    return [selectedAccount];
  }

  return getLoginAccessibleAccounts();
}

function getJiraModalAccount() {
  const accounts = getJiraModalAccounts();
  if (!accounts.length) {
    return null;
  }

  if (state.jiraModalAccountId) {
    return accounts.find((account) => account.id === state.jiraModalAccountId) || accounts[0];
  }

  return accounts[0];
}

function getDefaultJiraDescription(account) {
  if (!account) {
    return "";
  }

  return `Requesting CX visibility for ${account.owner} (${maskOrFallback(account.email)}).\n\nContext:\n- Contacted customer and need CX support or routing help.\n- Please review account context, current request, and recommend the next step.\n\nRequested outcome:\n- Confirm ownership and follow-up action.\n- Surface any blockers, missing setup items, or customer-facing updates.`;
}

function getJiraTicketUrl(key, externalUrl) {
  if (externalUrl) {
    return externalUrl;
  }

  if (!key || !state.jiraBaseUrl) {
    return "";
  }

  return `${state.jiraBaseUrl.replace(/\/$/, "")}/browse/${key}`;
}

function applyNetSuiteRevenueData(payload) {
  if (!payload?.byMask) {
    return;
  }

  state.accounts.forEach((account) => {
    const revenue = payload.byMask[maskOrFallback(account.email)];
    if (!revenue) {
      return;
    }

    account.spend.currentMonth = revenue.ytdRevenue || 0;
    account.spend.projected = revenue.projectedYearEnd || 0;
    account.spend.previousMonth = revenue.priorYearYtd || 0;
    account.spend.monthly = revenue.monthly || account.spend.monthly;
  });

  state.netSuiteIntegration = {
    configured: Boolean(payload.configured),
    mode: payload.mode || "mock",
    generatedAt: payload.generatedAt || "",
    overall: payload.overall || state.netSuiteIntegration.overall
  };

  const revenueItem = state.sharedItems.find((item) => item.title === "Revenue");
  if (revenueItem) {
    revenueItem.detail = `${formatCurrency(state.netSuiteIntegration.overall.ytdRevenue)} total YTD revenue from the NetSuite mock API.`;
  }

  const invoiceItem = state.sharedItems.find((item) => item.title === "Outstanding invoices");
  if (invoiceItem) {
    invoiceItem.detail = `${formatCurrency(state.netSuiteIntegration.overall.outstandingInvoices)} outstanding from the NetSuite mock API.`;
  }
}

async function loadIntegrationConfig() {
  try {
    const [jiraResponse, microsoftResponse, netSuiteRevenueResponse, genesysResponse] = await Promise.all([
      fetch("/api/jira/config"),
      fetch("/api/microsoft/config"),
      fetch("/api/netsuite/revenue"),
      fetch("/api/genesys/config")
    ]);

    if (jiraResponse.ok) {
      const jiraConfig = await jiraResponse.json();
      state.jiraBaseUrl = jiraConfig.baseUrl || "";
    }

    if (microsoftResponse.ok) {
      const microsoftConfig = await microsoftResponse.json();
      state.microsoftIntegration = {
        configured: Boolean(microsoftConfig.configured),
        connected: Boolean(microsoftConfig.connected),
        email: microsoftConfig.email || "",
        displayName: microsoftConfig.displayName || "",
        connectedAt: microsoftConfig.connectedAt || "",
        scopes: microsoftConfig.scopes || [],
        teamsUrl: microsoftConfig.teamsUrl || "https://teams.microsoft.com",
        officeUrl: microsoftConfig.officeUrl || "https://www.office.com",
        authUrl: microsoftConfig.authUrl || "/auth/microsoft/start"
      };
    }

    if (netSuiteRevenueResponse.ok) {
      const netSuiteRevenue = await netSuiteRevenueResponse.json();
      applyNetSuiteRevenueData(netSuiteRevenue);
    }

    if (genesysResponse.ok) {
      const genesysConfig = await genesysResponse.json();
      state.genesysIntegration = {
        configured: Boolean(genesysConfig.configured),
        mode: genesysConfig.mode || "mock",
        region: genesysConfig.region || "",
        orgName: genesysConfig.orgName || "",
        generatedAt: genesysConfig.generatedAt || "",
        portalUrl: genesysConfig.portalUrl || "",
        developerUrl: genesysConfig.developerUrl || "https://developer.genesys.cloud",
        scopes: genesysConfig.scopes || [],
        byRegion: genesysConfig.byRegion || {}
      };
    }
  } catch {
    state.jiraBaseUrl = state.jiraBaseUrl || "";
  }

  renderWorkspaceSummary();
}

function registerLocalJiraRequest(account, request) {
  const jiraSection = account?.managerSections.find((section) => section.title === "JIRA tickets");
  if (!jiraSection?.jiraData) {
    if (!state.privateWorkspaceJira[state.selectedLogin]) {
      state.privateWorkspaceJira[state.selectedLogin] = [];
    }
    state.privateWorkspaceJira[state.selectedLogin].unshift({
      ...request,
      account: account?.owner || "All accounts"
    });
    return true;
  }

  if (!jiraSection.jiraData.requests) {
    jiraSection.jiraData.requests = [];
  }

  jiraSection.jiraData.requests.unshift(request);
  jiraSection.jiraData.total += 1;
  const openEntry = jiraSection.jiraData.entries.find((entry) => entry.label === "Open CX visibility requests");
  if (openEntry) {
    openEntry.count += 1;
  }

  return true;
}

async function submitJiraRequest(account, request) {
  const response = await fetch("/api/jira/create", {
    method: "POST",
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      summary: request.title,
      description: request.description,
      priority: request.priority,
      accountMask: request.account,
      company: account.owner,
      contactName: request.contactName,
      contactEmail: request.contactEmail,
      cc: request.cc,
      attachments: request.attachments || []
    })
  });

  const result = await response.json();
  if (!response.ok) {
    return {
      ok: false,
      configured: result.code !== "jira_not_configured",
      message: result.message || "Jira create failed"
    };
  }

  return {
    ok: true,
    configured: true,
    key: result.key,
    browseUrl: result.browseUrl
  };
}

async function createHorizonAccessTicket(account, intake) {
  if (!account) {
    return false;
  }

  const users = intake.horizonTrainingUsers?.trim() || "Training attendees to be confirmed";
  const schedule = intake.horizonTrainingSchedule?.trim() || "Schedule pending";
  const request = {
    id: createLocalId("jira"),
    title: `Provision Horizon access for ${maskOrFallback(account.email)}`,
    priority: "High",
    status: "Open",
    createdAt: new Date().toISOString(),
    owner: state.selectedLogin,
    account: maskOrFallback(account.email),
    contactName: account.contactName,
    contactEmail: account.contactEmail,
    cc: [account.contactEmail, getCurrentLogin()?.microsoftEmail].filter(Boolean).join(", "),
    description: `AI generated Horizon access request.\n\nAccount: ${account.owner} (${maskOrFallback(account.email)})\nTraining schedule: ${schedule}\nUsers attending: ${users}\n\nRequested outcome:\n- Create Horizon access for listed users.\n- Confirm invites or onboarding steps before training.`,
    attachments: []
  };

  const remote = await submitJiraRequest(account, request);
  if (remote.ok) {
    request.status = "Submitted";
    request.externalKey = remote.key;
    request.externalUrl = remote.browseUrl;
    request.title = `${request.title} (${remote.key})`;
  }
  registerLocalJiraRequest(account, request);

  intake.horizonTrainingStatus = `AI queued Horizon access request on ${new Date().toLocaleDateString("en-US")}`;
  return remote.ok ? { created: true, remote } : { created: true, remote };
}

async function createAccountCreationTicket(account, intake) {
  const request = {
    id: createLocalId("jira"),
    title: `Create account structure for ${account.owner}`,
    priority: "High",
    status: "Open",
    createdAt: new Date().toISOString(),
    owner: state.selectedLogin,
    account: maskOrFallback(account.email),
    contactName: account.contactName,
    contactEmail: account.contactEmail,
    cc: [account.contactEmail, getCurrentLogin()?.microsoftEmail].filter(Boolean).join(", "),
    description: [
      `Create the CX account structure for ${account.owner} (${maskOrFallback(account.email)}).`,
      "",
      `Primary contact: ${account.contactName || "Not set"} (${account.contactEmail || "Not set"})`,
      `Projected volume: ${intake.projectedVolume || "Not set"}`,
      `Quote number: ${intake.quoteNumber || "Not set"}`,
      `Account structure set up: ${intake.accountStructureSetup || "Not set"}`,
      `Billing info: ${intake.billingInfo || "Not set"}`
    ].join("\n"),
    attachments: []
  };

  const remote = await submitJiraRequest(account, request);
  if (remote.ok) {
    request.status = "Submitted";
    request.externalKey = remote.key;
    request.externalUrl = remote.browseUrl;
    request.title = `${request.title} (${remote.key})`;
  }
  registerLocalJiraRequest(account, request);

  intake.accountCreationTicketAt = new Date().toLocaleDateString("en-US");
  intake.accountCreationTicketKey = remote.key || "";
  return { created: true, remote };
}

async function createEquipmentUploadTicket(account, intake) {
  const request = {
    id: createLocalId("jira"),
    title: `Upload equipment for ${account.owner}`,
    priority: "Medium",
    status: "Open",
    createdAt: new Date().toISOString(),
    owner: state.selectedLogin,
    account: maskOrFallback(account.email),
    contactName: account.contactName,
    contactEmail: account.contactEmail,
    cc: [account.contactEmail, getCurrentLogin()?.microsoftEmail].filter(Boolean).join(", "),
    description: [
      `Upload equipment for ${account.owner} (${maskOrFallback(account.email)}).`,
      "",
      `Primary contact: ${account.contactName || "Not set"} (${account.contactEmail || "Not set"})`,
      `Equipment list: ${intake.equipmentList || "Not set"}`
    ].join("\n"),
    attachments: []
  };

  const remote = await submitJiraRequest(account, request);
  if (remote.ok) {
    request.status = "Submitted";
    request.externalKey = remote.key;
    request.externalUrl = remote.browseUrl;
    request.title = `${request.title} (${remote.key})`;
  }
  registerLocalJiraRequest(account, request);

  intake.equipmentUploadTicketAt = new Date().toLocaleDateString("en-US");
  intake.equipmentUploadTicketKey = remote.key || "";
  return { created: true, remote };
}

function prepareNetSuiteHandoff(account, intake) {
  intake.accountingHandoffSummary = [
    `AI accounting handoff for ${account.owner} (${maskOrFallback(account.email)})`,
    `Primary contact: ${account.contactName || "Not set"} (${account.contactEmail || "Not set"})`,
    `Projected volume: ${intake.projectedVolume || "Not set"}`,
    `Billing info: ${intake.billingInfo || "Not set"}`,
    `Account structure: ${intake.accountStructureSetup || "Not set"}`,
    `Accounting requirements: ${intake.accountingRequirements || "Accounting to confirm required setup details."}`
  ].join("\n");
  intake.netSuiteHandoffPreparedAt = new Date().toLocaleDateString("en-US");
  return true;
}

function getMissingOnboardingFields(intake) {
  const requiredFields = [
    ["onboardingCompanyName", "Company name"],
    ["onboardingContactName", "Contact person"],
    ["onboardingContactEmail", "Contact email"],
    ["projectedVolume", "Projected Volume"],
    ["quoteNumber", "Quote number"],
    ["accountStructureSetup", "Account structure set up"],
    ["billingInfo", "Billing Info"],
    ["accountingRequirements", "Accounting requirements"],
    ["equipmentList", "Equipment list"],
    ["horizonTrainingSchedule", "Schedule for Horizon Training"],
    ["horizonTrainingUsers", "Users attending Horizon Training"],
    ["pricingReviewPlan", "Revisit Pricing at 90 days"]
  ];

  return requiredFields
    .filter(([key]) => !intake?.[key]?.trim())
    .map(([, label]) => label);
}

async function createOnboardingTicket(account, intake) {
  if (!account) {
    return { created: false, missingFields: [] };
  }

  const missingFields = getMissingOnboardingFields(intake);
  if (missingFields.length) {
    return { created: false, missingFields };
  }

  const request = {
    id: createLocalId("jira"),
    title: `Start onboarding for ${account.owner}`,
    priority: "High",
    status: "Open",
    createdAt: new Date().toISOString(),
    owner: state.selectedLogin,
    account: maskOrFallback(account.email),
    contactName: account.contactName,
    contactEmail: account.contactEmail,
    cc: [account.contactEmail, getCurrentLogin()?.microsoftEmail].filter(Boolean).join(", "),
    description: [
      `Onboarding intake ready for ${account.owner} (${maskOrFallback(account.email)}).`,
      "",
      `Primary contact: ${account.contactName || "Not set"} (${account.contactEmail || "Not set"})`,
      `Onboarding start date: ${intake.onboardingStartDate || getTodayIsoDate()}`,
      `Projected Volume: ${intake.projectedVolume}`,
      `Quote number: ${intake.quoteNumber}`,
      `Account structure set up: ${intake.accountStructureSetup}`,
      `Billing Info: ${intake.billingInfo}`,
      `Accounting requirements: ${intake.accountingRequirements || "Pending accounting guidance"}`,
      `Accounting handoff prepared: ${intake.netSuiteHandoffPreparedAt || "Not yet"}`,
      `CX account creation ticket: ${intake.accountCreationTicketKey || intake.accountCreationTicketAt || "Not yet"}`,
      `Equipment list: ${intake.equipmentList}`,
      `Equipment upload Jira: ${intake.equipmentUploadTicketKey || intake.equipmentUploadTicketAt || "Not yet"}`,
      `Horizon Training status: ${intake.horizonTrainingStatus || "Not set"}`,
      `Schedule for Horizon Training: ${intake.horizonTrainingSchedule}`,
      `Users attending Horizon Training: ${intake.horizonTrainingUsers}`,
      `Weekly check-in notes: ${intake.weeklyCheckInNotes || "Not set"}`,
      `45-day review notes: ${intake.day45ReviewNotes || "Not set"}`,
      `70-day review notes: ${intake.day70ReviewNotes || "Not set"}`,
      `Revisit Pricing at 90 days: ${intake.pricingReviewPlan}`
    ].join("\n"),
    attachments: []
  };

  const remote = await submitJiraRequest(account, request);
  if (remote.ok) {
    request.status = "Submitted";
    request.externalKey = remote.key;
    request.externalUrl = remote.browseUrl;
    request.title = `${request.title} (${remote.key})`;
  }
  registerLocalJiraRequest(account, request);

  intake.onboardingStartedAt = new Date().toLocaleDateString("en-US");
  return { created: true, missingFields: [], remote };
}

function renderJiraModalAttachmentList() {
  const files = Array.from(jiraModalAttachments?.files || []);
  if (!jiraModalAttachmentList) {
    return;
  }

  if (!files.length) {
    jiraModalAttachmentList.innerHTML = `<div class="empty-state"><p>No attachments selected yet.</p></div>`;
    return;
  }

  jiraModalAttachmentList.innerHTML = files
    .map(
      (file) => `
        <div class="shared-item">
          <strong>${file.name}</strong>
          <span>${Math.max(1, Math.round(file.size / 1024))} KB</span>
        </div>
      `
    )
    .join("");
}

function syncJiraModalFields() {
  const accounts = getJiraModalAccounts();
  if (!jiraModalAccountSelect) {
    return;
  }

  if (!accounts.length) {
    jiraModalAccountSelect.innerHTML = "";
    jiraModalCompany.value = "";
    jiraModalMask.value = "";
    jiraModalContactName.value = "";
    jiraModalContactEmail.value = "";
    jiraModalCc.value = "";
    jiraModalDescription.value = "";
    return;
  }

  if (!state.jiraModalAccountId || !accounts.some((account) => account.id === state.jiraModalAccountId)) {
    state.jiraModalAccountId = accounts[0].id;
  }

  jiraModalAccountSelect.innerHTML = accounts
    .map(
      (account) => `<option value="${account.id}">${account.owner} (${maskOrFallback(account.email)})</option>`
    )
    .join("");

  jiraModalAccountSelect.value = state.jiraModalAccountId;
  jiraModalAccountSelect.disabled = accounts.length === 1;

  const account = getJiraModalAccount();
  const login = getCurrentLogin();
  if (!account) {
    return;
  }

  jiraModalCompany.value = account.owner || "";
  jiraModalMask.value = maskOrFallback(account.email);
  jiraModalContactName.value = account.contactName || "";
  jiraModalContactEmail.value = account.contactEmail || "";
  jiraModalCc.value = [account.contactEmail, login?.microsoftEmail].filter(Boolean).join(", ");
  jiraModalDescription.value = getDefaultJiraDescription(account);
}

function openJiraModal() {
  const accounts = getJiraModalAccounts();
  if (!accounts.length) {
    showToast("No account is available for this login.");
    return;
  }

  state.jiraModalAccountId = accounts[0].id;
  jiraModalForm.reset();
  jiraModalPriority.value = "Medium";
  syncJiraModalFields();
  renderJiraModalAttachmentList();
  jiraModal.classList.remove("hidden");
  jiraModal.setAttribute("aria-hidden", "false");
  jiraModalTitleInput.focus();
}

function closeJiraModal() {
  jiraModal.classList.add("hidden");
  jiraModal.setAttribute("aria-hidden", "true");
  state.jiraModalAccountId = null;
}

function formatCurrency(value) {
  return new Intl.NumberFormat("en-US", {
    style: "currency",
    currency: "USD",
    maximumFractionDigits: 0
  }).format(value);
}

function createLocalId(prefix) {
  return `${prefix}-${Date.now()}-${Math.floor(Math.random() * 1000)}`;
}

function getImportedJiraBucket(account) {
  if (!account) {
    return null;
  }

  if (account.id === "all-accounts") {
    const visibleBuckets = getVisibleSimulatedJiraAccounts();
    return {
      accountId: "all-accounts",
      mask: "All private accounts",
      total: visibleBuckets.reduce((sum, bucket) => sum + bucket.total, 0),
      open: visibleBuckets.reduce((sum, bucket) => sum + bucket.open, 0),
      inProgress: visibleBuckets.reduce((sum, bucket) => sum + bucket.inProgress, 0),
      resolved: visibleBuckets.reduce((sum, bucket) => sum + bucket.resolved, 0),
      componentTypes: visibleBuckets.reduce((map, bucket) => {
        Object.entries(bucket.componentTypes || {}).forEach(([name, count]) => {
          map[name] = (map[name] || 0) + count;
        });
        return map;
      }, {}),
      recent: visibleBuckets.flatMap((bucket) => bucket.recent || []),
      items: visibleBuckets.flatMap((bucket) => bucket.items || [])
    };
  }

  return getSimulatedJiraBucket(account);
}

function getVisibleImportedJiraAccounts() {
  return getVisibleSimulatedJiraAccounts();
}

function buildAccountSummaryRows(accounts) {
  return accounts.map((account) => {
    const sampleSection = account.managerSections.find((section) => section.title === "Sample Tracker");
    const complaintSection = account.managerSections.find((section) => section.title === "Complaints Trackers");
    const importedJira = getSimulatedJiraBucket(account) || {
      total: 0,
      open: 0,
      inProgress: 0,
      resolved: 0
    };

    return {
      accountId: account.id,
      mask: account.email || "No mask",
      company: account.owner || "Untitled account",
      owner: account.team || "Unassigned",
      ytdRevenue: account.spend.currentMonth || 0,
      projectedYearEnd: account.spend.projected || 0,
      deltaPriorYear: (account.spend.currentMonth || 0) - (account.spend.previousMonth || 0),
      samplesReceived: sampleSection?.sampleData?.received || 0,
      samplesLoggedOnline: sampleSection?.sampleData?.loggedOnline || 0,
      complaints: complaintSection?.complaintData?.total || 0,
      jira30d: importedJira.total || 0,
      jiraOpen: importedJira.open || 0,
      jiraInProgress: importedJira.inProgress || 0,
      jiraResolved: importedJira.resolved || 0
    };
  });
}

function buildSummaryBreakdown(accountId, metric) {
  const account = getAccountById(accountId);
  if (!account) {
    return null;
  }

  if (metric === "ytdRevenue") {
    const quarterly = account.spend?.monthly || [];
    return {
      accountId,
      metric,
      sectionTitle: "Sample Tracker",
      title: `${account.owner} revenue breakdown`,
      subtitle: `${maskOrFallback(account.email)} · ${formatCurrency(account.spend.currentMonth || 0)} year to date`,
      items: [
        ...quarterly.map((item) => ({
          label: item.label,
          detail: "Quarterly contribution",
          value: formatCurrency(item.value || 0)
        })),
        {
          label: "Projected year end",
          detail: "Current projection",
          value: formatCurrency(account.spend.projected || 0)
        },
        {
          label: "Delta vs prior year",
          detail: "Current year compared with prior period",
          value: `${(account.spend.currentMonth || 0) - (account.spend.previousMonth || 0) >= 0 ? "+" : "-"}${formatCurrency(Math.abs((account.spend.currentMonth || 0) - (account.spend.previousMonth || 0)))}`
        }
      ]
    };
  }

  if (metric === "samplesReceived") {
    const sampleSection = account.managerSections.find((section) => section.title === "Sample Tracker");
    const statuses = sampleSection?.sampleData?.statuses || [];
    return {
      accountId,
      metric,
      sectionTitle: "Sample Tracker",
      title: `${account.owner} sample breakdown`,
      subtitle: `${maskOrFallback(account.email)} · ${sampleSection?.sampleData?.received || 0} received / ${sampleSection?.sampleData?.loggedOnline || 0} logged online`,
      items: [
        {
          label: "Samples received",
          detail: "Received by the lab",
          value: sampleSection?.sampleData?.received || 0
        },
        {
          label: "Logged online",
          detail: "Submitted through the online flow",
          value: sampleSection?.sampleData?.loggedOnline || 0
        },
        ...statuses.map((item) => ({
          label: item.label,
          detail: "Current sample status",
          value: item.count
        }))
      ]
    };
  }

  if (metric === "complaints") {
    const complaintSection = account.managerSections.find((section) => section.title === "Complaints Trackers");
    const records = complaintSection?.complaintData?.records || [];
    return {
      accountId,
      metric,
      sectionTitle: "Complaints Trackers",
      title: `${account.owner} complaint breakdown`,
      subtitle: `${maskOrFallback(account.email)} · ${records.length} listed complaints`,
      items: records.map((item) => ({
        label: item.title,
        detail: [item.detail, item.status, item.owner].filter(Boolean).join(" · "),
        value: item.date || ""
      }))
    };
  }

  if (metric === "jira30d" || metric === "jiraOpen") {
    const bucket = getSimulatedJiraBucket(account);
    const importedItems = bucket?.items || [];
    const filteredItems =
      metric === "jiraOpen"
        ? importedItems.filter((item) => String(item.status || "").toLowerCase() === "open")
        : importedItems;

    return {
      accountId,
      metric,
      sectionTitle: "JIRA tickets",
      title: `${account.owner} Jira breakdown`,
      subtitle: `${maskOrFallback(account.email)} · ${filteredItems.length} ${metric === "jiraOpen" ? "open" : "last 30 days"} Jira tickets`,
      items: filteredItems.map((item) => ({
        label: item.key || "Jira item",
        detail: [item.summary, item.status, item.componentName].filter(Boolean).join(" · "),
        value: item.createdLabel || item.priority || ""
      }))
    };
  }

  return null;
}

function renderSummaryBreakdown() {
  summaryBreakdown.classList.add("hidden");
  summaryBreakdown.innerHTML = "";
}

function getSelectedSummaryBreakdownData() {
  const selected = state.selectedSummaryBreakdown;
  if (state.viewMode === "shared" || !selected) {
    return null;
  }

  const breakdown = buildSummaryBreakdown(selected.accountId, selected.metric);
  if (!breakdown) {
    return null;
  }

  if (selected.metric === "complaints") {
    const complaintSection = getAccountById(selected.accountId)?.managerSections.find(
      (section) => section.title === "Complaints Trackers"
    );
    const records = complaintSection?.complaintData?.records || [];

    breakdown.items = records.map((item) => ({
      label: item.key || item.title || "Complaint",
      detail: [
        item.issueType ? `Type: ${item.issueType}` : "",
        buildAiComplaintDescription(item),
        item.status ? `Status: ${item.status}` : ""
      ]
        .filter(Boolean)
        .join(" - "),
      value: item.date || "",
      linkUrl: getJiraTicketUrl(item.key, item.externalUrl)
    }));
  }

  if (selected.metric === "jira30d" || selected.metric === "jiraOpen") {
    const account = state.accounts.find((item) => item.id === selected.accountId);
    const bucket = getSimulatedJiraBucket(account);
    const importedItems = bucket?.items || [];
    const filteredItems =
      selected.metric === "jiraOpen"
        ? importedItems.filter((item) => String(item.status || "").toLowerCase() === "open")
        : importedItems;

    breakdown.items = filteredItems.map((item) => ({
      label: item.key || "CS item",
      detail: [
        buildAiJiraDescription(item),
        item.componentName ? `Component: ${item.componentName}` : "",
        item.status ? `Status: ${item.status}` : ""
      ]
        .filter(Boolean)
        .join(" - "),
      value: item.createdLabel || item.priority || "",
      accountId: selected.accountId,
      sectionTitle: "JIRA tickets"
    }));
  }

  return breakdown;
}

function renderSummaryBreakdownMarkup(breakdown) {
  return `
    <div class="summary-breakdown inline-breakdown">
      <div class="summary-breakdown-head">
        <div>
          <strong>${breakdown.title}</strong>
          <p>${breakdown.subtitle}</p>
        </div>
        <div class="summary-breakdown-actions">
          <button type="button" class="secondary-button compact-button" data-summary-open-workspace="${breakdown.accountId}" data-summary-section-title="${breakdown.sectionTitle}">
            Open in workspace
          </button>
          <button type="button" class="secondary-button compact-button" data-summary-close>
            Close
          </button>
        </div>
      </div>
      <div class="summary-breakdown-list">
        ${
          breakdown.items.length
            ? breakdown.items
                .map(
                  (item) => `
                    <div class="summary-breakdown-item">
                      <div>
                        <strong>${
                          item.accountId && item.sectionTitle
                            ? `<button type="button" class="summary-link-button jira-inline-link" data-workspace-account-link="${item.accountId}" data-workspace-section-link="${item.sectionTitle}">${item.label}</button>`
                            : item.linkUrl
                              ? `<a class="jira-ticket-link" href="${item.linkUrl}" target="_blank" rel="noreferrer">${item.label}</a>`
                            : item.label
                        }</strong>
                        ${item.detail ? `<span>${item.detail}</span>` : ""}
                      </div>
                      <strong>${item.value}</strong>
                    </div>
                  `
                )
                .join("")
            : `
              <div class="empty-state">
                <p>No detail rows are available for this summary yet.</p>
              </div>
            `
        }
      </div>
    </div>
  `;
}

function renderSummaryTable() {
  const isShared = state.viewMode === "shared";
  summaryTablePanel.classList.toggle("hidden", isShared);

  if (isShared) {
    renderSummaryBreakdown();
    return;
  }

  const selectedAccount = getSelectedAccount();
  const accountsForSummary =
    selectedAccount && selectedAccount.id !== "all-accounts" ? [selectedAccount] : getLoginAccessibleAccounts();
  const rows = buildAccountSummaryRows(accountsForSummary);
  const selectedBreakdown = getSelectedSummaryBreakdownData();
  summaryTableMeta.textContent =
    selectedAccount && selectedAccount.id !== "all-accounts"
      ? `${selectedAccount.owner} summary for ${state.selectedLogin}.`
      : `${rows.length} account${rows.length === 1 ? "" : "s"} summarized for ${state.selectedLogin}.`;

  if (rows.length === 0) {
    summaryTableBody.innerHTML = `
      <tr>
        <td colspan="8">No summary rows available for this login yet.</td>
      </tr>
    `;
    return;
  }

  summaryTableBody.innerHTML = rows
    .sort((a, b) => b.ytdRevenue - a.ytdRevenue)
    .map(
      (row) => `
        <tr>
          <td>${row.mask}</td>
          <td>${row.company}</td>
          <td>${row.owner}</td>
          <td>
            <button type="button" class="summary-link-button" data-summary-section-link="Sample Tracker" data-summary-account-link="${row.accountId}" data-summary-metric="ytdRevenue">
              ${formatCurrency(row.ytdRevenue)}
            </button>
          </td>
          <td>
            <button type="button" class="summary-link-button" data-summary-section-link="Sample Tracker" data-summary-account-link="${row.accountId}" data-summary-metric="samplesReceived">
              ${row.samplesReceived}
            </button>
          </td>
          <td>
            <button type="button" class="summary-link-button" data-summary-section-link="Complaints Trackers" data-summary-account-link="${row.accountId}" data-summary-metric="complaints">
              ${row.complaints}
            </button>
          </td>
          <td>
            <button type="button" class="summary-link-button" data-summary-section-link="JIRA tickets" data-summary-account-link="${row.accountId}" data-summary-metric="jira30d">
              ${row.jira30d}
            </button>
          </td>
          <td>
            <button type="button" class="summary-link-button" data-summary-section-link="JIRA tickets" data-summary-account-link="${row.accountId}" data-summary-metric="jiraOpen">
              ${row.jiraOpen}
            </button>
          </td>
        </tr>
        ${
          selectedBreakdown && selectedBreakdown.accountId === row.accountId
            ? `
              <tr class="summary-breakdown-row">
                <td colspan="8">${renderSummaryBreakdownMarkup(selectedBreakdown)}</td>
              </tr>
            `
            : ""
        }
      `
    )
    .join("");
  renderSummaryBreakdown();
}

function openManagerSectionFromSummary(accountId, sectionTitle) {
  const account = getAccountById(accountId);
  if (!account) {
    return;
  }

  state.viewMode = "private";
  state.selectedId = accountId;
  state.selectedImportedJiraAccountId = accountId;
  state.selectedImportedJiraComponent = "All";
  state.selectedSummaryBreakdown = null;
  render();
  const targetCard = managerWorkspace.querySelector(`[data-manager-section="${sectionTitle}"]`);
  if (targetCard) {
    targetCard.scrollIntoView({ behavior: "smooth", block: "start" });
    return;
  }
  managerWorkspacePanel.scrollIntoView({ behavior: "smooth", block: "start" });
}

function renderImportedJiraItems(items, emptyMessage) {
  if (!items || items.length === 0) {
    return `
      <div class="empty-state">
        <strong>No Jira activity</strong>
        <p>${emptyMessage}</p>
      </div>
    `;
  }

  return items
    .map(
      (ticket) => `
        <div class="jira-request-item imported-item">
          <strong>${ticket.key} · ${ticket.summary}</strong>
          <span>${ticket.matchedMask || "No mask"} · ${ticket.priority} priority · ${ticket.status} · ${ticket.createdLabel}</span>
        </div>
      `
    )
    .join("");
}

function renderImportedJiraMaskSummary(account) {
  if (account?.id !== "all-accounts") {
    return "";
  }

  const maskCards = getVisibleImportedJiraAccounts()
    .sort((a, b) => b.total - a.total)
    .map(
      (bucket) => `
        <button class="manager-list-item jira-mask-button ${state.selectedImportedJiraAccountId === bucket.accountId ? "active" : ""}" type="button" data-jira-mask-account="${bucket.accountId}">
          <span>${bucket.mask}</span>
          <strong>${bucket.total}</strong>
        </button>
      `
    )
    .join("");

  if (!maskCards) {
    return "";
  }

  return `
    <div class="jira-mask-summary">
      <div class="outlook-heading">Tickets by account mask</div>
      <div class="manager-list">${maskCards}</div>
    </div>
  `;
}

function renderImportedJiraTicketList(items, emptyMessage, accountId) {
  if (!items || items.length === 0) {
    return `
      <div class="empty-state">
        <strong>No Jira activity</strong>
        <p>${emptyMessage}</p>
      </div>
    `;
  }

  return items
    .map(
      (ticket) => `
        <div class="jira-request-item imported-item">
          <strong>${
            accountId
              ? `<button type="button" class="summary-link-button jira-inline-link" data-workspace-account-link="${accountId}" data-workspace-section-link="JIRA tickets">${ticket.key}</button>`
              : ticket.key
          } - ${ticket.summary}</strong>
          <span>${ticket.matchedMask || "No mask"} - ${ticket.priority} priority - ${ticket.status} - ${ticket.createdLabel}</span>
          <span class="jira-age-meta">${formatJiraAgeLabel(ticket)}</span>
        </div>
      `
    )
    .join("");
}

function buildAiJiraDescription(ticket) {
  const summary = String(ticket.summary || "").trim();
  if (!summary) {
    return "Simulated Jira activity tied to this account.";
  }

  const cleaned = summary
    .replace(/^RE:\s*/i, "")
    .replace(/^FW:\s*/i, "")
    .replace(/^FWD:\s*/i, "")
    .replace(/\s+/g, " ")
    .trim();

  return `AI summary: ${cleaned}.`;
}

function renderJiraComponentChart(bucket, selectedComponent) {
  const entries = Object.entries(bucket?.componentTypes || {}).sort((a, b) => b[1] - a[1]);
  if (entries.length === 0) {
    return `
      <div class="empty-state">
        <strong>No component data</strong>
        <p>No Jira component types are available for this account right now.</p>
      </div>
    `;
  }

  const maxValue = Math.max(...entries.map(([, count]) => count), 1);
  return `
    <div class="jira-chart-list">
      ${entries
        .map(
          ([name, count]) => `
            <button class="jira-chart-row ${selectedComponent === name ? "active" : ""}" type="button" data-jira-component-button="${name}">
              <span class="jira-chart-code jira-chart-name">${name}</span>
              <div class="jira-chart-track">
                <div class="jira-chart-value" style="width: ${(count / maxValue) * 100}%"></div>
              </div>
              <strong>${count}</strong>
            </button>
          `
        )
        .join("")}
    </div>
  `;
}

function renderImportedJiraDetails(account, importedBucket) {
  if (!importedBucket) {
    return `
      <div class="empty-state">
        <strong>No Jira activity</strong>
        <p>No simulated Jira activity is available for this account.</p>
      </div>
    `;
  }

  const selectedBucket =
    account?.id === "all-accounts" && state.selectedImportedJiraAccountId
      ? getVisibleImportedJiraAccounts().find((bucket) => bucket.accountId === state.selectedImportedJiraAccountId) || null
      : importedBucket;

  if (!selectedBucket) {
    return `
      <div class="empty-state">
        <strong>Select an account mask</strong>
        <p>Choose an account above to see the Jira activity tied to that account.</p>
      </div>
    `;
  }

  const componentEntries = Object.entries(selectedBucket.componentTypes || {}).sort((a, b) => b[1] - a[1]);
  const selectedComponent = state.selectedImportedJiraComponent || "All";
  if (selectedComponent === "All") {
    return `
      <div class="jira-workspace-shell">
        <div class="jira-toolbar">
          <div class="jira-toolbar-copy">
            <strong>${selectedBucket.mask || maskOrFallback(account?.email)}</strong>
            <span>Select a component type to show the simulated CS tickets for this account.</span>
          </div>
          <label class="tool-label compact-label jira-toolbar-filter" for="jira-component-filter">
            Component type
            <select id="jira-component-filter" data-jira-component-filter>
              <option value="All" selected>Choose a component</option>
              ${componentEntries
                .map(
                  ([name, count]) => `
                    <option value="${name}">${name} (${count})</option>
                  `
                )
                .join("")}
            </select>
          </label>
        </div>
      </div>
    `;
  }

  const filteredTickets = (selectedBucket.recent || [])
    .filter((ticket) => selectedComponent === "All" || ticket.componentType === selectedComponent)
    .slice(0, 12);
  if (filteredTickets.length === 0 && selectedComponent !== "All") {
    return `
      <div class="empty-state">
        <strong>No Jira activity</strong>
        <p>No simulated CS tickets are available for this component filter.</p>
      </div>
    `;
  }

  const openCount = filteredTickets.filter((ticket) => ticket.statusGroup === "Open").length;
  const inProgressCount = filteredTickets.filter((ticket) => ticket.statusGroup === "In progress").length;
  const resolvedCount = filteredTickets.filter((ticket) => ticket.statusGroup === "Resolved").length;
  const avgTurnaround = getAverageJiraTurnaroundDays(
    filteredTickets.filter((ticket) => ticket.statusGroup === "Resolved")
  );
  const oldestOpenAge = filteredTickets
    .filter((ticket) => ticket.statusGroup !== "Resolved")
    .reduce((maxAge, ticket) => Math.max(maxAge, getJiraTicketAgeDays(ticket) ?? 0), 0);
  const importedQueue = renderImportedJiraTicketList(
    filteredTickets,
    selectedComponent === "All"
      ? "No simulated Jira tickets matched this account in the last 30 days."
      : "No simulated Jira tickets matched this component filter.",
    selectedBucket.accountId
  );

  return `
    <div class="jira-workspace-shell">
      <div class="jira-toolbar">
        <div class="jira-toolbar-copy">
          <strong>${selectedBucket.mask || maskOrFallback(account?.email)}</strong>
          <span>${filteredTickets.length} visible CS tickets in the current filter</span>
        </div>
        <label class="tool-label compact-label jira-toolbar-filter" for="jira-component-filter">
          Component type
          <select id="jira-component-filter" data-jira-component-filter>
            <option value="All">All components</option>
            ${componentEntries
              .map(
                ([name, count]) => `
                  <option value="${name}" ${selectedComponent === name ? "selected" : ""}>${name} (${count})</option>
                `
              )
              .join("")}
          </select>
        </label>
      </div>
      <div class="jira-counter-grid jira-counter-grid-compact">
        <div class="manager-stat">
          <span>Total tickets</span>
          <strong>${filteredTickets.length}</strong>
        </div>
        <div class="manager-stat">
          <span>Open</span>
          <strong>${openCount}</strong>
        </div>
        <div class="manager-stat">
          <span>In progress</span>
          <strong>${inProgressCount}</strong>
        </div>
        <div class="manager-stat">
          <span>Resolved</span>
          <strong>${resolvedCount}</strong>
        </div>
        <div class="manager-stat">
          <span>Avg turnaround</span>
          <strong>${avgTurnaround === null ? "--" : `${avgTurnaround}d`}</strong>
        </div>
        <div class="manager-stat">
          <span>Oldest open</span>
          <strong>${oldestOpenAge}d</strong>
        </div>
      </div>
      <div class="jira-chart-block jira-chart-block-slim jira-component-queue">
        <div class="jira-queue-header">
          <strong>Simulated ticket queue</strong>
          <span>Filter by component type and review the matching CS tickets below.</span>
        </div>
        <div class="jira-detail-list">
          ${importedQueue}
        </div>
      </div>
    </div>
  `;
}

function showToast(message) {
  toast.textContent = message;
  toast.classList.add("visible");
  window.clearTimeout(showToast.timeoutId);
  showToast.timeoutId = window.setTimeout(() => {
    toast.classList.remove("visible");
  }, 2200);
}

function syncSelectedAccount() {
  const visibleAccounts = getFilteredVisibleAccounts();

  if (state.viewMode === "shared") {
    state.selectedId = null;
    return;
  }

  if (visibleAccounts.length === 0) {
    state.selectedId = "all-accounts";
    return;
  }

  if (state.selectedId === "all-accounts") {
    return;
  }

  if (!visibleAccounts.some((account) => account.id === state.selectedId)) {
    state.selectedId = "all-accounts";
  }
}

function renderAccounts() {
  accountList.innerHTML = "";
  accountFilterInput.value = state.accountFilterQuery;
  const visibleAccounts = getFilteredVisibleAccounts();

  if (state.viewMode === "shared") {
    accountList.innerHTML = `
      <div class="empty-state empty-state-sidebar">
        <strong>Private accounts hidden</strong>
        <p>Shared view does not expose any owner-specific account data.</p>
      </div>
    `;
    return;
  }

  const listItems = [
    {
      type: "all",
      id: "all-accounts",
      title: "All",
      meta: "Private workspace + onboarding",
      submeta: `${visibleAccounts.length} account${visibleAccounts.length === 1 ? "" : "s"}${visibleAccounts.length === 0 ? " available yet" : ""}`,
      active: state.selectedId === "all-accounts",
      statusActive: true
    },
    ...visibleAccounts.map((account) => ({
      type: "account",
      id: account.id,
      title: account.owner || "Untitled account",
      meta: account.team || "No owner assigned",
      submeta: account.email || "No account mask added",
      active: account.id === state.selectedId,
      statusActive: account.status === "active"
    }))
  ];

  if (visibleAccounts.length === 0) {
    accountList.innerHTML = `
      <div class="account-virtual-spacer" style="height:${96}px;">
        <div class="account-virtual-window" style="transform: translateY(0);">
          <button type="button" class="account-card ${state.selectedId === "all-accounts" ? "active" : ""}" data-account-item="all-accounts">
            <div class="account-card-header">
              <div>
                <strong>All</strong>
                <p class="account-meta">Private workspace + onboarding</p>
              </div>
              <span class="status-dot active"></span>
            </div>
            <p class="account-meta">${state.accountFilterQuery ? "No matching accounts, but onboarding is still available." : "No private accounts yet. Use this space for onboarding."}</p>
          </button>
        </div>
      </div>
    `;
    return;
  }

  const itemHeight = 96;
  const viewportHeight = 460;
  const buffer = 6;
  const startIndex = Math.max(0, Math.floor(state.accountListScroll / itemHeight) - buffer);
  const visibleCount = Math.ceil(viewportHeight / itemHeight) + buffer * 2;
  const endIndex = Math.min(listItems.length, startIndex + visibleCount);
  const renderedItems = listItems.slice(startIndex, endIndex);

  accountList.innerHTML = `
    <div class="account-virtual-spacer" style="height:${listItems.length * itemHeight}px;">
      <div class="account-virtual-window" style="transform: translateY(${startIndex * itemHeight}px);">
        ${renderedItems
          .map(
            (item) => `
              <button type="button" class="account-card ${item.active ? "active" : ""}" data-account-item="${item.id}">
                <div class="account-card-header">
                  <div>
                    <strong>${item.title}</strong>
                    <p class="account-meta">${item.meta}</p>
                  </div>
                  <span class="status-dot ${item.statusActive ? "active" : ""}"></span>
                </div>
                <p class="account-meta">${item.submeta}</p>
              </button>
            `
          )
          .join("")}
      </div>
    </div>
  `;
}

function renderLoginOptions() {
  const currentLogin = syncSelectedLogin();
  if (!currentLogin) {
    loginSelect.innerHTML = "";
    return;
  }

  loginSelect.innerHTML = state.loginDirectory
    .map((person) => `<option value="${person.name}">${person.name} (${person.role})</option>`)
    .join("");

  loginSelect.value = state.selectedLogin;
}

function renderLoginSummary() {
  const login = syncSelectedLogin();
  if (!login) {
    loginSummary.innerHTML = `<p class="workspace-meta">No login profiles are available yet.</p>`;
    return;
  }

  loginSummary.innerHTML = `
    <strong>${login.name}</strong>
    <p class="workspace-meta">${login.role}</p>
    <p class="workspace-meta">${login.microsoftEmail}</p>
    <p class="workspace-meta">${isCreatorView() ? "Creator access with editing controls enabled." : "Shared dashboard access enabled."}</p>
  `;
}

function renderLoginOverlay() {
  if (!loginOverlay || !loginCardList) {
    return;
  }

  loginOverlay.classList.toggle("hidden", !state.loginOverlayOpen);
  loginCardList.innerHTML = state.loginDirectory
    .map(
      (person) => `
        <button class="login-card" type="button" data-login-card="${person.name}">
          <span class="avatar-chip">${person.avatar}</span>
          <span class="login-card-copy">
            <strong>${person.name}</strong>
            <span>${person.role}</span>
            <span>${person.microsoftEmail}</span>
          </span>
        </button>
      `
    )
    .join("");
}

function renderActiveUserCard() {
  const login = syncSelectedLogin();
  if (!login) {
    activeUserCard.innerHTML = `<p class="workspace-meta">No active user loaded.</p>`;
    return;
  }

  activeUserCard.innerHTML = `
    <div class="active-user-head">
      <span class="avatar-chip">${login.avatar}</span>
      <div>
        <strong>${login.name}</strong>
        <p class="workspace-meta">${login.role}</p>
      </div>
    </div>
    <p class="workspace-meta">${login.calendarConnected ? "Microsoft calendar connected" : "Microsoft calendar needs setup"}</p>
  `;
}

function getHomepageCards(login) {
  if (login.role === "Account Manager") {
    return [
      { title: "Private focus", detail: "Review assigned accounts, sample activity, customer notes, and new business readiness." },
      { title: "Today in Outlook", detail: "Prep meetings, check new emails, and build follow-up actions before customer touch points." },
      { title: "CX visibility", detail: "Create JIRA requests for anything that needs support or routing help from CX." }
    ];
  }

  if (login.role.includes("Customer Engagement")) {
    return [
      { title: "Engagement queue", detail: "Use the shared view to coordinate outreach, support requests, and follow-up timing." },
      { title: "Calendar context", detail: "See Microsoft-connected availability and activity that informs engagement planning." },
      { title: "Cross-team support", detail: "Partner with AMs through shared metrics and the private workspace sections they surface." }
    ];
  }

  if (login.role === "Salesforce Admin") {
    return [
      { title: "System visibility", detail: "Review shared metrics, workflow needs, and support requests that affect CRM operations." },
      { title: "Routing watch", detail: "Use the shared workspace for calendar coverage and operational follow-up." },
      { title: "Support intake", detail: "Monitor JIRA requests and prepare for Salesforce-related assistance." }
    ];
  }

  if (login.role === "Sales/Marketing") {
    return [
      { title: "Shared performance", detail: "Focus on shared revenue, invoice exposure, and upcoming team activity." },
      { title: "Launch coordination", detail: "Use shared notes and calendar timing to support campaigns and messaging." },
      { title: "Cross-functional context", detail: "Review what AMs and CX are surfacing without exposing private account internals." }
    ];
  }

  return [
    { title: "Creator controls", detail: "You have full sandbox editing access plus both private and shared dashboard views." },
    { title: "Presentation mode", detail: "Use the account selector to move between AM experiences and show different role-shaped views." },
    { title: "Integration-ready", detail: "The current build is ready for future JIRA and eoilreports.com data connections." }
  ];
}

function renderHomepagePanel() {
  const login = getCurrentLogin();
  const cards = getHomepageCards(login);
  homepagePanel.innerHTML = `
    <div class="homepage-header">
      <div>
        <p class="eyebrow">Homepage</p>
        <h3>${login.role} home</h3>
      </div>
    </div>
    <div class="homepage-grid">
      ${cards
        .map(
          (card) => `
            <article class="homepage-card">
              <strong>${card.title}</strong>
              <p>${card.detail}</p>
            </article>
          `
        )
        .join("")}
    </div>
  `;
}

function renderWorkspaceSummary() {
  const microsoft = state.microsoftIntegration;
  const genesys = state.genesysIntegration;
  const microsoftTester = getCurrentLogin()?.microsoftEmail || "dwilliams@polarislabs.com";
  const connectUrl = `${microsoft.authUrl}?login_hint=${encodeURIComponent(microsoftTester)}`;
  const activeGenesysRegion = state.pipelineFilterRegion === "ALL" ? getEveRegion() : state.pipelineFilterRegion;
  const genesysRegionSignal = genesys.byRegion?.[activeGenesysRegion];
  const account = getSelectedAccount();
  const workspaceCopy =
    state.viewMode === "shared"
      ? "Shared workspace for team metrics, calendar coordination, AI support, and cross-functional tools."
      : account?.id === "all-accounts"
        ? "Private All accounts workspace with combined portfolio metrics and tools."
        : `${account?.owner || "Account"} workspace with account-specific metrics, notes, search, and AI support.`;
  const jiraCopy =
    state.viewMode === "shared"
      ? "Simulated Jira activity is available across the shared dashboard."
      : account?.id === "all-accounts"
        ? `Jira activity is scoped to the accessible private account set for ${state.selectedLogin}.`
        : `Jira activity is scoped to ${account?.owner || "this account"} (${maskOrFallback(account?.email)}).`;
  workspaceSummary.innerHTML = `
    <article class="shared-card workspace-summary-card">
      <strong>Workspace</strong>
      <p>${workspaceCopy}</p>
      <p>${jiraCopy}</p>
    </article>
    <article class="shared-card workspace-summary-card">
      <strong>Microsoft Teams + Office</strong>
      <p>${microsoft.configured ? "Microsoft Graph integration scaffold is configured for Teams, Outlook calendar, and Office access." : "Microsoft Graph is not configured yet. The workspace is ready for Teams and Office integration once credentials are added."}</p>
      <p>Test user: ${microsoftTester}</p>
      <p>${microsoft.connected ? `Connected account: ${microsoft.displayName || microsoft.email}${microsoft.connectedAt ? ` on ${new Date(microsoft.connectedAt).toLocaleString("en-US")}` : ""}.` : "No Microsoft account is connected in the sandbox yet."}</p>
      <p>${microsoft.scopes.length ? `Scopes: ${microsoft.scopes.join(", ")}` : "Scopes will appear here once configured."}</p>
      <div class="workspace-action-row">
        <a class="external-link-button" href="${connectUrl}" target="_blank" rel="noreferrer">${microsoft.connected ? "Reconnect Microsoft" : "Connect Microsoft"}</a>
        <a class="external-link-button" href="${microsoft.teamsUrl}" target="_blank" rel="noreferrer">Open Teams</a>
        <a class="external-link-button" href="${microsoft.officeUrl}" target="_blank" rel="noreferrer">Open Office</a>
      </div>
    </article>
    <article class="shared-card workspace-summary-card">
      <strong>Genesys Cloud</strong>
      <p>${genesys.configured ? `${genesys.orgName || "Genesys Cloud"} is configured for queue, conversation, and regional pipeline signal reads.` : "Genesys Cloud is not configured yet. The backend scaffold is ready for read-only queue and interaction signals."}</p>
      <p>${genesysRegionSignal ? `${activeGenesysRegion} signal: ${genesysRegionSignal.hotOpportunities} hot opportunities, ${genesysRegionSignal.openCallbacks} open callbacks, ${genesysRegionSignal.abandonedInteractions} abandoned interactions.` : "Regional Genesys signal data will appear here once configured."}</p>
      <p>${genesys.scopes.length ? `Scopes: ${genesys.scopes.join(", ")}` : "Read scopes will appear here once configured."}</p>
      <div class="workspace-action-row">
        ${genesys.portalUrl ? `<a class="external-link-button" href="${genesys.portalUrl}" target="_blank" rel="noreferrer">Open Genesys</a>` : ""}
        <a class="external-link-button" href="${genesys.developerUrl}" target="_blank" rel="noreferrer">Genesys docs</a>
      </div>
    </article>
  `;
}

function getCurrentWorkspaceNotes() {
  if (state.viewMode === "shared") {
    return state.sharedWorkspaceNotes;
  }

  const account = getSelectedAccount();
  if (!account) {
    return [];
  }

  return account.notes || [];
}

function renderWorkspaceNotes() {
  workspaceNoteInput.value = state.workspaceNoteDraft;
  const notes = getCurrentWorkspaceNotes();

  if (state.viewMode === "shared") {
    workspaceNoteInput.placeholder = "Add a note to the shared workspace...";
  } else if (state.selectedId === "all-accounts") {
    workspaceNoteInput.placeholder = "Add a note to your private All accounts workspace...";
  } else {
    workspaceNoteInput.placeholder = "Add a note for this account...";
  }

  if (notes.length === 0) {
    workspaceNotes.innerHTML = `
      <div class="empty-state">
        <strong>No notes yet</strong>
        <p>Add a workspace note to track follow-up, context, or next steps.</p>
      </div>
    `;
    return;
  }

  workspaceNotes.innerHTML = notes
    .map(
      (note) => `
        <article class="shared-card note-card">
          <strong>${note.scope || "Workspace note"}</strong>
          <p>${note.text}</p>
          <span class="note-meta">${note.author}</span>
        </article>
      `
    )
    .join("");
}

function renderSharedPanel() {
  const isShared = state.viewMode === "shared";
  sharedPanel.classList.toggle("hidden", !isShared);

  if (!isShared) {
    return;
  }

  const jiraCard = `
    <article class="shared-card">
      <strong>Jira activity</strong>
      <p>${getVisibleSimulatedJiraAccounts().reduce((sum, bucket) => sum + bucket.total, 0)} simulated tickets across your accessible accounts. ${getVisibleSimulatedJiraAccounts().reduce((sum, bucket) => sum + bucket.open, 0)} open, ${getVisibleSimulatedJiraAccounts().reduce((sum, bucket) => sum + bucket.inProgress, 0)} in progress, ${getVisibleSimulatedJiraAccounts().reduce((sum, bucket) => sum + bucket.resolved, 0)} resolved.</p>
    </article>
  `;

  sharedList.innerHTML = jiraCard + state.sharedItems
    .map(
      (item) => `
        <article class="shared-card">
          <strong>${item.title}</strong>
          <p>${item.detail}</p>
        </article>
      `
    )
    .join("") + state.teamCalendar
    .map(
      (item) => `
        <article class="shared-card">
          <strong>${item.title}</strong>
          <p>${item.owner} / ${item.time}</p>
        </article>
      `
    )
    .join("");
}

function getCalendarSection(account) {
  if (state.viewMode === "shared") {
    return {
      title: "Everybody calendar",
      detail: "Shared Microsoft calendar across account managers, engagement, marketing, admin, and creator access.",
      meetings: state.teamCalendar.map((item) => ({
        title: item.title,
        time: item.time,
        attendees: item.owner
      })),
      emails: [
        { subject: "Shared revenue review prep", sender: "shared-workspace@polarislabs.com", time: "7:10 AM" },
        { subject: "Team calendar changes detected", sender: "outlook@polarislabs.com", time: "Yesterday" }
      ]
    };
  }

  return account?.managerSections?.find((section) => section.title === "Personal Calendar")?.outlookData
    ? {
        title: "Personal calendar",
        detail: "Outlook-style meetings and new emails tied to the current private view.",
        meetings: account.managerSections.find((section) => section.title === "Personal Calendar").outlookData.meetings,
        emails: account.managerSections.find((section) => section.title === "Personal Calendar").outlookData.emails
      }
    : null;
}

function renderManagerWorkspace(account) {
  const isShared = state.viewMode === "shared";
  managerWorkspacePanel.classList.toggle("hidden", isShared);

  if (isShared) {
    return;
  }

  if (!account) {
    managerWorkspace.innerHTML = `
      <div class="empty-state">
        <strong>No private workspace yet</strong>
        <p>Select an account to see Sample Tracker, Call Log, Complaints Trackers, JIRA tickets, Onboarding Space, and Personal Calendar.</p>
      </div>
    `;
    return;
  }

  managerWorkspace.innerHTML = account.managerSections
    .map((section) => {
      if (section.title === "Sample Tracker" && section.sampleData) {
        return `
          <article class="manager-card" data-manager-section="${section.title}">
            <strong>${section.title}</strong>
            <p>${section.detail}</p>
            ${section.sampleData.autoSummary ? `<div class="manager-inline-summary">${section.sampleData.autoSummary}</div>` : ""}
            <div class="manager-stats">
              <div class="manager-stat">
                <span>Samples received</span>
                <strong>${section.sampleData.received}</strong>
              </div>
              <div class="manager-stat">
                <span>Logged online</span>
                <strong>${section.sampleData.loggedOnline}</strong>
              </div>
            </div>
            <div class="manager-list">
              ${section.sampleData.statuses
                .map(
                  (status) => `
                    <div class="manager-list-item">
                      <span>${status.label}</span>
                      <strong>${status.count}</strong>
                    </div>
                  `
                )
                .join("")}
            </div>
            <div class="external-link-row">
              <a class="external-link-button" href="https://www.eoilreports.com/login" target="_blank" rel="noreferrer">
                Open HORIZON login
              </a>
              <span>Go directly to eoilreports.com to review sample detail and sign in.</span>
            </div>
          </article>
        `;
      }

      if (section.title === ACCOUNT_ONBOARDING_TITLE && section.intake) {
        const completion = getOnboardingCompletionPercent(account, section.intake);
        const postMaskItems = completion.items.filter((item) =>
          ["netSuiteHandoff", "equipmentUploadTicket", "horizonTraining", "weeklyCheckIns", "day45Review", "day70Review"].includes(item.key)
        );

        return `
          <article class="manager-card manager-card-wide" data-manager-section="${section.title}">
            <div class="manager-card-topline">
              <strong>${section.title}</strong>
              <span class="pill ${completion.percent === 100 ? "success" : "warning"}">${completion.percent}% complete</span>
            </div>
            <p>${section.detail}</p>
            <div class="manager-inline-summary">Quote ${section.intake.quoteNumber || "not set"} · Account ${maskOrFallback(account.email)} · Post-mask work now lives with this account.</div>
            <div class="checklist onboarding-checklist">
              ${postMaskItems
                .map(
                  (item) => `
                    <button type="button" class="checklist-item checklist-toggle ${item.complete ? "is-complete" : ""}" data-onboarding-checklist-toggle="${item.key}">
                      <span class="checklist-box" aria-hidden="true"></span>
                      <div>
                        <strong>${item.label}</strong>
                        <span>${item.detail}</span>
                      </div>
                    </button>
                  `
                )
                .join("")}
            </div>
            <form class="new-business-form">
              <label>
                Quote number
                <input type="text" data-newbiz-field="quoteNumber" value="${section.intake.quoteNumber || ""}" placeholder="Quote tied to this account" />
              </label>
              <label class="span-two-field">
                Billing Info
                <textarea rows="3" data-newbiz-field="billingInfo" placeholder="Billing contact, PO notes, invoice requirements, and remittance details">${section.intake.billingInfo || ""}</textarea>
              </label>
              <label class="span-two-field">
                Accounting requirements
                <textarea rows="3" data-newbiz-field="accountingRequirements" placeholder="What accounting needs to complete the NetSuite setup">${section.intake.accountingRequirements || ""}</textarea>
                <div class="newbiz-inline-actions">
                  <button type="button" class="secondary-button compact-button" data-action="prepare-netsuite-handoff">
                    AI prepare accounting handoff
                  </button>
                  <span>${section.intake.netSuiteHandoffPreparedAt ? `Prepared ${section.intake.netSuiteHandoffPreparedAt}.` : "Send the required account info to accounting for NetSuite."}</span>
                </div>
              </label>
              <label class="span-two-field">
                Equipment list
                <textarea rows="3" data-newbiz-field="equipmentList" placeholder="List units, assets, sample points, or equipment groups">${section.intake.equipmentList || ""}</textarea>
                <div class="newbiz-inline-actions">
                  <button type="button" class="secondary-button compact-button" data-action="create-equipment-upload-ticket">
                    Create equipment upload Jira
                  </button>
                  <span>${section.intake.equipmentUploadTicketAt ? `Created ${section.intake.equipmentUploadTicketAt}${section.intake.equipmentUploadTicketKey ? ` as ${section.intake.equipmentUploadTicketKey}` : ""}.` : "Create the CX request to upload equipment to the live account."}</span>
                </div>
              </label>
              <label>
                Horizon Training status
                <input type="text" data-newbiz-field="horizonTrainingStatus" value="${section.intake.horizonTrainingStatus || ""}" placeholder="Scheduled, completed, or still needed" />
              </label>
              <label>
                Schedule for Horizon Training
                <input type="text" data-newbiz-field="horizonTrainingSchedule" value="${section.intake.horizonTrainingSchedule || ""}" placeholder="Date, time, and meeting details" />
              </label>
              <label class="span-two-field">
                Users attending Horizon Training
                <textarea rows="3" data-newbiz-field="horizonTrainingUsers" placeholder="List the users who should attend and need Horizon access">${section.intake.horizonTrainingUsers || ""}</textarea>
                <div class="newbiz-inline-actions">
                  <button type="button" class="secondary-button compact-button" data-action="ai-create-horizon-access">
                    AI create Horizon access ticket
                  </button>
                  <span>Uses the live account number to prepare the Horizon access request.</span>
                </div>
              </label>
              <label class="span-two-field">
                Weekly check-in notes
                <textarea rows="3" data-newbiz-field="weeklyCheckInNotes" placeholder="Capture weekly check-in highlights and action items">${section.intake.weeklyCheckInNotes || ""}</textarea>
                <div class="newbiz-inline-actions">
                  <button type="button" class="secondary-button compact-button" data-action="mark-weekly-checkins-complete">
                    Mark weekly check-ins complete
                  </button>
                  <span>${section.intake.weeklyCheckInsCompletedAt ? `Completed ${section.intake.weeklyCheckInsCompletedAt}.` : "Use when the weekly onboarding cadence has been completed."}</span>
                </div>
              </label>
              <label class="span-two-field">
                45-day review
                <textarea rows="3" data-newbiz-field="day45ReviewNotes" placeholder="Review projected volume, sample volume, and complaints at day 45">${section.intake.day45ReviewNotes || ""}</textarea>
                <div class="newbiz-inline-actions">
                  <button type="button" class="secondary-button compact-button" data-action="mark-day45-review-complete">
                    Mark 45-day review complete
                  </button>
                  <span>${section.intake.day45ReviewCompletedAt ? `Completed ${section.intake.day45ReviewCompletedAt}.` : "Confirm the day-45 checkpoint when the review is done."}</span>
                </div>
              </label>
              <label class="span-two-field">
                70-day review
                <textarea rows="3" data-newbiz-field="day70ReviewNotes" placeholder="Check sample volume and note whether pricing needs to be revisited">${section.intake.day70ReviewNotes || ""}</textarea>
                <div class="newbiz-inline-actions">
                  <button type="button" class="secondary-button compact-button" data-action="mark-day70-review-complete">
                    Mark 70-day review complete
                  </button>
                  <span>${section.intake.day70ReviewCompletedAt ? `Completed ${section.intake.day70ReviewCompletedAt}.` : "Use when the 70-day sample and pricing check has been completed."}</span>
                </div>
              </label>
              <label class="span-two-field">
                90-day pricing direction
                <textarea rows="3" data-newbiz-field="pricingReviewPlan" placeholder="If projected volume matches the estimate, keep pricing. If not, have a conversation with the customer.">${section.intake.pricingReviewPlan || ""}</textarea>
              </label>
            </form>
          </article>
        `;
      }

      if (section.title === "Call Log" && section.callLogData) {
        return `
          <article class="manager-card" data-manager-section="${section.title}">
            <strong>${section.title}</strong>
            <p>${section.detail}</p>
            <div class="manager-stats">
              <div class="manager-stat">
                <span>Total calls</span>
                <strong>${section.callLogData.total}</strong>
              </div>
              <div class="manager-stat">
                <span>${account.id === "all-accounts" ? "Scope" : "Account mask"}</span>
                <strong>${account.id === "all-accounts" ? "All private accounts" : account.email || "No mask"}</strong>
              </div>
            </div>
            <div class="manager-list">
              ${section.callLogData.entries
                .map(
                  (entry) => `
                    <div class="manager-list-item">
                      <span>${entry.label}</span>
                      <strong>${entry.count}</strong>
                    </div>
                  `
                )
                .join("")}
            </div>
          </article>
        `;
      }

      if (section.title === "Complaints Trackers" && section.complaintData) {
        return `
          <article class="manager-card" data-manager-section="${section.title}">
            <strong>${section.title}</strong>
            <p>${section.detail}</p>
            <div class="manager-stats">
              <div class="manager-stat">
                <span>Recorded complaints</span>
                <strong>${section.complaintData.total}</strong>
              </div>
              <div class="manager-stat">
                <span>${account.id === "all-accounts" ? "Scope" : "Account mask"}</span>
                <strong>${account.id === "all-accounts" ? "All private accounts" : account.email || "No mask"}</strong>
              </div>
            </div>
            <div class="manager-list">
              ${section.complaintData.entries
                .map(
                  (entry) => `
                    <div class="manager-list-item">
                      <span>${entry.label}</span>
                      <strong>${entry.count}</strong>
                    </div>
                  `
                )
                .join("")}
            </div>
          </article>
        `;
      }

      if (section.title === "JIRA tickets" && section.jiraData) {
        const importedBucket = getImportedJiraBucket(account);
        const canCreateFromContext = state.viewMode === "private" && getJiraModalAccounts().length > 0;
        const avgTurnaround = getAverageJiraTurnaroundDays([
          ...((importedBucket?.items || []).filter((ticket) => ticket.statusGroup === "Resolved")),
          ...((section.jiraData.requests || []).filter((request) => request.status === "Resolved" || request.status === "Done"))
        ]);
        const oldestCreatedOpen = Math.max(
          0,
          ...(section.jiraData.requests || [])
            .filter((request) => request.status !== "Done" && request.status !== "Resolved")
            .map((request) => getJiraTicketAgeDays(request) ?? 0)
        );
        return `
          <article class="manager-card manager-card-wide" data-manager-section="${section.title}">
            <div class="manager-card-topline">
              <strong>${section.title}</strong>
              <button type="button" class="secondary-button compact-button" data-action="open-jira-modal" ${canCreateFromContext ? "" : "disabled"}>
                Create CX ticket
              </button>
            </div>
            <p>${section.detail}</p>
            <div class="manager-stats">
              <div class="manager-stat">
                <span>Total CX requests</span>
                <strong>${section.jiraData.total}</strong>
              </div>
              <div class="manager-stat">
                <span>Visibility scope</span>
                <strong>AM to CX</strong>
              </div>
              <div class="manager-stat">
                <span>Simulated last 30 days</span>
                <strong>${importedBucket?.total || 0}</strong>
              </div>
              <div class="manager-stat">
                <span>Avg turnaround</span>
                <strong>${avgTurnaround === null ? "--" : `${avgTurnaround}d`}</strong>
              </div>
              <div class="manager-stat">
                <span>Oldest CX request</span>
                <strong>${oldestCreatedOpen}d</strong>
              </div>
            </div>
            <div class="manager-list">
              ${section.jiraData.entries
                .map(
                  (entry) => `
                    <div class="manager-list-item">
                      <span>${entry.label}</span>
                      <strong>${entry.count}</strong>
                    </div>
                  `
                )
                .join("")}
            </div>
            <div class="jira-import-block">
              <div class="outlook-heading">Simulated Jira activity</div>
              ${renderImportedJiraMaskSummary(account)}
              ${renderImportedJiraDetails(account, importedBucket)}
            </div>
            <div class="jira-queue-block">
              <div class="jira-queue-header">
                <strong>Created CX requests</strong>
                <span>Locally created requests from this workspace.</span>
              </div>
              <div class="jira-request-list">
              ${(section.jiraData.requests || [])
                .map(
                  (request) => `
                    <div class="jira-request-item">
                      <strong>${
                        request.externalKey && getJiraTicketUrl(request.externalKey, request.externalUrl)
                          ? `<a class="jira-ticket-link" href="${getJiraTicketUrl(request.externalKey, request.externalUrl)}" target="_blank" rel="noreferrer">${request.externalKey}</a> - ${request.title}`
                          : request.title
                      }</strong>
                      <span>${request.priority} priority · ${request.status}${request.account ? ` · ${request.account}` : ""}</span>
                      <span class="jira-age-meta">${formatJiraAgeLabel(request)}</span>
                    </div>
                  `
                )
                .join("")}
              </div>
            </div>
            <div class="jira-composer jira-composer-hint">
              <span>Open the pop-out ticket form from this card to prefill the account mask, contact, CC line, description, and attachments.</span>
            </div>
          </article>
        `;
      }

      if (section.title === ONBOARDING_SECTION_TITLE) {
        return "";
      }

      return `
        <article class="manager-card" data-manager-section="${section.title}">
          <strong>${section.title}</strong>
          <p>${section.detail}</p>
        </article>
      `;
    })
    .join("");
}

function getOperationsSections(account) {
  return (account?.managerSections || []).filter((section) => section.title !== "Personal Calendar" && section.title !== ONBOARDING_SECTION_TITLE);
}

function renderOperationSectionCard(section, account) {
  if (section.title === "Sample Tracker" && section.sampleData) {
    return `
      <article class="manager-card manager-card-detail" data-manager-section="${section.title}">
        <strong>${section.title}</strong>
        <p>${section.detail}</p>
        ${section.sampleData.autoSummary ? `<div class="manager-inline-summary">${section.sampleData.autoSummary}</div>` : ""}
        <div class="manager-stats">
          <div class="manager-stat"><span>Samples received</span><strong>${section.sampleData.received}</strong></div>
          <div class="manager-stat"><span>Logged online</span><strong>${section.sampleData.loggedOnline}</strong></div>
        </div>
        <div class="manager-list">${section.sampleData.statuses.map((status) => `<div class="manager-list-item"><span>${status.label}</span><strong>${status.count}</strong></div>`).join("")}</div>
        <div class="external-link-row">
          <a class="external-link-button" href="https://www.eoilreports.com/login" target="_blank" rel="noreferrer">Open HORIZON login</a>
          <span>Go directly to eoilreports.com to review sample detail and sign in.</span>
        </div>
      </article>
    `;
  }

  if (section.title === "Call Log" && section.callLogData) {
    return `
      <article class="manager-card manager-card-detail" data-manager-section="${section.title}">
        <strong>${section.title}</strong>
        <p>${section.detail}</p>
        <div class="manager-stats">
          <div class="manager-stat"><span>Total calls</span><strong>${section.callLogData.total}</strong></div>
          <div class="manager-stat"><span>${account.id === "all-accounts" ? "Scope" : "Account mask"}</span><strong>${account.id === "all-accounts" ? "All private accounts" : account.email || "No mask"}</strong></div>
        </div>
        <div class="manager-list">${section.callLogData.entries.map((entry) => `<div class="manager-list-item"><span>${entry.label}</span><strong>${entry.count}</strong></div>`).join("")}</div>
      </article>
    `;
  }

  if (section.title === "Complaints Trackers" && section.complaintData) {
    return `
      <article class="manager-card manager-card-detail" data-manager-section="${section.title}">
        <strong>${section.title}</strong>
        <p>${section.detail}</p>
        <div class="manager-stats">
          <div class="manager-stat"><span>Recorded complaints</span><strong>${section.complaintData.total}</strong></div>
          <div class="manager-stat"><span>${account.id === "all-accounts" ? "Scope" : "Account mask"}</span><strong>${account.id === "all-accounts" ? "All private accounts" : account.email || "No mask"}</strong></div>
        </div>
        <div class="manager-list">${section.complaintData.entries.map((entry) => `<div class="manager-list-item"><span>${entry.label}</span><strong>${entry.count}</strong></div>`).join("")}</div>
      </article>
    `;
  }

  if (section.title === "JIRA tickets" && section.jiraData) {
    const importedBucket = getImportedJiraBucket(account);
    const canCreateFromContext = state.viewMode === "private" && getJiraModalAccounts().length > 0;
    const avgTurnaround = getAverageJiraTurnaroundDays([
      ...((importedBucket?.items || []).filter((ticket) => ticket.statusGroup === "Resolved")),
      ...((section.jiraData.requests || []).filter((request) => request.status === "Resolved" || request.status === "Done"))
    ]);
    const oldestCreatedOpen = Math.max(
      0,
      ...(section.jiraData.requests || [])
        .filter((request) => request.status !== "Done" && request.status !== "Resolved")
        .map((request) => getJiraTicketAgeDays(request) ?? 0)
    );
    return `
      <article class="manager-card manager-card-detail" data-manager-section="${section.title}">
        <div class="manager-card-topline">
          <strong>${section.title}</strong>
          <button type="button" class="secondary-button compact-button" data-action="open-jira-modal" ${canCreateFromContext ? "" : "disabled"}>Create CX ticket</button>
        </div>
        <p>${section.detail}</p>
        <div class="manager-stats">
          <div class="manager-stat"><span>Total CX requests</span><strong>${section.jiraData.total}</strong></div>
          <div class="manager-stat"><span>Visibility scope</span><strong>AM to CX</strong></div>
          <div class="manager-stat"><span>Simulated last 30 days</span><strong>${importedBucket?.total || 0}</strong></div>
          <div class="manager-stat"><span>Avg turnaround</span><strong>${avgTurnaround === null ? "--" : `${avgTurnaround}d`}</strong></div>
          <div class="manager-stat"><span>Oldest CX request</span><strong>${oldestCreatedOpen}d</strong></div>
        </div>
        <div class="manager-list">${section.jiraData.entries.map((entry) => `<div class="manager-list-item"><span>${entry.label}</span><strong>${entry.count}</strong></div>`).join("")}</div>
        <div class="jira-import-block"><div class="outlook-heading">Simulated Jira activity</div>${renderImportedJiraMaskSummary(account)}${renderImportedJiraDetails(account, importedBucket)}</div>
        <div class="jira-queue-block">
          <div class="jira-queue-header"><strong>Created CX requests</strong><span>Locally created requests from this workspace.</span></div>
          <div class="jira-request-list">${(section.jiraData.requests || []).map((request) => `<div class="jira-request-item"><strong>${request.externalKey && getJiraTicketUrl(request.externalKey, request.externalUrl) ? `<a class="jira-ticket-link" href="${getJiraTicketUrl(request.externalKey, request.externalUrl)}" target="_blank" rel="noreferrer">${request.externalKey}</a> - ${request.title}` : request.title}</strong><span>${request.priority} priority · ${request.status}${request.account ? ` · ${request.account}` : ""}</span><span class="jira-age-meta">${formatJiraAgeLabel(request)}</span></div>`).join("")}</div>
        </div>
      </article>
    `;
  }

  if (section.title === ACCOUNT_ONBOARDING_TITLE && section.intake) {
    const completion = getOnboardingCompletionPercent(account, section.intake);
    const postMaskItems = completion.items.filter((item) => ["netSuiteHandoff", "equipmentUploadTicket", "horizonTraining", "weeklyCheckIns", "day45Review", "day70Review"].includes(item.key));
    return `
      <article class="manager-card manager-card-detail" data-manager-section="${section.title}">
        <div class="manager-card-topline"><strong>${section.title}</strong><span class="pill ${completion.percent === 100 ? "success" : "warning"}">${completion.percent}% complete</span></div>
        <p>${section.detail}</p>
        <div class="manager-inline-summary">Quote ${section.intake.quoteNumber || "not set"} · Account ${maskOrFallback(account.email)} · Post-mask work now lives with this account.</div>
        <div class="checklist onboarding-checklist">${postMaskItems.map((item) => `<button type="button" class="checklist-item checklist-toggle ${item.complete ? "is-complete" : ""}" data-onboarding-checklist-toggle="${item.key}"><span class="checklist-box" aria-hidden="true"></span><div><strong>${item.label}</strong><span>${item.detail}</span></div></button>`).join("")}</div>
        <form class="new-business-form">
          <label><span>Quote number</span><input type="text" data-newbiz-field="quoteNumber" value="${section.intake.quoteNumber || ""}" placeholder="Quote tied to this account" /></label>
          <label class="span-two-field"><span>Billing Info</span><textarea rows="3" data-newbiz-field="billingInfo" placeholder="Billing contact, PO notes, invoice requirements, and remittance details">${section.intake.billingInfo || ""}</textarea></label>
          <label class="span-two-field"><span>Accounting requirements</span><textarea rows="3" data-newbiz-field="accountingRequirements" placeholder="What accounting needs to complete the NetSuite setup">${section.intake.accountingRequirements || ""}</textarea><div class="newbiz-inline-actions"><button type="button" class="secondary-button compact-button" data-action="prepare-netsuite-handoff">AI prepare accounting handoff</button><span>${section.intake.netSuiteHandoffPreparedAt ? `Prepared ${section.intake.netSuiteHandoffPreparedAt}.` : "Send the required account info to accounting for NetSuite."}</span></div></label>
          <label class="span-two-field"><span>Equipment list</span><textarea rows="3" data-newbiz-field="equipmentList" placeholder="List units, assets, sample points, or equipment groups">${section.intake.equipmentList || ""}</textarea><div class="newbiz-inline-actions"><button type="button" class="secondary-button compact-button" data-action="create-equipment-upload-ticket">Create equipment upload Jira</button><span>${section.intake.equipmentUploadTicketAt ? `Created ${section.intake.equipmentUploadTicketAt}${section.intake.equipmentUploadTicketKey ? ` as ${section.intake.equipmentUploadTicketKey}` : ""}.` : "Create the CX request to upload equipment to the live account."}</span></div></label>
          <label><span>Horizon Training status</span><input type="text" data-newbiz-field="horizonTrainingStatus" value="${section.intake.horizonTrainingStatus || ""}" placeholder="Scheduled, completed, or still needed" /></label>
          <label><span>Schedule for Horizon Training</span><input type="text" data-newbiz-field="horizonTrainingSchedule" value="${section.intake.horizonTrainingSchedule || ""}" placeholder="Date, time, and meeting details" /></label>
          <label class="span-two-field"><span>Users attending Horizon Training</span><textarea rows="3" data-newbiz-field="horizonTrainingUsers" placeholder="List the users who should attend and need Horizon access">${section.intake.horizonTrainingUsers || ""}</textarea><div class="newbiz-inline-actions"><button type="button" class="secondary-button compact-button" data-action="ai-create-horizon-access">AI create Horizon access ticket</button><span>Uses the live account number to prepare the Horizon access request.</span></div></label>
        </form>
      </article>
    `;
  }

  return `<article class="manager-card manager-card-detail" data-manager-section="${section.title}"><strong>${section.title}</strong><p>${section.detail}</p></article>`;
}

function renderManagerWorkspace(account) {
  const isShared = state.viewMode === "shared";
  managerWorkspacePanel.classList.toggle("hidden", isShared);

  if (isShared) return;

  if (!account) {
    managerWorkspace.innerHTML = `
      <div class="empty-state">
        <strong>No private workspace yet</strong>
        <p>Select an account to see the account operations workspace.</p>
      </div>
    `;
    return;
  }

  const operationsSections = getOperationsSections(account);
  if (!operationsSections.length) {
    managerWorkspace.innerHTML = `
      <div class="empty-state">
        <strong>No operations yet</strong>
        <p>Select or create an account to populate samples, calls, complaints, Jira, and onboarding operations.</p>
      </div>
    `;
    return;
  }

  if (!operationsSections.some((section) => section.title === state.activeOperationsTab)) {
    state.activeOperationsTab = operationsSections[0].title;
  }

  const activeSection = operationsSections.find((section) => section.title === state.activeOperationsTab) || operationsSections[0];

  managerWorkspace.innerHTML = `
    <div class="operations-shell">
      <div class="operations-tabs" role="tablist" aria-label="Account operations">
        ${operationsSections.map((section) => `<button type="button" class="operations-tab ${section.title === activeSection.title ? "active" : ""}" data-operations-tab="${section.title}" role="tab" aria-selected="${section.title === activeSection.title ? "true" : "false"}">${section.title}</button>`).join("")}
      </div>
      ${renderOperationSectionCard(activeSection, account)}
    </div>
  `;
}

function renderOnboardingWorkspace(account) {
  const showPanel = state.viewMode === "private" && account?.id === "all-accounts";
  onboardingPanel.classList.toggle("hidden", !showPanel);

  if (!showPanel) {
    onboardingWorkspace.innerHTML = "";
    return;
  }

  const section = account.managerSections.find((item) => item.title === ONBOARDING_SECTION_TITLE);
  if (!section?.intake) {
    onboardingWorkspace.innerHTML = `
      <div class="empty-state">
        <strong>No onboarding draft yet</strong>
        <p>Use the All accounts view to manage onboarding before CX returns the account number.</p>
      </div>
    `;
    return;
  }

  const completion = getOnboardingCompletionPercent(account, section.intake);
  const timeline = getOnboardingTimeline(section.intake);
  const automationFeed =
    section.intake.automationFeed?.length > 0
      ? section.intake.automationFeed
      : getOnboardingAutomationFeed(account, section.intake);

  onboardingWorkspace.innerHTML = `
    <article class="manager-card manager-card-wide" data-manager-section="${section.title}">
      <div class="manager-card-topline">
        <strong>${section.title}</strong>
        <span class="pill ${completion.percent === 100 ? "success" : "warning"}">${completion.percent}% complete</span>
      </div>
      <p>${section.detail}</p>
      <div class="onboarding-top-grid">
        <div class="onboarding-top-column">
          <div class="onboarding-progress">
            <div class="onboarding-progress-bar">
              <span style="width: ${completion.percent}%"></span>
            </div>
            <div class="onboarding-progress-copy">${completion.completed} of ${completion.total} checklist items complete</div>
          </div>
          <div class="checklist onboarding-checklist">
            ${completion.items
              .map(
                (item) => `
                  <button type="button" class="checklist-item checklist-toggle ${item.complete ? "is-complete" : ""}" data-onboarding-checklist-toggle="${item.key}">
                    <span class="checklist-box" aria-hidden="true"></span>
                    <div>
                      <strong>${item.label}</strong>
                      <span>${item.detail}</span>
                    </div>
                  </button>
                `
              )
              .join("")}
          </div>
        </div>
        <div class="onboarding-top-column">
          <form class="onboarding-core-form">
            <label>
              Company name
              <input type="text" data-newbiz-field="onboardingCompanyName" value="${section.intake.onboardingCompanyName || ""}" placeholder="Brand-new account name before CX returns the account number" />
            </label>
            <label>
              Contact person
              <input type="text" data-newbiz-field="onboardingContactName" value="${section.intake.onboardingContactName || ""}" placeholder="Primary onboarding contact" />
            </label>
            <label class="span-two-field">
              Contact email
              <input type="email" data-newbiz-field="onboardingContactEmail" value="${section.intake.onboardingContactEmail || ""}" placeholder="Primary contact email for onboarding and CX follow-up" />
            </label>
            <label>
              Returned account number
              <input type="text" data-newbiz-field="returnedAccountMask" value="${section.intake.returnedAccountMask || ""}" placeholder="Enter the account number once CX returns it" />
            </label>
            <label>
              Onboarding start date
              <input type="date" data-newbiz-field="onboardingStartDate" value="${section.intake.onboardingStartDate || getTodayIsoDate()}" />
            </label>
            <label>
              Projected Volume
              <input type="text" data-newbiz-field="projectedVolume" value="${section.intake.projectedVolume || ""}" placeholder="Expected sample volume, annual usage, or ramp plan" />
            </label>
            <label>
              Quote number
              <input type="text" data-newbiz-field="quoteNumber" value="${section.intake.quoteNumber || ""}" placeholder="Enter the quote number for CX and onboarding reference" />
            </label>
            <label class="span-two-field">
              Account structure set up
              <input type="text" data-newbiz-field="accountStructureSetup" value="${section.intake.accountStructureSetup || ""}" placeholder="Sublevels, naming conventions, user setup notes, and account structure approach" />
            </label>
            <div class="span-two-field newbiz-inline-actions">
              <button type="button" class="primary-button" data-action="start-onboarding">
                Start Onboarding
              </button>
              <button type="button" class="primary-button" data-action="convert-onboarding-to-account">
                Convert to live account
              </button>
              <span>${section.intake.onboardingStartedAt ? `Onboarding ticket created on ${section.intake.onboardingStartedAt}.` : "Start onboarding from here, then convert it once CX returns the account number."}</span>
            </div>
          </form>
        </div>
      </div>
      <div class="onboarding-timeline">
        ${timeline
          .map(
            (item) => `
              <div class="timeline-item">
                <strong>${item.label}</strong>
                <span>${item.date}</span>
                <p>${item.detail}</p>
              </div>
            `
          )
          .join("")}
      </div>
      <div class="automation-shell">
        <div class="manager-card-topline">
          <strong>Automation</strong>
          <span class="pill ${section.intake.automationEnabled ? "success" : "warning"}">${section.intake.automationEnabled ? "Active" : "Paused"}</span>
        </div>
        <div class="newbiz-inline-actions">
          <button type="button" class="secondary-button compact-button" data-action="toggle-onboarding-automation">
            ${section.intake.automationEnabled ? "Pause automation" : "Resume automation"}
          </button>
          <button type="button" class="secondary-button compact-button" data-action="run-onboarding-automation">
            Run automation sweep
          </button>
          <span>${section.intake.automationLastRunAt ? `Last run ${section.intake.automationLastRunAt}.` : "Automation has not been run yet in this onboarding draft."}</span>
        </div>
        <div class="automation-list">
          ${automationFeed
            .map(
              (item) => `
                <div class="automation-item">
                  <strong>${item.label}</strong>
                  <span>${item.status}</span>
                  <p>${item.detail}</p>
                </div>
              `
            )
            .join("")}
        </div>
      </div>
      <form class="new-business-form">
        <label class="span-two-field">
          Billing Info
          <textarea rows="3" data-newbiz-field="billingInfo" placeholder="Billing contact, PO notes, invoice requirements, and remittance details">${section.intake.billingInfo || ""}</textarea>
        </label>
        <label class="span-two-field">
          Accounting requirements
          <textarea rows="3" data-newbiz-field="accountingRequirements" placeholder="What accounting needs to set this account up correctly in NetSuite">${section.intake.accountingRequirements || ""}</textarea>
          <div class="newbiz-inline-actions">
            <button type="button" class="secondary-button compact-button" data-action="prepare-netsuite-handoff">
              AI prepare accounting handoff
            </button>
            <span>${section.intake.netSuiteHandoffPreparedAt ? `Prepared ${section.intake.netSuiteHandoffPreparedAt}.` : "Send the required onboarding info to accounting once requirements are confirmed."}</span>
          </div>
        </label>
        <label class="span-two-field">
          CX account creation
          <div class="newbiz-inline-actions">
            <button type="button" class="secondary-button compact-button" data-action="create-account-creation-ticket">
              Create CX account creation Jira
            </button>
            <span>${section.intake.accountCreationTicketAt ? `Created ${section.intake.accountCreationTicketAt}${section.intake.accountCreationTicketKey ? ` as ${section.intake.accountCreationTicketKey}` : ""}.` : "Creates the CX Jira request to build the account."}</span>
          </div>
        </label>
        <label class="span-two-field">
          Equipment list
          <textarea rows="3" data-newbiz-field="equipmentList" placeholder="List units, assets, sample points, or equipment groups">${section.intake.equipmentList || ""}</textarea>
          <div class="newbiz-inline-actions">
            <a class="external-link-button" href="/equipment-list-template.csv" download="${section.intake.equipmentTemplateName || "equipment-list-template.csv"}">
              Download current equipment form
            </a>
            <button type="button" class="secondary-button compact-button" data-action="create-equipment-upload-ticket">
              Create equipment upload Jira
            </button>
            <span>${section.intake.equipmentUploadTicketAt ? `Created ${section.intake.equipmentUploadTicketAt}${section.intake.equipmentUploadTicketKey ? ` as ${section.intake.equipmentUploadTicketKey}` : ""}.` : "Keep the latest form attached and send the upload request to CX."}</span>
          </div>
        </label>
        <label>
          Horizon Training status
          <input type="text" data-newbiz-field="horizonTrainingStatus" value="${section.intake.horizonTrainingStatus || ""}" placeholder="Scheduled, completed, or still needed" />
        </label>
        <label>
          Schedule for Horizon Training
          <input type="text" data-newbiz-field="horizonTrainingSchedule" value="${section.intake.horizonTrainingSchedule || ""}" placeholder="Date, time, and meeting details" />
        </label>
        <label class="span-two-field">
          Users attending Horizon Training
          <textarea rows="3" data-newbiz-field="horizonTrainingUsers" placeholder="List the users who should attend and need Horizon access">${section.intake.horizonTrainingUsers || ""}</textarea>
          <div class="newbiz-inline-actions">
            <button type="button" class="secondary-button compact-button" data-action="ai-create-horizon-access">
              AI create Horizon access ticket
            </button>
            <span>Uses the training schedule and listed users to create a JIRA request for Horizon access.</span>
          </div>
        </label>
        <label class="span-two-field">
          Weekly check-in notes
          <textarea rows="3" data-newbiz-field="weeklyCheckInNotes" placeholder="Capture weekly check-in highlights and action items">${section.intake.weeklyCheckInNotes || ""}</textarea>
          <div class="newbiz-inline-actions">
            <button type="button" class="secondary-button compact-button" data-action="mark-weekly-checkins-complete">
              Mark weekly check-ins complete
            </button>
            <span>${section.intake.weeklyCheckInsCompletedAt ? `Completed ${section.intake.weeklyCheckInsCompletedAt}.` : "Use when the weekly onboarding cadence has been completed."}</span>
          </div>
        </label>
        <label class="span-two-field">
          45-day review
          <textarea rows="3" data-newbiz-field="day45ReviewNotes" placeholder="Review projected volume, sample volume, and complaints at day 45">${section.intake.day45ReviewNotes || ""}</textarea>
          <div class="newbiz-inline-actions">
            <button type="button" class="secondary-button compact-button" data-action="mark-day45-review-complete">
              Mark 45-day review complete
            </button>
            <span>${section.intake.day45ReviewCompletedAt ? `Completed ${section.intake.day45ReviewCompletedAt}.` : "Confirm the day-45 checkpoint when the review is done."}</span>
          </div>
        </label>
        <label class="span-two-field">
          70-day review
          <textarea rows="3" data-newbiz-field="day70ReviewNotes" placeholder="Check sample volume and note whether pricing needs to be revisited">${section.intake.day70ReviewNotes || ""}</textarea>
          <div class="newbiz-inline-actions">
            <button type="button" class="secondary-button compact-button" data-action="mark-day70-review-complete">
              Mark 70-day review complete
            </button>
            <span>${section.intake.day70ReviewCompletedAt ? `Completed ${section.intake.day70ReviewCompletedAt}.` : "Use when the 70-day sample and pricing check has been completed."}</span>
          </div>
        </label>
        <label class="span-two-field">
          90-day pricing direction
          <textarea rows="3" data-newbiz-field="pricingReviewPlan" placeholder="If projected volume matches the estimate, keep pricing. If not, have a conversation with the customer.">${section.intake.pricingReviewPlan || ""}</textarea>
        </label>
      </form>
    </article>
  `;
}

function renderMicrosoftAccess() {
  const login = getCurrentLogin();

  msBadge.textContent = login.calendarConnected ? "Connected" : "Needs setup";
  msBadge.className = login.calendarConnected ? "pill success" : "pill warning";

  msAccess.innerHTML = `
    <article class="shared-card">
      <strong>Microsoft login</strong>
      <p>${login.microsoftEmail}</p>
    </article>
    <article class="shared-card">
      <strong>Calendar access</strong>
      <p>${login.calendarConnected ? "Calendar sync is available for dashboard scheduling and meeting context." : "Calendar sync is not connected for this login yet."}</p>
    </article>
    <article class="shared-card">
      <strong>SSO status</strong>
      <p>Single sign-on is enabled for every sandbox user and tied to dashboard access.</p>
    </article>
    <article class="shared-card">
      <strong>Account-linked access</strong>
      <p>${state.viewMode === "shared" ? "Shared view stays separate from private account records." : "Private view remains limited to the account records tied to this login."}</p>
    </article>
  `;
}

function renderCalendarPanel() {
  const login = getCurrentLogin();
  const account = getSelectedAccount();
  const calendar = getCalendarSection(account);

  if (!calendar) {
    calendarPanelContent.innerHTML = `
      <div class="empty-state">
        <strong>No calendar context</strong>
        <p>Select an account or shared view to see meetings and new email activity.</p>
      </div>
    `;
    return;
  }

  calendarPanelContent.innerHTML = `
    <div class="calendar-head">
      <div>
        <strong>${calendar.title}</strong>
        <p class="workspace-meta">${login.microsoftEmail}</p>
      </div>
      <span class="pill ${login.calendarConnected ? "success" : "warning"}">${login.calendarConnected ? "Connected" : "Needs setup"}</span>
    </div>
    <p class="workspace-meta">${calendar.detail}</p>
    <div class="outlook-panel">
      <div class="outlook-column">
        <div class="outlook-heading">Upcoming meetings</div>
        <div class="manager-list">
          ${calendar.meetings
            .map(
              (meeting) => `
                <div class="outlook-item">
                  <strong>${meeting.title}</strong>
                  <span>${meeting.time}</span>
                  <p>${meeting.attendees}</p>
                </div>
              `
            )
            .join("")}
        </div>
      </div>
      <div class="outlook-column">
        <div class="outlook-heading">New emails</div>
        <div class="manager-list">
          ${calendar.emails
            .map(
              (email) => `
                <div class="outlook-item">
                  <strong>${email.subject}</strong>
                  <span>${email.sender}</span>
                  <p>${email.time}</p>
                </div>
              `
            )
            .join("")}
        </div>
      </div>
    </div>
  `;
}

function getWorkspaceSearchResults() {
  const query = state.workspaceSearchQuery.trim().toLowerCase();
  if (!query) {
    return [];
  }

  const selectedAccount = getSelectedAccount();
  const searchableAccounts =
    state.viewMode === "shared"
      ? []
      : selectedAccount?.id === "all-accounts"
        ? getLoginAccessibleAccounts()
        : selectedAccount
          ? [selectedAccount]
          : [];

  const accountMatches = searchableAccounts
    .filter((account) =>
      [account.owner, account.email, account.team, account.summary].some((value) =>
        String(value || "").toLowerCase().includes(query)
      )
    )
    .map((account) => ({
      title: account.owner,
      detail: `Account match / Mask: ${account.email} / Owner: ${account.team}`
    }));

  const workspaceMatches = state.sharedItems
    .filter((item) =>
      [item.title, item.detail].some((value) => value.toLowerCase().includes(query))
    )
    .map((item) => ({
      title: item.title,
      detail: `Shared metric / ${item.detail}`
    }));

  const calendarMatches = state.teamCalendar
    .filter((item) =>
      [item.title, item.owner, item.time].some((value) => value.toLowerCase().includes(query))
    )
    .map((item) => ({
      title: item.title,
      detail: `Calendar / ${item.owner} / ${item.time}`
    }));

  const jiraMatches = getVisibleImportedJiraAccounts()
    .filter((bucket) => searchableAccounts.some((account) => account.id === bucket.accountId))
    .flatMap((bucket) =>
      (bucket.items || [])
        .filter((item) =>
          [item.key, item.summary, item.componentName, item.status, item.matchedMask].some((value) =>
            String(value || "").toLowerCase().includes(query)
          )
        )
        .slice(0, 8)
        .map((item) => ({
          title: item.key || "CS ticket",
          detail: `JIRA / ${bucket.mask} / ${buildAiJiraDescription(item)}`,
          accountId: bucket.accountId,
          sectionTitle: "JIRA tickets"
        }))
    );

  return [...jiraMatches, ...accountMatches, ...workspaceMatches, ...calendarMatches];
}

function renderWorkspaceSearch() {
  workspaceSearch.value = state.workspaceSearchQuery;
  const results = getWorkspaceSearchResults();

  if (!state.workspaceSearchQuery.trim()) {
    workspaceSearchResults.innerHTML = `
      <div class="empty-state">
        <strong>Search the workspace</strong>
        <p>${state.viewMode === "shared" ? "Look up shared metrics, calendar items, and team context from one place." : state.selectedId === "all-accounts" ? "Look up accounts, masks, notes, and CS tickets across your private portfolio." : "Look up details for the selected account, including masks, notes, and CS tickets."}</p>
      </div>
    `;
    return;
  }

  if (results.length === 0) {
    workspaceSearchResults.innerHTML = `
      <div class="empty-state">
        <strong>No matches found</strong>
        <p>${state.viewMode === "shared" ? "Try another shared metric, calendar item, or team keyword." : state.selectedId === "all-accounts" ? "Try another account name, owner, account mask, or CS number." : "Try another keyword, account mask, or CS number for this account."}</p>
      </div>
    `;
    return;
  }

  workspaceSearchResults.innerHTML = results
    .map(
      (result) => `
        <article class="shared-card" ${result.accountId ? `data-workspace-account-link="${result.accountId}" data-workspace-section-link="${result.sectionTitle}"` : ""}>
          <strong>${result.title}</strong>
          <p>${result.detail}</p>
        </article>
      `
    )
    .join("");
}

function getEveConversation() {
  if (!state.eveConversations[state.selectedLogin]) {
    state.eveConversations[state.selectedLogin] = [
      {
        id: createLocalId("eve-msg"),
        sender: EVE_NAME,
        text: "I’m ready to flag urgent trends and regional pipeline opportunities."
      }
    ];
  }

  return state.eveConversations[state.selectedLogin];
}

function getEveRegion() {
  const selectedAccount = getSelectedAccount();
  if (state.viewMode === "private" && selectedAccount && selectedAccount.id !== "all-accounts") {
    return selectedAccount.region || getCurrentLogin()?.primaryRegion || "NA";
  }

  return getCurrentLogin()?.primaryRegion || "NA";
}

function getEveVisibleAccounts() {
  if (state.viewMode === "shared") {
    return state.accounts;
  }

  const selectedAccount = getSelectedAccount();
  if (selectedAccount && selectedAccount.id !== "all-accounts") {
    return [selectedAccount];
  }

  return getVisibleAccounts();
}

function getEveUrgentSignals() {
  return getEveVisibleAccounts()
    .flatMap((account) => {
      const sampleSection = account.managerSections.find((section) => section.title === "Sample Tracker");
      const complaintSection = account.managerSections.find((section) => section.title === "Complaints Trackers");
      const jiraBucket = getSimulatedJiraBucket(account);
      const statuses = sampleSection?.sampleData?.statuses || [];
      const onHoldCount = statuses.find((item) => item.label === "On Hold")?.count || 0;
      const loggedOnline = sampleSection?.sampleData?.loggedOnline || 0;
      const received = sampleSection?.sampleData?.received || 0;
      const revenueDelta = (account.spend.currentMonth || 0) - (account.spend.previousMonth || 0);
      const alerts = [];

      if ((complaintSection?.complaintData?.total || 0) >= 3) {
        alerts.push({
          severity: "high",
          accountId: account.id,
          sectionTitle: "Complaints Trackers",
          title: `${account.owner}: complaint trend is elevated`,
          detail: `${complaintSection.complaintData.total} complaints are open or recently recorded. Eve recommends review before the next customer touchpoint.`
        });
      }

      if (onHoldCount >= 2) {
        alerts.push({
          severity: "warning",
          accountId: account.id,
          sectionTitle: "Sample Tracker",
          title: `${account.owner}: samples are sitting on hold`,
          detail: `${onHoldCount} samples are on hold. Eve recommends checking blockers and customer expectations.`
        });
      }

      if (loggedOnline - received >= 4) {
        alerts.push({
          severity: "warning",
          accountId: account.id,
          sectionTitle: "Sample Tracker",
          title: `${account.owner}: intake is outpacing receipts`,
          detail: `${loggedOnline - received} more samples are logged online than received. Eve sees possible intake-to-lab friction.`
        });
      }

      if ((jiraBucket?.open || 0) >= 3) {
        alerts.push({
          severity: "high",
          accountId: account.id,
          sectionTitle: "JIRA tickets",
          title: `${account.owner}: CX queue needs attention`,
          detail: `${jiraBucket.open} Jira items are still open. Eve recommends checking for hidden blockers before they age further.`
        });
      }

      if (revenueDelta < 0) {
        alerts.push({
          severity: "warning",
          accountId: account.id,
          sectionTitle: "Sample Tracker",
          title: `${account.owner}: YTD revenue is trailing prior year`,
          detail: `${formatCurrency(Math.abs(revenueDelta))} below prior year to date. Eve recommends pairing revenue drift with sample and complaint trends.`
        });
      }

      return alerts;
    })
    .sort((left, right) => {
      const score = { high: 2, warning: 1, neutral: 0 };
      return score[right.severity] - score[left.severity];
    })
    .slice(0, 6);
}

function getEveRegionalPipelines() {
  const region = getEveRegion();
  return (state.pipelineLeads || [])
    .filter((lead) => lead.region === region)
    .sort((left, right) => right.fitScore - left.fitScore)
    .slice(0, 4);
}

function getVisiblePipelineLeads() {
  const regionFilter = state.pipelineFilterRegion === "ALL" ? null : state.pipelineFilterRegion;
  const stageFilter = state.pipelineFilterStage === "ALL" ? null : state.pipelineFilterStage;

  return (state.pipelineLeads || [])
    .filter((lead) => !regionFilter || lead.region === regionFilter)
    .filter((lead) => !stageFilter || lead.stage === stageFilter)
    .sort((left, right) => right.fitScore - left.fitScore);
}

function getAccountManagerOptions() {
  return state.loginDirectory.filter((person) => person.role === "Account Manager");
}

function renderPipelinesPanel() {
  const showPanel = state.viewMode === "shared" || state.selectedId === "all-accounts";
  pipelinesPanel.classList.toggle("hidden", !showPanel);
  if (!showPanel) {
    pipelinesPanelContent.innerHTML = "";
    return;
  }

  const visibleLeads = getVisiblePipelineLeads();
  const regionOptions = ["ALL", "NA", "EMEA", "APAC"];
  const accountManagers = getAccountManagerOptions();
  const averageFit = visibleLeads.length
    ? Math.round(visibleLeads.reduce((sum, lead) => sum + (lead.fitScore || 0), 0) / visibleLeads.length)
    : 0;
  const genesysRegion = state.pipelineFilterRegion === "ALL" ? getEveRegion() : state.pipelineFilterRegion;
  const genesysSignal = state.genesysIntegration.byRegion?.[genesysRegion];

  pipelinesPanelContent.innerHTML = `
    <article class="shared-card pipeline-card">
      <div class="pipeline-head">
        <div>
          <strong>Pipeline workspace</strong>
          <p>Configure pipeline stages, add opportunities, and keep regional routing clean for Eve alerts.</p>
        </div>
        <span class="pill neutral">${visibleLeads.length} visible leads</span>
      </div>
      <div class="manager-stats">
        <div class="manager-stat">
          <span>Average fit score</span>
          <strong>${averageFit}%</strong>
        </div>
        <div class="manager-stat">
          <span>Configured stages</span>
          <strong>${state.pipelineStages.length}</strong>
        </div>
        <div class="manager-stat">
          <span>Regional focus</span>
          <strong>${state.pipelineFilterRegion === "ALL" ? getEveRegion() : state.pipelineFilterRegion}</strong>
        </div>
        <div class="manager-stat">
          <span>Top stage count</span>
          <strong>${visibleLeads[0]?.stage || "None"}</strong>
        </div>
      </div>
    </article>
    <article class="shared-card pipeline-card">
      <strong>Genesys routing signal</strong>
      <div class="manager-stats">
        <div class="manager-stat">
          <span>Status</span>
          <strong>${state.genesysIntegration.configured ? "Connected" : "Scaffolded"}</strong>
        </div>
        <div class="manager-stat">
          <span>Hot opportunities</span>
          <strong>${genesysSignal?.hotOpportunities ?? 0}</strong>
        </div>
        <div class="manager-stat">
          <span>Open callbacks</span>
          <strong>${genesysSignal?.openCallbacks ?? 0}</strong>
        </div>
        <div class="manager-stat">
          <span>Abandoned interactions</span>
          <strong>${genesysSignal?.abandonedInteractions ?? 0}</strong>
        </div>
      </div>
      <p class="workspace-meta">${genesysSignal ? `${genesysRegion} queue view from the Genesys scaffold. Eve can use this to flag pipeline leads that deserve AM follow-up.` : "Genesys regional signals will show here once the integration is configured."}</p>
    </article>
    <article class="shared-card pipeline-card">
      <strong>Pipeline configuration</strong>
      <div class="pipeline-config-grid">
        <label class="tool-label">
          Region filter
          <select id="pipeline-region-filter">
            ${regionOptions
              .map((option) => `<option value="${option}" ${state.pipelineFilterRegion === option ? "selected" : ""}>${option === "ALL" ? "All regions" : option}</option>`)
              .join("")}
          </select>
        </label>
        <label class="tool-label">
          Stage filter
          <select id="pipeline-stage-filter">
            <option value="ALL">All stages</option>
            ${state.pipelineStages
              .map((stage) => `<option value="${stage}" ${state.pipelineFilterStage === stage ? "selected" : ""}>${stage}</option>`)
              .join("")}
          </select>
        </label>
        <label class="tool-label span-two-field">
          Configure stages
          <input id="pipeline-stage-input" type="text" value="${state.pipelineStages.join(", ")}" placeholder="Discovery, Qualification, Evaluation..." />
        </label>
      </div>
      <div class="workspace-action-row">
        <button id="save-pipeline-stages-button" class="secondary-button compact-button" type="button">Save stage configuration</button>
      </div>
    </article>
    <article class="shared-card pipeline-card">
      <strong>Add pipeline opportunity</strong>
      <form id="pipeline-lead-form" class="pipeline-form-grid">
        <label>
          Company
          <input name="company" type="text" placeholder="Company name" />
        </label>
        <label>
          Region
          <select name="region">
            ${["NA", "EMEA", "APAC"].map((region) => `<option value="${region}">${region}</option>`).join("")}
          </select>
        </label>
        <label>
          Stage
          <select name="stage">
            ${state.pipelineStages.map((stage) => `<option value="${stage}">${stage}</option>`).join("")}
          </select>
        </label>
        <label>
          Owner
          <input name="owner" type="text" placeholder="Sales or AM owner" />
        </label>
        <label>
          Fit score
          <input name="fitScore" type="number" min="1" max="100" placeholder="0-100" />
        </label>
        <label class="span-two-field">
          Signal
          <input name="signal" type="text" placeholder="Why Eve should care about this lead" />
        </label>
        <label class="span-two-field">
          Next step
          <input name="nextStep" type="text" placeholder="What the team should do next" />
        </label>
        <div class="span-two-field workspace-action-row">
          <button class="primary-button" type="submit">Add opportunity</button>
        </div>
      </form>
    </article>
    <article class="shared-card pipeline-card">
      <strong>Visible opportunities</strong>
      <div class="pipeline-lead-list">
        ${
          visibleLeads.length
            ? visibleLeads
                .map(
                  (lead) => `
                    <div class="pipeline-lead-item">
                      <div class="pipeline-lead-head">
                        <strong>${lead.company}</strong>
                        <span class="pill neutral">${lead.stage}</span>
                      </div>
                      <p>${lead.signal}</p>
                      <div class="eve-meta-row">
                        <span>${lead.region}</span>
                        <span>${lead.fitScore}% fit</span>
                        <span>${lead.owner}</span>
                      </div>
                      <p class="workspace-meta">Next step: ${lead.nextStep}</p>
                      <div class="pipeline-route-row">
                        <label class="tool-label compact-label">
                          Route opportunity
                          <select data-pipeline-route-select="${lead.id}">
                            <option value="">Choose AM</option>
                            ${accountManagers
                              .map(
                                (manager) => `<option value="${manager.name}" ${lead.routedTo === manager.name ? "selected" : ""}>${manager.name}</option>`
                              )
                              .join("")}
                          </select>
                        </label>
                        <button type="button" class="secondary-button compact-button" data-pipeline-route-button="${lead.id}">Route</button>
                      </div>
                      ${lead.routedTo ? `<p class="workspace-meta">Currently routed to ${lead.routedTo}.</p>` : ""}
                    </div>
                  `
                )
                .join("")
            : `<div class="empty-state"><p>No pipeline opportunities match the current filters.</p></div>`
        }
      </div>
    </article>
  `;
}

function buildEveReply(prompt) {
  const normalized = String(prompt || "").toLowerCase();
  const urgentSignals = getEveUrgentSignals();
  const pipelineLeads = getEveRegionalPipelines();
  const region = getEveRegion();

  if (normalized.includes("pipeline") || normalized.includes("lead") || normalized.includes("region")) {
    if (!pipelineLeads.length) {
      return `I do not see any strong pipeline signals in ${region} yet. I’d keep watching for high-fit opportunities before pushing an AM alert.`;
    }

    return `For ${region}, I would surface ${pipelineLeads
      .slice(0, 2)
      .map((lead) => `${lead.company} in ${lead.stage}`)
      .join(" and ")}. These look like the strongest potential customer opportunities right now.`;
  }

  if (normalized.includes("urgent") || normalized.includes("risk") || normalized.includes("issue")) {
    if (!urgentSignals.length) {
      return "I do not see any urgent hidden issues right now. Current account trends look stable across the visible view.";
    }

    return `The most urgent item I see is ${urgentSignals[0].title.toLowerCase()}. ${urgentSignals[0].detail}`;
  }

  return `I’m watching ${urgentSignals.length} urgent trend signal${urgentSignals.length === 1 ? "" : "s"} and ${pipelineLeads.length} regional pipeline opportunit${pipelineLeads.length === 1 ? "y" : "ies"} for ${region}. Ask me about urgent issues or pipeline leads and I’ll narrow it down.`;
}

function renderAiPanel() {
  const login = getCurrentLogin();
  const region = getEveRegion();
  const urgentSignals = getEveUrgentSignals();
  const pipelineLeads = getEveRegionalPipelines();
  const conversation = getEveConversation();

  aiPanel.innerHTML = `
    <article class="shared-card eve-card">
      <div class="eve-head">
        <div>
          <p class="eyebrow">AI chat</p>
          <strong>${EVE_NAME}</strong>
          <p>${login.aiMode}</p>
        </div>
        <span class="pill ${urgentSignals.length ? "warning" : "success"}">${urgentSignals.length} urgent</span>
      </div>
      <div class="eve-meta-row">
        <span>Scope: ${state.viewMode === "shared" ? "shared workspace" : `${login.name}'s private view`}</span>
        <span>Region watch: ${region}</span>
      </div>
      <div class="eve-quick-actions">
        <button type="button" class="secondary-button compact-button" data-eve-prompt="What urgent issues am I missing?">Urgent issues</button>
        <button type="button" class="secondary-button compact-button" data-eve-prompt="Show me pipeline opportunities in my region.">Regional pipeline</button>
        <button type="button" class="secondary-button compact-button" data-eve-prompt="Summarize current account risk.">Account risk</button>
      </div>
    </article>
    <article class="shared-card eve-card">
      <strong>Urgent issues Eve sees</strong>
      <div class="eve-alert-list">
        ${
          urgentSignals.length
            ? urgentSignals
                .map(
                  (alert) => `
                    <button type="button" class="eve-alert-card ${alert.severity}" data-workspace-account-link="${alert.accountId}" data-workspace-section-link="${alert.sectionTitle}">
                      <span class="pill ${alert.severity === "high" ? "warning" : "neutral"}">${alert.severity === "high" ? "Urgent" : "Watch"}</span>
                      <strong>${alert.title}</strong>
                      <p>${alert.detail}</p>
                    </button>
                  `
                )
                .join("")
            : `<div class="empty-state"><p>Eve does not see any urgent hidden issues in the current view.</p></div>`
        }
      </div>
    </article>
    <article class="shared-card eve-card">
      <strong>Regional pipeline alerts</strong>
      <div class="eve-pipeline-list">
        ${
          pipelineLeads.length
            ? pipelineLeads
                .map(
                  (lead) => `
                    <div class="eve-pipeline-card">
                      <div class="eve-pipeline-head">
                        <strong>${lead.company}</strong>
                        <span class="pill neutral">${lead.stage}</span>
                      </div>
                      <p>${lead.signal}</p>
                      <div class="eve-meta-row">
                        <span>${lead.region} region</span>
                        <span>${lead.fitScore}% fit</span>
                      </div>
                      <p class="workspace-meta">Next step: ${lead.nextStep}</p>
                    </div>
                  `
                )
                .join("")
            : `<div class="empty-state"><p>No pipeline opportunities are currently flagged for ${region}.</p></div>`
        }
      </div>
    </article>
    <article class="shared-card eve-card">
      <strong>Chat with ${EVE_NAME}</strong>
      <div class="eve-chat-log">
        ${conversation
          .slice(-6)
          .map(
            (message) => `
              <div class="eve-message ${message.sender === EVE_NAME ? "eve" : "user"}">
                <span class="eve-message-author">${message.sender}</span>
                <p>${message.text}</p>
              </div>
            `
          )
          .join("")}
      </div>
      <div class="eve-composer">
        <textarea id="eve-input" rows="3" placeholder="Ask Eve about urgent risks, regional leads, or account trends...">${state.eveDraft || ""}</textarea>
        <div class="workspace-action-row">
          <button id="eve-send-button" class="primary-button" type="button">Send to Eve</button>
        </div>
      </div>
    </article>
  `;
}

function syncViewButtons() {
  privateViewButton.classList.toggle("active", state.viewMode === "private");
  sharedViewButton.classList.toggle("active", state.viewMode === "shared");
}

function updateSecurityBadge(account) {
  const enabledCount = Object.values(account.security).filter(Boolean).length;
  if (enabledCount === 3) {
    securityBadge.textContent = "Hardened";
    securityBadge.className = "pill success";
    return;
  }

  securityBadge.textContent = "Action needed";
  securityBadge.className = "pill warning";
}

function renderToggles(account) {
  document.querySelectorAll("[data-toggle]").forEach((button) => {
    const key = button.dataset.toggle;
    button.classList.toggle("enabled", Boolean(account.security[key]));
  });
}

function renderDetails() {
  const account = getSelectedAccount();
  const creatorView = isCreatorView();
  const login = getCurrentLogin();

  profilePanel.classList.toggle("hidden", !creatorView);

  if (state.viewMode === "shared") {
    heroLabel.textContent = "Shared workspace";
    accountName.textContent = "Shared view";
    heroPlan.textContent = "Executive summary";
    heroSummary.textContent = "This shared workspace is visible to everyone and shows shared metrics only: revenue, sample volume, invoice exposure, and the consolidated Microsoft calendar.";
    metricMrr.textContent = "$482,300";
    metricHealth.textContent = "184 / 161";
    metricSpend.textContent = "$73,450";
    ownerInput.value = "";
    emailInput.value = "";
    teamInput.value = "";
    regionInput.value = "NA";
    spendCurrent.textContent = "--";
    spendProjected.textContent = "--";
    spendDelta.textContent = "--";
    spendBars.innerHTML = `
      <div class="shared-card">
        <strong>YTD revenue</strong>
        <p>$482,300 total shared company contribution view</p>
      </div>
      <div class="shared-card">
        <strong>Samples</strong>
        <p>184 submitted online / 161 received</p>
      </div>
      <div class="shared-card">
        <strong>Outstanding invoices</strong>
        <p>$73,450 awaiting payment</p>
      </div>
    `;
    securityBadge.textContent = "Shared access";
    securityBadge.className = "pill neutral";
    renderToggles({
      security: { mfa: false, sso: true, audit: false }
    });
    saveButton.disabled = true;
    toggleStatusButton.disabled = true;
    newAccountButton.disabled = !creatorView;
    renderOnboardingWorkspace(null);
    renderManagerWorkspace(null);
    return;
  }

  if (!account) {
    heroLabel.textContent = "Private workspace";
    accountName.textContent = "Your account details will appear here";
    heroPlan.textContent = "Custom";
    heroSummary.textContent = "Use New to create a blank account for this login, then enter only the company, account mask, and owner information you want to review.";
    metricMrr.textContent = "$0";
    metricHealth.textContent = "0%";
    metricSpend.textContent = "$0";
    ownerInput.value = "";
    emailInput.value = "";
    teamInput.value = login?.name || "";
    regionInput.value = "NA";
    spendCurrent.textContent = "$0";
    spendProjected.textContent = "$0";
    spendDelta.textContent = "$0";
    spendBars.innerHTML = `
      <div class="empty-state">
        <strong>No YTD revenue yet</strong>
        <p>Revenue details will populate after you add real company data.</p>
      </div>
    `;
    securityBadge.textContent = "No account selected";
    securityBadge.className = "pill neutral";
    renderToggles({
      security: { mfa: false, sso: true, audit: false }
    });
    saveButton.disabled = true;
    toggleStatusButton.disabled = true;
    newAccountButton.disabled = false;
    renderOnboardingWorkspace(null);
    renderManagerWorkspace(null);
    return;
  }

  if (account.id === "all-accounts") {
    const visibleImportedAccounts = getVisibleImportedJiraAccounts();
    if (!visibleImportedAccounts.some((bucket) => bucket.accountId === state.selectedImportedJiraAccountId)) {
      state.selectedImportedJiraAccountId = visibleImportedAccounts[0]?.accountId || null;
    }
    const spendDeltaValue = account.spend.currentMonth - account.spend.previousMonth;
    const spendMax = Math.max(...account.spend.monthly.map((entry) => entry.value), 1);

    heroLabel.textContent = "Private workspace";
    saveButton.disabled = true;
    toggleStatusButton.disabled = true;
    newAccountButton.disabled = false;

    accountName.textContent = "All accounts";
    heroPlan.textContent = account.plan;
    heroSummary.textContent = account.summary;
    metricMrr.textContent = formatCurrency(account.mrr);
    metricHealth.textContent = `${account.loginHealth}%`;
    metricSpend.textContent = formatCurrency(account.spend.currentMonth);
    ownerInput.value = account.owner;
    emailInput.value = account.email;
    teamInput.value = account.team;
    regionInput.value = "NA";
    spendCurrent.textContent = formatCurrency(account.spend.currentMonth);
    spendProjected.textContent = formatCurrency(account.spend.projected);
    spendDelta.textContent = `${spendDeltaValue >= 0 ? "+" : "-"}${formatCurrency(Math.abs(spendDeltaValue))}`;

    spendBars.innerHTML = account.spend.monthly
      .map(
        (entry) => `
          <div class="spend-bar-row">
            <span>${entry.label}</span>
            <div class="spend-bar-track">
              <div class="spend-bar-value" style="width: ${(entry.value / spendMax) * 100}%"></div>
            </div>
            <strong>${formatCurrency(entry.value)}</strong>
          </div>
        `
      )
      .join("");

    securityBadge.textContent = "Portfolio view";
    securityBadge.className = "pill neutral";
    renderToggles({
      security: { mfa: false, sso: true, audit: false }
    });
    renderOnboardingWorkspace(account);
    renderManagerWorkspace(account);
    return;
  }

  if (state.selectedImportedJiraAccountId !== account.id) {
    state.selectedImportedJiraAccountId = account.id;
    state.selectedImportedJiraComponent = "All";
  }

  const spendDeltaValue = account.spend.currentMonth - account.spend.previousMonth;
  const spendMax = Math.max(...account.spend.monthly.map((entry) => entry.value), 1);

  heroLabel.textContent = "Private workspace";
  saveButton.disabled = false;
  toggleStatusButton.disabled = false;
  newAccountButton.disabled = false;

  accountName.textContent = account.owner || "Untitled account";
  heroPlan.textContent = account.plan || "Custom";
  heroSummary.textContent = account.summary || "Add your account notes, lifecycle context, or owner guidance here.";
  metricMrr.textContent = formatCurrency(account.mrr);
  metricHealth.textContent = `${account.loginHealth}%`;
  metricSpend.textContent = formatCurrency(account.spend.currentMonth);
  ownerInput.value = account.owner;
  emailInput.value = account.email;
  teamInput.value = account.team;
  regionInput.value = account.region;
  spendCurrent.textContent = formatCurrency(account.spend.currentMonth);
  spendProjected.textContent = formatCurrency(account.spend.projected);
  spendDelta.textContent = `${spendDeltaValue >= 0 ? "+" : "-"}${formatCurrency(Math.abs(spendDeltaValue))}`;

  spendBars.innerHTML = account.spend.monthly
    .map(
      (entry) => `
        <div class="spend-bar-row">
          <span>${entry.label}</span>
          <div class="spend-bar-track">
            <div class="spend-bar-value" style="width: ${(entry.value / spendMax) * 100}%"></div>
          </div>
          <strong>${formatCurrency(entry.value)}</strong>
        </div>
      `
    )
    .join("");

  updateSecurityBadge(account);
  renderToggles(account);
  renderOnboardingWorkspace(account);
  renderManagerWorkspace(account);
}

function render() {
  syncSelectedLogin();
  syncSelectedAccount();
  syncViewButtons();
  renderLoginOverlay();
  renderLoginOptions();
  renderActiveUserCard();
  renderLoginSummary();
  renderHomepagePanel();
  renderCalendarPanel();
  renderWorkspaceSummary();
  renderAccounts();
  renderSummaryTable();
  renderMicrosoftAccess();
  renderSharedPanel();
  renderPipelinesPanel();
  renderWorkspaceSearch();
  renderAiPanel();
  renderWorkspaceNotes();
  renderDetails();
  renderDashboardLayout();
  if (!jiraModal.classList.contains("hidden")) {
    syncJiraModalFields();
    renderJiraModalAttachmentList();
  }
}

function persistFormValues() {
  const account = getSelectedAccount();
  if (!account || state.viewMode === "shared") {
    return;
  }

  account.owner = ownerInput.value.trim() || account.owner;
  account.email = emailInput.value.trim() || account.email;
  account.team = teamInput.value.trim() || account.team;
  account.region = regionInput.value;
  queueDashboardStateSync();
}

document.querySelector("#profile-form").addEventListener("input", persistFormValues);

workspaceSearch.addEventListener("input", () => {
  state.workspaceSearchQuery = workspaceSearch.value;
  renderWorkspaceSearch();
});

workspaceSearchResults.addEventListener("click", (event) => {
  const trigger = event.target.closest("[data-workspace-account-link]");
  if (!trigger) {
    return;
  }

  openManagerSectionFromSummary(trigger.dataset.workspaceAccountLink, trigger.dataset.workspaceSectionLink);
  state.workspaceSearchQuery = "";
  renderWorkspaceSearch();
});

aiPanel.addEventListener("input", (event) => {
  const input = event.target.closest("#eve-input");
  if (!input) {
    return;
  }

  state.eveDraft = input.value;
});

aiPanel.addEventListener("click", (event) => {
  const workspaceTrigger = event.target.closest("[data-workspace-account-link]");
  if (workspaceTrigger) {
    openManagerSectionFromSummary(workspaceTrigger.dataset.workspaceAccountLink, workspaceTrigger.dataset.workspaceSectionLink);
    return;
  }

  const quickPrompt = event.target.closest("[data-eve-prompt]");
  const sendTrigger = event.target.closest("#eve-send-button");
  if (!quickPrompt && !sendTrigger) {
    return;
  }

  const prompt = quickPrompt ? quickPrompt.dataset.evePrompt : state.eveDraft.trim();
  if (!prompt) {
    showToast("Add a message for Eve first.");
    return;
  }

  const conversation = getEveConversation();
  conversation.push({
    id: createLocalId("eve-msg"),
    sender: getCurrentLogin().name,
    text: prompt
  });
  conversation.push({
    id: createLocalId("eve-msg"),
    sender: EVE_NAME,
    text: buildEveReply(prompt)
  });
  state.eveDraft = "";
  renderAiPanel();
  queueDashboardStateSync();
});

pipelinesPanelContent.addEventListener("change", (event) => {
  const regionFilter = event.target.closest("#pipeline-region-filter");
  if (regionFilter) {
    state.pipelineFilterRegion = regionFilter.value;
    renderPipelinesPanel();
    return;
  }

  const stageFilter = event.target.closest("#pipeline-stage-filter");
  if (stageFilter) {
    state.pipelineFilterStage = stageFilter.value;
    renderPipelinesPanel();
  }
});

pipelinesPanelContent.addEventListener("click", (event) => {
  const routeButton = event.target.closest("[data-pipeline-route-button]");
  if (routeButton) {
    const leadId = routeButton.dataset.pipelineRouteButton;
    const select = pipelinesPanelContent.querySelector(`[data-pipeline-route-select="${leadId}"]`);
    const routedTo = String(select?.value || "").trim();
    if (!routedTo) {
      showToast("Choose an AM to route this opportunity.");
      return;
    }

    const lead = state.pipelineLeads.find((item) => item.id === leadId);
    if (!lead) {
      showToast("Could not find that pipeline opportunity.");
      return;
    }

    lead.routedTo = routedTo;
    lead.owner = routedTo;
    lead.nextStep = `Routed to ${routedTo} for account-manager follow-up.`;
    renderPipelinesPanel();
    renderAiPanel();
    queueDashboardStateSync();
    showToast(`Opportunity routed to ${routedTo}.`);
    return;
  }

  const saveStages = event.target.closest("#save-pipeline-stages-button");
  if (!saveStages) {
    return;
  }

  const stageInput = pipelinesPanelContent.querySelector("#pipeline-stage-input");
  const stages = String(stageInput?.value || "")
    .split(",")
    .map((item) => item.trim())
    .filter(Boolean);

  if (!stages.length) {
    showToast("Add at least one pipeline stage.");
    return;
  }

  state.pipelineStages = stages;
  if (!stages.includes(state.pipelineFilterStage)) {
    state.pipelineFilterStage = "ALL";
  }
  renderPipelinesPanel();
  renderAiPanel();
  queueDashboardStateSync();
  showToast("Pipeline stages updated.");
});

pipelinesPanelContent.addEventListener("submit", (event) => {
  const form = event.target.closest("#pipeline-lead-form");
  if (!form) {
    return;
  }

  event.preventDefault();
  const formData = new FormData(form);
  const company = String(formData.get("company") || "").trim();
  const region = String(formData.get("region") || "NA");
  const stage = String(formData.get("stage") || state.pipelineStages[0] || "Discovery");
  const owner = String(formData.get("owner") || "").trim();
  const fitScore = Number(formData.get("fitScore") || 0);
  const signal = String(formData.get("signal") || "").trim();
  const nextStep = String(formData.get("nextStep") || "").trim();

  if (!company || !owner || !signal || !nextStep || !fitScore) {
    showToast("Complete all pipeline fields first.");
    return;
  }

  state.pipelineLeads.unshift({
    id: createLocalId("lead"),
    company,
    region,
    stage,
    fitScore,
    signal,
    owner,
    nextStep
  });

  form.reset();
  state.pipelineFilterRegion = region;
  state.pipelineFilterStage = "ALL";
  renderPipelinesPanel();
  renderAiPanel();
  queueDashboardStateSync();
  showToast("Pipeline opportunity added.");
});

accountFilterInput.addEventListener("input", () => {
  state.accountFilterQuery = accountFilterInput.value;
  state.accountListScroll = 0;
  accountList.scrollTop = 0;
  render();
});

accountList.addEventListener("scroll", () => {
  state.accountListScroll = accountList.scrollTop;
  renderAccounts();
});

accountList.addEventListener("click", (event) => {
  const trigger = event.target.closest("[data-account-item]");
  if (!trigger) {
    return;
  }

  state.selectedId = trigger.dataset.accountItem;
  render();
});

summaryTableBody.addEventListener("click", (event) => {
  const trigger = event.target.closest("[data-summary-section-link]");
  if (!trigger) {
    return;
  }

  state.selectedSummaryBreakdown = {
    accountId: trigger.dataset.summaryAccountLink,
    metric: trigger.dataset.summaryMetric,
    sectionTitle: trigger.dataset.summarySectionLink
  };
  renderSummaryBreakdown();
});

summaryBreakdown.addEventListener("click", (event) => {
  const closeTrigger = event.target.closest("[data-summary-close]");
  if (closeTrigger) {
    state.selectedSummaryBreakdown = null;
    renderSummaryBreakdown();
    return;
  }

  const workspaceTrigger = event.target.closest("[data-summary-open-workspace]");
  if (!workspaceTrigger) {
    return;
  }

  openManagerSectionFromSummary(workspaceTrigger.dataset.summaryOpenWorkspace, workspaceTrigger.dataset.summarySectionTitle);
});

workspaceNoteInput.addEventListener("input", () => {
  state.workspaceNoteDraft = workspaceNoteInput.value;
});

addWorkspaceNoteButton.addEventListener("click", () => {
  const text = state.workspaceNoteDraft.trim();
  if (!text) {
    showToast("Add note text first.");
    return;
  }

  if (state.viewMode === "shared") {
    state.sharedWorkspaceNotes.unshift({
      id: createLocalId("shared-note"),
      scope: "Shared workspace",
      author: state.selectedLogin,
      text
    });
  } else if (state.selectedId === "all-accounts") {
    if (!state.privateWorkspaceNotes[state.selectedLogin]) {
      state.privateWorkspaceNotes[state.selectedLogin] = [];
    }
    state.privateWorkspaceNotes[state.selectedLogin].unshift({
      id: createLocalId("private-note"),
      scope: "All accounts",
      author: state.selectedLogin,
      text
    });
  } else {
    const account = getSelectedAccount();
    if (!account) {
      showToast("Select an account first.");
      return;
    }
    account.notes.unshift({
      id: createLocalId("account-note"),
      scope: account.owner,
      author: state.selectedLogin,
      text
    });
  }

  state.workspaceNoteDraft = "";
  renderWorkspaceNotes();
  queueDashboardStateSync();
  showToast("Workspace note added.");
});

document.querySelectorAll("[data-toggle]").forEach((button) => {
  button.addEventListener("click", () => {
    const account = getSelectedAccount();
    if (!account || state.viewMode === "shared") {
      return;
    }
    const key = button.dataset.toggle;
    account.security[key] = !account.security[key];
    account.loginHealth = Object.values(account.security).filter(Boolean).length * 17 + 47;
    renderDetails();
    queueDashboardStateSync();
  });
});

async function handleWorkspaceClick(event) {
  const operationsTab = event.target.closest("[data-operations-tab]");
  if (operationsTab) {
    state.activeOperationsTab = operationsTab.dataset.operationsTab;
    renderManagerWorkspace(getSelectedAccount());
    return;
  }

  const maskTrigger = event.target.closest("[data-jira-mask-account]");
  if (maskTrigger) {
    state.selectedImportedJiraAccountId = maskTrigger.dataset.jiraMaskAccount;
    state.selectedImportedJiraComponent = "All";
    render();
    return;
  }

  const componentTrigger = event.target.closest("[data-jira-component-button]");
  if (componentTrigger) {
    const value = componentTrigger.dataset.jiraComponentButton;
    state.selectedImportedJiraComponent = state.selectedImportedJiraComponent === value ? "All" : value;
    render();
    return;
  }

  const trigger = event.target.closest("[data-action='open-jira-modal']");
  if (trigger) {
    openJiraModal();
    return;
  }

  const onboardingTrigger = event.target.closest("[data-action='start-onboarding']");
  if (onboardingTrigger) {
    if (state.viewMode === "shared") {
      return;
    }

    const intake = getEditableOnboardingIntake();
    const account = getOnboardingContextAccount();
    if (!intake || !account) {
      return;
    }

    const result = await createOnboardingTicket(account, intake);
    if (!result.created) {
      showToast(`Complete these fields first: ${result.missingFields.slice(0, 3).join(", ")}${result.missingFields.length > 3 ? "..." : ""}`);
      return;
    }

    render();
    queueDashboardStateSync();
    showToast(result.remote?.ok ? `Onboarding pushed to Jira as ${result.remote.key}.` : `Onboarding saved locally for ${account.owner}.`);
    return;
  }

  const accountCreationTrigger = event.target.closest("[data-action='create-account-creation-ticket']");
  if (accountCreationTrigger) {
    const account = getOnboardingContextAccount();
    const intake = getEditableOnboardingIntake();
    if (!intake || !account) {
      return;
    }

    const created = await createAccountCreationTicket(account, intake);
    render();
    queueDashboardStateSync();
    showToast(created.remote?.ok ? `Account creation pushed to Jira as ${created.remote.key}.` : `Account creation request saved locally for ${account.owner}.`);
    return;
  }

  const convertTrigger = event.target.closest("[data-action='convert-onboarding-to-account']");
  if (convertTrigger) {
    const intake = getEditableOnboardingIntake();
    if (!intake || state.viewMode === "shared" || state.selectedId !== "all-accounts") {
      return;
    }

    const result = convertOnboardingDraftToAccount(intake);
    if (!result.created) {
      showToast(`Complete these fields first: ${result.missing.slice(0, 3).join(", ")}${result.missing.length > 3 ? "..." : ""}`);
      return;
    }

    render();
    queueDashboardStateSync();
    showToast(`Converted onboarding into live account ${result.account.owner} (${result.account.email}).`);
    return;
  }

  const checklistToggle = event.target.closest("[data-onboarding-checklist-toggle]");
  if (checklistToggle) {
    const intake = getEditableOnboardingIntake();
    const account = getOnboardingContextAccount();
    if (!intake || !account) {
      return;
    }

    if (!intake.checklistStates) {
      intake.checklistStates = {};
    }

    const items = getOnboardingChecklistItems(account, intake);
    const target = items.find((item) => item.key === checklistToggle.dataset.onboardingChecklistToggle);
    if (!target) {
      return;
    }

    intake.checklistStates[target.key] = !target.complete;
    render();
    queueDashboardStateSync();
    showToast(`${target.label} ${!target.complete ? "marked complete" : "reopened"}.`);
    return;
  }

  const automationToggle = event.target.closest("[data-action='toggle-onboarding-automation']");
  if (automationToggle) {
    const intake = getEditableOnboardingIntake();
    if (!intake) {
      return;
    }

    intake.automationEnabled = !intake.automationEnabled;
    render();
    queueDashboardStateSync();
    showToast(`Onboarding automation ${intake.automationEnabled ? "resumed" : "paused"}.`);
    return;
  }

  const automationRun = event.target.closest("[data-action='run-onboarding-automation']");
  if (automationRun) {
    const intake = getEditableOnboardingIntake();
    const account = getOnboardingContextAccount();
    if (!intake || !account) {
      return;
    }

    runOnboardingAutomationSweep(account, intake);
    render();
    queueDashboardStateSync();
    showToast("Onboarding automation sweep completed.");
    return;
  }

  const equipmentTrigger = event.target.closest("[data-action='create-equipment-upload-ticket']");
  if (equipmentTrigger) {
    const account = getOnboardingContextAccount();
    const intake = getEditableOnboardingIntake();
    if (!intake || !account) {
      return;
    }

    const created = await createEquipmentUploadTicket(account, intake);
    render();
    queueDashboardStateSync();
    showToast(created.remote?.ok ? `Equipment upload pushed to Jira as ${created.remote.key}.` : `Equipment upload request saved locally for ${account.owner}.`);
    return;
  }

  const accountingTrigger = event.target.closest("[data-action='prepare-netsuite-handoff']");
  if (accountingTrigger) {
    const account = getOnboardingContextAccount();
    const intake = getEditableOnboardingIntake();
    if (!intake || !account) {
      return;
    }

    prepareNetSuiteHandoff(account, intake);
    render();
    queueDashboardStateSync();
    showToast("Accounting handoff prepared for NetSuite.");
    return;
  }

  const weeklyCheckInTrigger = event.target.closest("[data-action='mark-weekly-checkins-complete']");
  if (weeklyCheckInTrigger) {
    const intake = getEditableOnboardingIntake();
    if (!intake) {
      return;
    }

    intake.weeklyCheckInsCompletedAt = new Date().toLocaleDateString("en-US");
    render();
    queueDashboardStateSync();
    showToast("Weekly check-ins marked complete.");
    return;
  }

  const day45Trigger = event.target.closest("[data-action='mark-day45-review-complete']");
  if (day45Trigger) {
    const intake = getEditableOnboardingIntake();
    if (!intake) {
      return;
    }

    intake.day45ReviewCompletedAt = new Date().toLocaleDateString("en-US");
    render();
    queueDashboardStateSync();
    showToast("45-day review marked complete.");
    return;
  }

  const day70Trigger = event.target.closest("[data-action='mark-day70-review-complete']");
  if (day70Trigger) {
    const intake = getEditableOnboardingIntake();
    if (!intake) {
      return;
    }

    intake.day70ReviewCompletedAt = new Date().toLocaleDateString("en-US");
    render();
    queueDashboardStateSync();
    showToast("70-day review marked complete.");
    return;
  }

  const horizonTrigger = event.target.closest("[data-action='ai-create-horizon-access']");
  if (!horizonTrigger || state.viewMode === "shared") {
    return;
  }

  const account = getOnboardingContextAccount();
  const intake = getEditableOnboardingIntake();
  if (!intake || !account) {
    return;
  }

  const created = await createHorizonAccessTicket(account, intake);
  if (!created?.created) {
    showToast("Could not create a Horizon access ticket.");
    return;
  }

  render();
  queueDashboardStateSync();
  showToast(created.remote?.ok ? `Horizon access pushed to Jira as ${created.remote.key}.` : `Horizon access request saved locally for ${account.owner}.`);
}

function handleWorkspaceChange(event) {
  const filter = event.target.closest("[data-jira-component-filter]");
  if (!filter) {
    return;
  }

  state.selectedImportedJiraComponent = filter.value;
  render();
}

function handleWorkspaceInput(event) {
  const field = event.target.closest("[data-newbiz-field]");
  if (!field || state.viewMode === "shared") {
    return;
  }

  const intake = getEditableOnboardingIntake();
  if (!intake) {
    return;
  }

  intake[field.dataset.newbizField] = field.value;
  queueDashboardStateSync();
}

managerWorkspace.addEventListener("click", handleWorkspaceClick);
managerWorkspace.addEventListener("change", handleWorkspaceChange);
managerWorkspace.addEventListener("input", handleWorkspaceInput);

onboardingWorkspace.addEventListener("click", handleWorkspaceClick);
onboardingWorkspace.addEventListener("change", handleWorkspaceChange);
onboardingWorkspace.addEventListener("input", handleWorkspaceInput);

jiraModal.addEventListener("click", (event) => {
  if (event.target.closest("[data-close-jira-modal]")) {
    closeJiraModal();
  }
});

jiraModalAccountSelect.addEventListener("change", () => {
  state.jiraModalAccountId = jiraModalAccountSelect.value;
  syncJiraModalFields();
  renderJiraModalAttachmentList();
});

jiraModalAttachments.addEventListener("change", renderJiraModalAttachmentList);

jiraModalForm.addEventListener("submit", async (event) => {
  event.preventDefault();

  const account = getAccountById(jiraModalAccountSelect.value);
  const title = jiraModalTitleInput.value.trim();
  const description = jiraModalDescription.value.trim();
  const priority = jiraModalPriority.value || "Medium";

  if (!account) {
    showToast("Choose an account first.");
    return;
  }

  if (!title) {
    showToast("Add a JIRA title first.");
    return;
  }

  const attachmentNames = Array.from(jiraModalAttachments.files || []).map((file) => file.name);
  account.owner = jiraModalCompany.value.trim() || account.owner;
  account.email = jiraModalMask.value.trim() || account.email;
  account.contactName = jiraModalContactName.value.trim() || account.contactName;
  account.contactEmail = jiraModalContactEmail.value.trim() || account.contactEmail;

  const request = {
    id: createLocalId("jira"),
    title,
    priority,
    status: "Open",
    createdAt: new Date().toISOString(),
    owner: state.selectedLogin,
    account: maskOrFallback(account.email),
    contactName: account.contactName,
    contactEmail: account.contactEmail,
    cc: jiraModalCc.value.trim(),
    description,
    attachments: attachmentNames
  };

  const remote = await submitJiraRequest(account, request);
  if (remote.ok) {
    request.status = "Submitted";
    request.externalKey = remote.key;
    request.externalUrl = remote.browseUrl;
    request.title = `${request.title} (${remote.key})`;
  }
  registerLocalJiraRequest(account, request);

  closeJiraModal();
  render();
  queueDashboardStateSync();
  showToast(remote.ok ? `Jira created: ${remote.key}` : remote.configured ? `Jira create failed; saved locally for ${account.owner}.` : `Jira not configured; saved locally for ${account.owner}.`);
});

document.addEventListener("keydown", (event) => {
  if (event.key === "Escape" && !jiraModal.classList.contains("hidden")) {
    closeJiraModal();
  }
});

async function bootstrapApp() {
  await loadDashboardState();
  state.accounts.forEach((account) => {
    applyImportedComplaintData(account);
  });
  await loadIntegrationConfig();
  render();
  queueDashboardStateSync();
}

dashboardGrid.addEventListener("click", (event) => {
  const collapseTrigger = event.target.closest("[data-panel-collapse]");
  if (!collapseTrigger) {
    return;
  }

  const layout = getCurrentPanelLayout();
  const id = collapseTrigger.dataset.panelCollapse;
  layout.collapsed[id] = !layout.collapsed[id];
  saveCurrentPanelLayout(layout);
  renderDashboardLayout();
});

dashboardGrid.addEventListener("dragstart", (event) => {
  const handle = event.target.closest("[data-panel-handle]");
  if (!handle) {
    return;
  }

  state.draggedPanelId = handle.dataset.panelHandle;
  event.dataTransfer.effectAllowed = "move";
});

dashboardGrid.addEventListener("dragover", (event) => {
  if (!state.draggedPanelId) {
    return;
  }
  event.preventDefault();
});

dashboardGrid.addEventListener("drop", (event) => {
  const targetPanel = event.target.closest(".dashboard-panel");
  if (!state.draggedPanelId || !targetPanel) {
    return;
  }

  event.preventDefault();
  const targetId = targetPanel.dataset.panelId;
  if (targetId === state.draggedPanelId) {
    state.draggedPanelId = null;
    return;
  }

  const layout = getCurrentPanelLayout();
  const order = [...layout.order.filter((id) => id !== state.draggedPanelId)];
  const targetIndex = order.indexOf(targetId);
  const rect = targetPanel.getBoundingClientRect();
  const insertAfter = event.clientY > rect.top + rect.height / 2;
  const insertIndex = targetIndex + (insertAfter ? 1 : 0);
  order.splice(insertIndex, 0, state.draggedPanelId);
  layout.order = order;
  saveCurrentPanelLayout(layout);
  state.draggedPanelId = null;
  renderDashboardLayout();
});

dashboardGrid.addEventListener("dragend", () => {
  state.draggedPanelId = null;
});

privateViewButton.addEventListener("click", () => {
  state.viewMode = "private";
  render();
});

sharedViewButton.addEventListener("click", () => {
  state.viewMode = "shared";
  render();
});

loginSelect.addEventListener("change", () => {
  state.selectedLogin = loginSelect.value;
  render();
});

if (loginCardList) {
  loginCardList.addEventListener("click", (event) => {
    const trigger = event.target.closest("[data-login-card]");
    if (!trigger) {
      return;
    }

    state.selectedLogin = trigger.dataset.loginCard;
    state.loginOverlayOpen = false;
    state.viewMode = "private";
    state.selectedId = "all-accounts";
    render();
    showToast(`Signed in as ${state.selectedLogin}.`);
  });
}

saveButton.addEventListener("click", () => {
  if (!getSelectedAccount() || state.viewMode === "shared") {
    return;
  }
  persistFormValues();
  render();
  queueDashboardStateSync();
  showToast("Changes saved to the shared workspace.");
});

toggleStatusButton.addEventListener("click", () => {
  const account = getSelectedAccount();
  if (!account || state.viewMode === "shared") {
    return;
  }
  const transitions = {
    active: "review",
    review: "inactive",
    inactive: "active"
  };

  account.status = transitions[account.status];
  showToast(`Status changed to ${account.status}.`);
  render();
  queueDashboardStateSync();
});

newAccountButton.addEventListener("click", () => {
  if (state.viewMode === "shared") {
    return;
  }

  const nextId = `acct-${Math.floor(Math.random() * 900 + 100)}`;
  const login = getCurrentLogin();
  const newAccount = createAccount(nextId, "", "", login?.name || "");
  newAccount.status = "review";
  applyImportedComplaintData(newAccount);

  state.accounts.unshift(newAccount);
  state.selectedId = newAccount.id;
  render();
  queueDashboardStateSync();
  showToast("New account created.");
});

resetLayoutButton.addEventListener("click", () => {
  saveCurrentPanelLayout({
    order: [...DEFAULT_PANEL_ORDER],
    collapsed: {},
    viewOption: "comfortable"
  });
  renderDashboardLayout();
  showToast("Layout reset.");
});

viewOptionSelect.addEventListener("change", () => {
  const layout = getCurrentPanelLayout();
  layout.viewOption = viewOptionSelect.value;
  saveCurrentPanelLayout(layout);
  renderDashboardLayout();
  showToast(`View set to ${viewOptionSelect.value}.`);
});

bootstrapApp();
