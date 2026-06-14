const EMPTY_LINE_ITEMS_MESSAGE = "AI analysis will populate line items here.";
const DEFAULT_PROFILE_ID = "koncept";
const DEFAULT_PRICING_REFERENCE_ID = "koncept-exhibition-quotation";
const DEFAULT_SAMPLE_ID = "kent-group";
const CSRF_HEADER_NAME = "X-Swooshz-CSRF";
const LEGACY_QUOTE_PRESETS_STORAGE_KEY = "swooshz_quote_detail_presets_v1";
const QUOTE_SESSION_STORAGE_KEY = "swooshz_quote_session_v1";
const QUOTE_SESSION_FILE_DB_NAME = "swooshz_quote_session_files_v1";
const QUOTE_SESSION_FILE_STORE_NAME = "reference_files";
const QUOTE_SESSION_FILE_DB_VERSION = 1;
const QUOTE_SESSION_STATE_VERSION = 4;
const OUTPUT_SORT_MODES = ["pricing_reference", "category", "name", "category_name"];
const FINAL_JOB_STATUSES = new Set(["completed", "degraded", "needs_review", "blocked", "failed"]);
const PROFILE_PRESET_PREFIX = "profile:";
const PRICING_REFERENCE_FILE_ACCEPT = ".xlsx,.csv,.md";
const REFERENCE_FILE_ACCEPT = "image/png,image/jpeg,image/webp,application/pdf,.pdf";
const MAX_PRICING_REFERENCE_FILE_BYTES = 10 * 1024 * 1024;
const MAX_REFERENCE_IMAGES = 8;
const DEFAULT_DATE_LABEL = "Date:";
const DEFAULT_TERMS_HEADING = "Terms & Conditions:";
const DEFAULT_NOTES_HEADING = "Note:";
const DEFAULT_ACCEPTANCE_TEXT = "We accept the quotation amount and the terms";
const DEFAULT_PERSON_LABEL = "Person in charge";
const DEFAULT_STAMP_LABEL = "Company name & stamp";
const DEFAULT_TAX_LABEL = "GST";
const DEFAULT_TAX_RATE = 0.09;
const DEFAULT_CURRENCY_LABEL = "SGD";
const CUSTOM_CURRENCY_VALUE = "__CUSTOM__";
const CURRENCY_OPTIONS = [
  ["SGD", "Singapore Dollar"],
  ["AUD", "Australian Dollar"],
  ["CNY", "Chinese Yuan"],
  ["EUR", "Euro"],
  ["GBP", "Pound Sterling"],
  ["IDR", "Indonesian Rupiah"],
  ["MYR", "Malaysian Ringgit"],
  ["THB", "Thai Baht"],
  ["USD", "US Dollar"],
];
const DEFAULT_QUOTE_COMPANY_RICH_TEXT = {
  termsHeading: "<div>Terms &amp; Conditions:</div>",
  notesHeading: "<div>Note:</div>",
  acceptanceText: "<div>We accept the quotation amount and the terms</div>",
  personLabel: "<div>Person in charge</div>",
  stampLabel: "<div>Company name &amp; stamp</div>",
  konceptDateLabel: "<div>Date:</div>",
  dateLabel: "<div>Date:</div>",
};
const DEFAULT_BOOTH_DIMENSIONS = {
  booth_width: "6",
  booth_depth: "6",
  booth_size: "6m x 6m",
  dimension_source: "default",
};

const QUOTE_COPY = {
  label: "Quotation",
  intakeSubtitle: "Drop reference images or PDFs for a real quote, or load the demo fixture for a quick test run.",
  dropTitle: "Drop reference images or PDFs to start",
  dropMeta: `JPG, PNG, WebP, or PDF. Up to ${MAX_REFERENCE_IMAGES} references, 12 MB each; remote AI analysis sends selected refs to the configured provider.`,
  imageNoun: "reference file",
  analyzeLabel: "Start Analysis",
  fallbackAction: "review the reference files",
};

const SIDE_PANEL_SEQUENCE = ["images", "customer", "quote_company", "basis", "output"];
const RICH_TEXT_SOURCE_IDS = [
  "clientName",
  "clientAttention",
  "clientTitle",
  "clientAddress",
  "projectTitle",
  "projectNumber",
  "headerDetails",
  "termsHeading",
  "paymentTerms",
  "notesHeading",
  "standardNotes",
  "quoteCompanyName",
  "acceptanceText",
  "konceptSignatory",
  "konceptTitle",
  "konceptDateLabel",
  "personLabel",
  "stampLabel",
  "dateLabel",
];

const EMPTY_BASIS = {
  surfaces: "",
  counters: "",
  platform: "",
  graphics: "",
  furniture: "",
  electrical: "",
};

const BASIS_FIELDS = [
  ["surfaces", "Surfaces / Structures"],
  ["counters", "Cabinets / Counters"],
  ["platform", "Platform / Flooring"],
  ["graphics", "Graphics / Signage"],
  ["furniture", "Furniture / Plants / AV"],
  ["electrical", "Electrical"],
];

const BASIS_TAGS = [
  ["Include", "Include", "Confirmed in the draft"],
  ["Exclude", "Exclude", "Not included unless requested"],
  ["Custom", "AI Proposal", "Not found in pricing reference"],
  ["Confirm", "Confirm", "Needs include, exclude, or revision"],
];

const state = {
  profileId: "",
  pricingReferenceId: "",
  pricingReferenceSource: "",
  selectedPresetValue: "",
  profiles: [],
  pricingReferences: [],
  defaultPricingReferenceId: DEFAULT_PRICING_REFERENCE_ID,
  images: [],
  headerLogo: null,
  workflowStage: "needs_images",
  quoteBasis: { ...EMPTY_BASIS },
  quoteBasisSections: [],
  lineItems: [],
  outputRows: [],
  originalOutputRows: [],
  outputErrors: [],
  outputSortMode: "pricing_reference",
  analysisFindings: [],
  blockingClarificationQuestions: [],
  boothDimensions: { ...DEFAULT_BOOTH_DIMENSIONS },
  originalAnalysisSnapshot: null,
  basisConfirmed: false,
  isAnalysisRunning: false,
  isGenerating: false,
  aiFailed: false,
  draftSource: "",
  lastAnalysisMode: "standard",
  pendingAnalysisMode: "standard",
  isBooting: true,
  isPageUnloading: false,
  csrfHeaderName: CSRF_HEADER_NAME,
  csrfToken: "",
  pendingFeedback: "",
  activeSidePanel: "images",
  downloadFile: null,
  pricingMatches: [],
  pricingIssues: [],
  activeJob: null,
  pendingPricingReference: null,
  pricingReferenceImportBusy: false,
  pricingReferenceImportToken: "",
  permissions: {
    role: "viewer",
    canManageSettings: false,
    canManagePricingReferences: false,
    canManageProfiles: false,
    canImportPricingReferences: false,
  },
  quoteDateFormat: {
    bold: false,
    italic: false,
    underline: false,
  },
  basisChat: {
    scope: "quote",
    field: "",
    sectionId: "",
    lineIndex: -1,
    line: "",
    quantity: "",
    unit: "",
    quantityLabel: "",
    proposal: null,
  },
};

const qs = (selector) => document.querySelector(selector);
let analysisElapsedTimerId = 0;
let sessionFileDbPromise = null;

const elements = {
  healthText: qs("#healthText"),
  statusDot: qs("#statusDot"),
  topbarStatus: qs("#topbarStatus"),
  imageIntake: qs("#imageIntake"),
  dropTitle: qs("#dropTitle"),
  dropMeta: qs("#dropMeta"),
  dropzone: qs(".dropzone"),
  imageInput: qs("#imageInput"),
  imageUploadStatus: qs("#imageUploadStatus"),
  fileList: qs("#fileList"),
  newQuoteButton: qs("#newQuoteButton"),
  sideWorkspace: qs("#sideWorkspace"),
  sideDrawerTitle: qs("#sideDrawerTitle"),
  sideDrawerEyebrow: qs("#sideDrawerEyebrow"),
  sideDrawerSubtitle: qs("#sideDrawerSubtitle"),
  sampleDetailsButton: qs("#sampleDetailsButton"),
  clientName: qs("#clientName"),
  clientAttention: qs("#clientAttention"),
  clientTitle: qs("#clientTitle"),
  clientAddress: qs("#clientAddress"),
  projectTitle: qs("#projectTitle"),
  quoteDate: qs("#quoteDate"),
  quoteDateFormatButtons: Array.from(document.querySelectorAll("[data-date-format-command]")),
  projectNumber: qs("#projectNumber"),
  headerDetails: qs("#headerDetails"),
  headerLogoInput: qs("#headerLogoInput"),
  headerLogoPreview: qs("#headerLogoPreview"),
  termsHeading: qs("#termsHeading"),
  paymentTerms: qs("#paymentTerms"),
  notesHeading: qs("#notesHeading"),
  standardNotes: qs("#standardNotes"),
  quoteCompanyName: qs("#quoteCompanyName"),
  acceptanceText: qs("#acceptanceText"),
  konceptSignatory: qs("#konceptSignatory"),
  konceptTitle: qs("#konceptTitle"),
  konceptDateLabel: qs("#konceptDateLabel"),
  personLabel: qs("#personLabel"),
  stampLabel: qs("#stampLabel"),
  dateLabel: qs("#dateLabel"),
  taxLabel: qs("#taxLabel"),
  taxRate: qs("#taxRate"),
  aiFailureBanner: qs("#aiFailureBanner"),
  basisReviewSurface: qs("#basisReviewSurface"),
  resultStatus: qs("#resultStatus"),
  outputStatusPill: qs("#outputStatusPill"),
  outputSourceLabel: qs("#outputSourceLabel"),
  outputTotalLines: qs("#outputTotalLines"),
  messageList: qs("#messageList"),
  matchSummary: qs("#matchSummary"),
  pricingMatchesBody: qs("#pricingMatchesBody"),
  pricingTableWrap: qs("#pricingTableWrap"),
  pricingEmptyState: qs("#pricingEmptyState"),
  pricingReviewMessages: qs("#pricingReviewMessages"),
  profileSelect: qs("#profileSelect"),
  presetNameInput: qs("#presetNameInput"),
  presetSelect: qs("#presetSelect"),
  savePresetButton: qs("#savePresetButton"),
  loadPresetButton: qs("#loadPresetButton"),
  deletePresetButton: qs("#deletePresetButton"),
  resetImagesButton: qs("#resetImagesButton"),
  clearCustomerButton: qs("#clearCustomerButton"),
  clearQuoteCompanyButton: qs("#clearQuoteCompanyButton"),
  analyseAgainButton: qs("#analyseAgainButton"),
  resetQuoteBasisButton: qs("#resetQuoteBasisButton"),
  resetOutputButton: qs("#resetOutputButton"),
  presetStatus: qs("#presetStatus"),
  sideBackButton: qs("#sideBackButton"),
  sideNextButton: qs("#sideNextButton"),
  sideDownloadButton: qs("#sideDownloadButton"),
  basisChatOverlay: qs("#basisChatOverlay"),
  basisChatTitle: qs("#basisChatTitle"),
  basisChatContext: qs("#basisChatContext"),
  basisChatMessages: qs("#basisChatMessages"),
  basisChatProposal: qs("#basisChatProposal"),
  basisChatProposalActions: qs("#basisChatProposalActions"),
  basisChatForm: qs("#basisChatForm"),
  basisChatPrompt: qs("#basisChatPrompt"),
  basisChatSendButton: qs("#basisChatSendButton"),
  basisChatApplyButton: qs("#basisChatApplyButton"),
  basisChatKeepButton: qs("#basisChatKeepButton"),
  basisChatCloseButton: qs("#basisChatCloseButton"),
  newPricingReferenceButton: qs("#newPricingReferenceButton"),
  deletePricingReferenceButton: qs("#deletePricingReferenceButton"),
  pricingReferenceModal: qs("#pricingReferenceModal"),
  pricingReferenceForm: qs("#pricingReferenceForm"),
  pricingReferenceName: qs("#pricingReferenceName"),
  pricingReferenceTemplateButton: qs("#pricingReferenceTemplateButton"),
  pricingReferenceFile: qs("#pricingReferenceFile"),
  pricingReferenceFileName: qs("#pricingReferenceFileName"),
  pricingReferencePreview: qs("#pricingReferencePreview"),
  pricingReferenceNoAccess: qs("#pricingReferenceNoAccess"),
  pricingReferenceEditorBody: qs("#pricingReferenceEditorBody"),
  pricingReferenceDeleteSection: qs("#pricingReferenceDeleteSection"),
  deletePricingReferenceSelect: qs("#deletePricingReferenceSelect"),
  pricingReferenceTableOverlay: qs("#pricingReferenceTableOverlay"),
  pricingReferenceTableSummary: qs("#pricingReferenceTableSummary"),
  pricingReferenceTableBody: qs("#pricingReferenceTableBody"),
  pricingReferenceTableCloseButton: qs("#pricingReferenceTableCloseButton"),
  settingsButton: qs("#settingsButton"),
  pricingReferenceTaxLabel: qs("#pricingReferenceTaxLabel"),
  pricingReferenceTaxRate: qs("#pricingReferenceTaxRate"),
  pricingReferenceCurrency: qs("#pricingReferenceCurrency"),
  pricingReferenceCurrencyCustom: qs("#pricingReferenceCurrencyCustom"),
  selectedPricingReferenceSummary: qs("#selectedPricingReferenceSummary"),
  selectedPricingReferenceCurrency: qs("#selectedPricingReferenceCurrency"),
  selectedPricingReferenceTax: qs("#selectedPricingReferenceTax"),
  outputSortMode: qs("#outputSortMode"),
  pricingReferenceSaveButton: qs("#pricingReferenceSaveButton"),
  pricingReferenceCancelButton: qs("#pricingReferenceCancelButton"),
  pricingReferenceCloseButton: qs("#pricingReferenceCloseButton"),
  analysisConfirmModal: qs("#analysisConfirmModal"),
  analysisConfirmCancelButton: qs("#analysisConfirmCancelButton"),
  analysisConfirmStartButton: qs("#analysisConfirmStartButton"),
  richTextEditors: Array.from(document.querySelectorAll("[data-rich-text-source]")),
  richTextToolbar: Array.from(document.querySelectorAll("[data-rich-command]")),
};

function normalizeAnalysisMode(value) {
  void value;
  return "standard";
}

function analysisRunningMessage() {
  return "Reading the reference files and preparing the quote basis. Please wait.";
}

function pricingReferenceSelectValue(reference = {}) {
  const referenceId = String(reference.id || "").trim();
  if (!referenceId) return "";
  const source = reference.source === "company" ? "company" : reference.source === "local" ? "local" : "bundled";
  return `${source}::${referenceId}`;
}

function pricingReferenceSelectionFromValue(value = "") {
  const text = String(value || "").trim();
  if (!text) return { pricingReferenceId: "", source: "" };
  const delimiterIndex = text.indexOf("::");
  if (delimiterIndex < 0) {
    return { pricingReferenceId: text, source: "" };
  }
  return {
    source: text.slice(0, delimiterIndex) || "bundled",
    pricingReferenceId: text.slice(delimiterIndex + 2),
  };
}

function mergePricingReferences(bundled = []) {
  const seen = new Set();
  return [...bundled].filter((reference) => {
    const source = String(reference?.source || "bundled").trim() || "bundled";
    if (source !== "bundled") return false;
    const key = pricingReferenceSelectValue(reference);
    if (!key || seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

function sortedPricingReferencesForDisplay(references = []) {
  return [...references].sort((left, right) => {
    const leftLabel = String(left?.label || left?.id || "").trim();
    const rightLabel = String(right?.label || right?.id || "").trim();
    return leftLabel.localeCompare(rightLabel, undefined, { sensitivity: "base" })
      || String(left?.id || "").localeCompare(String(right?.id || ""), undefined, { sensitivity: "base" });
  });
}

function safeId(value = "", fallback = "item") {
  const slug = String(value || "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "");
  return slug || fallback;
}

function basisDisplayTitle(value = "") {
  return String(value || "")
    .replace(/\u2013|\u2014/g, "-")
    .replace(/\s*-\s*quote\s+basis\s+to\s+confirm\s*$/i, "")
    .trim();
}

function normalizeCategoryTitle(value = "") {
  const text = basisDisplayTitle(value);
  const compact = text.toLowerCase().replace(/[^a-z0-9]+/g, " ").trim().replace(/\s+/g, " ");
  if (!compact) return "General";
  if (/\b(counter|counters|cabinet|cabinets|cabinetry|reception counter|storage cabinet|display counter|lockable cabinet)\b/.test(compact)) return "COUNTERS AND CABINETS";
  if (/\b(rigging|overhead structure|aluminium box truss|aluminum box truss|box truss|truss|suspended structure|hanging frame|hanging structure)\b/.test(compact)) return "Hanging Structure";
  if (/\b(platform|flooring|floor|carpet|raised platform|vinyl flooring)\b/.test(compact)) return "Floor Design";
  if (/\b(graphic|graphics|signage|sign|logo|lightbox|print|printed|vinyl)\b/.test(compact)) return "Graphics";
  if (/\b(furniture|chair|table)\b/.test(compact)) return "Furniture Rental";
  if (/\b(decor|plant|plants|rental item|loose rental)\b/.test(compact)) return "Rental Items";
  if (/\b(electrical|power|socket|light|lighting|spotlight)\b/.test(compact)) return "Electrical Fittings ( Excluding connection fees by Organiser)";
  if (/\b(av|audio|visual|screen|monitor|tv)\b/.test(compact)) return "AV Equipment Rental Items";
  if (/\b(water|sink|tap|plumbing)\b/.test(compact)) return "Water Connection";
  if (/\b(coffee|tea|beverage|drink)\b/.test(compact)) return "Coffee / Tea (Subject to approval by Venue owner and Organiser)";
  const boothTerms = new Set(["booth", "booth dimension", "booth dimensions", "booth structure", "booth structures", "structure", "structures", "walls", "wall", "partitions", "partition", "fascia", "system profile partitions", "entrance frame", "curved frame", "back wall", "side wall", "header fascia structure", "header"]);
  if (boothTerms.has(compact) || /\b(booth|wall|walls|partition|partitions|fascia|entrance frame|curved frame|back wall|side wall|system profile|header)\b/.test(compact)) return "Booth Structure";
  return text;
}

function sectionTitleKey(value = "") {
  return basisDisplayTitle(value).toLowerCase().replace(/[^a-z0-9]+/g, " ").trim();
}

function referenceSectionTitleAliases(value = "") {
  const title = basisDisplayTitle(value);
  if (!title) return [];
  const aliasesByTitle = {
    "floor design": ["Flooring / Platform", "Platform / Flooring", "Flooring", "Platform"],
    "counters and cabinets": ["Counters & Cabinets", "Counters", "Cabinets"],
    "graphics": ["Graphics / Signage", "Signage"],
    "furniture rental": ["Furniture / Decor", "Furniture", "Furniture Decor"],
    "rental items": ["Decor", "Plant", "Plants", "Loose Rental", "Rental Item"],
    "av equipment rental items": ["Electrical / AV", "AV", "Audio Visual", "Screens", "Monitor", "TV"],
    "electrical fittings excluding connection fees by organiser": ["Electrical / AV", "Electrical", "Lighting", "Lights", "Power"],
    "booth structure": ["Surfaces / Structures", "Walls / Structures", "Walls", "Partitions", "Fascia"],
  };
  const aliases = new Set([title, normalizeCategoryTitle(title)]);
  (aliasesByTitle[sectionTitleKey(title)] || []).forEach((alias) => aliases.add(alias));
  return Array.from(aliases).filter((alias) => basisDisplayTitle(alias));
}

function exactPricingReferenceSectionTitle(value = "") {
  const text = basisDisplayTitle(value);
  if (!text) return "";
  const appState = typeof state === "object" && state ? state : {};
  const references = Array.isArray(appState.pricingReferences) ? appState.pricingReferences : [];
  const selectedId = String(appState.pricingReferenceId || "").trim();
  const selectedSource = String(appState.pricingReferenceSource || "").trim();
  const reference = references.find((item) => String(item?.id || "") === selectedId && (!selectedSource || String(item?.source || "") === selectedSource))
    || references.find((item) => String(item?.id || "") === selectedId)
    || null;
  const items = Array.isArray(reference?.items) ? reference.items : [];
  const lookup = new Map();
  items.forEach((item) => {
    const title = basisDisplayTitle(item?.reference_section || item?.source_section || item?.section || "");
    if (!title) return;
    referenceSectionTitleAliases(title).forEach((candidate) => {
      const key = sectionTitleKey(candidate);
      if (key && !lookup.has(key)) lookup.set(key, title);
    });
  });
  return referenceSectionTitleAliases(text).map(sectionTitleKey).map((key) => lookup.get(key)).find(Boolean) || "";
}

function normalizeQuoteBasisTitle(value = "") {
  return exactPricingReferenceSectionTitle(value) || normalizeCategoryTitle(value);
}

function cleanCustomerQuoteLineText(value = "") {
  let text = normalizeUnit(String(value || "").trim().replace(/\s+/g, " "));
  text = text.replace(/(^|[^A-Za-z0-9])m\s*(?:2|\^2)(?=$|[^A-Za-z0-9])/gi, "$1sqm");
  text = text.replace(/(^|[^A-Za-z0-9])sq\.?\s*m\.?(?=$|[^A-Za-z0-9])/gi, "$1sqm");
  const replacements = [
    /\s+taken\s+from\s+quotation\s+title\s*:?\s*/gi,
    /\s+from\s+quotation\s+title\s*:?\s*/gi,
    /\s+visible\s+in\s+image(?:s)?\s*(?:at\s+)?/gi,
    /\s+as\s+seen\s+in\s+render\s*/gi,
    /\s+as\s+per\s+image\s*/gi,
    /\s+based\s+on\s+uploaded\s+reference\s*/gi,
    /\s+ai\s+detected\s*:?\s*/gi,
    /\s+assumed\s+from\s+image\s*/gi,
    /\s+suggested\s+by\s+image\s*/gi,
    /\s+from\s+reference\s+image\s*/gi,
    /\s+appears\s+to\s+be\s+/gi,
    /\s+likely\s+/gi,
  ];
  replacements.forEach((pattern) => { text = text.replace(pattern, " "); });
  return text.replace(/\s+(:|,|\.)/g, "$1").replace(/:\s*([0-9])/g, " $1").replace(/\s{2,}/g, " ").replace(/^[ ;]+|[ ;]+$/g, "");
}

function pricingReferenceLineText(value = "") {
  return String(value || "").trim().replace(/\s+/g, " ");
}

function bracketedCatalogReferenceParts(value = "") {
  const match = cleanCustomerQuoteLineText(value).match(/^\[\s*(.+?)\s*\](?:\s*-\s*(.*))?$/);
  if (!match) return null;
  const reference = cleanCustomerQuoteLineText(match[1] || "");
  const detail = cleanCustomerQuoteLineText(match[2] || "");
  return reference ? { reference, detail } : null;
}

function outputCatalogDescription(value = {}) {
  const reference = pricingReferenceLineText(value.pricing_reference_description || value.catalog_description || "");
  if (reference) return cleanCustomerQuoteLineText(reference);
  if (value.pricing_keyword) {
    const bracketed = bracketedCatalogReferenceParts(value.text || value.description || "");
    if (bracketed?.reference) return bracketed.reference;
  }
  return cleanCustomerQuoteLineText(value.text || value.description || "");
}

function leadingNumber(value = "") {
  const match = String(value ?? "").replaceAll(",", "").trim().match(/^-?\d+(?:\.\d+)?/);
  if (!match) return null;
  const number = Number(match[0]);
  return Number.isFinite(number) ? number : null;
}

function formatQuantityNumber(value) {
  return Number.isInteger(value) ? value : value;
}

function normalizeQuantityPrefixUnit(value = "") {
  const unit = String(value || "").toLowerCase().replace(/\s+/g, " ").replace(/\.+$/g, "").trim();
  if (["sqm", "m2", "m^2", "sq m", "sq.m", "sq.m."].includes(unit)) return "sqm";
  if (["nos", "no", "pc", "pcs", "piece", "pieces", "unit", "units"].includes(unit)) return "nos";
  if (["lot", "lots"].includes(unit)) return "lot";
  if (["set", "sets"].includes(unit)) return "set";
  if (["each", "ea"].includes(unit)) return "each";
  if (["m length", "m run"].includes(unit)) return "m length";
  return normalizeUnit(value);
}

function leadingQuantityPrefix(value = "") {
  const cleaned = cleanCustomerQuoteLineText(value);
  const match = cleaned.match(/^(\d[\d,]*(?:\.\d+)?)\s+(m\s+(?:length|run)|sq\.?\s*m\.?|m\s*(?:2|\^2)|sqm|nos?\.?|no\.?|pcs?\.?|pieces?|units?|lots?|sets?|each|ea)(?=$|[\s:;,\-.])/i);
  if (!match) return null;
  const quantity = leadingNumber(match[1]);
  if (quantity === null || quantity <= 0) return null;
  let text = cleaned.slice(match[0].length).replace(/^[\s:;,\-.]+/, "").replace(/^of\s+/i, "").trim();
  if (!text) return null;
  return {
    text,
    quantity: formatQuantityNumber(quantity),
    unit: normalizeQuantityPrefixUnit(match[2]),
  };
}

function quantityUnitAliases(unit = "") {
  const normalized = normalizeUnit(unit).toLowerCase().replace(/\s+/g, " ").trim();
  const aliases = new Set();
  if (normalized) aliases.add(normalized);
  if (normalized === "sqm") {
    ["sqm", "m2", "m^2", "sq m", "sq.m", "sq.m."].forEach((alias) => aliases.add(alias));
  }
  if (/^nos?\.?$/.test(normalized)) {
    ["nos", "nos.", "no", "no.", "pcs", "pieces", "units"].forEach((alias) => aliases.add(alias));
  }
  if (normalized === "m length") {
    ["m length", "m"].forEach((alias) => aliases.add(alias));
  }
  return Array.from(aliases).filter(Boolean).sort((a, b) => b.length - a.length);
}

function startsWithQuantityUnit(textAfterNumber = "", unit = "") {
  const after = String(textAfterNumber || "").trimStart().toLowerCase();
  if (!after) return true;
  if (!/^[a-z]/.test(after)) return true;
  return quantityUnitAliases(unit).some((alias) => {
    const escaped = alias.replace(/[.*+?^${}()|[\]\\]/g, "\\$&").replace(/\\ /g, "\\s+");
    return new RegExp(`^${escaped}(?=$|[^a-z0-9])`, "i").test(after);
  });
}

function stripLeadingQuantityCountFromLineText(text = "", quantity = "", unit = "") {
  const cleaned = cleanCustomerQuoteLineText(text);
  const prefix = leadingQuantityPrefix(cleaned);
  if (prefix) {
    const quantityNumber = leadingNumber(quantity);
    if (quantityNumber === null || Math.abs(Number(prefix.quantity) - quantityNumber) <= 0.0001) {
      return prefix.text;
    }
  }
  const quantityNumber = leadingNumber(quantity);
  if (quantityNumber === null || quantityNumber <= 0 || !cleaned) return cleaned;
  const match = cleaned.replaceAll(",", "").match(/^-?\d+(?:\.\d+)?/);
  if (!match) return cleaned;
  const textNumber = Number(match[0]);
  if (!Number.isFinite(textNumber) || Math.abs(textNumber - quantityNumber) > 0.0001) return cleaned;
  const originalMatch = cleaned.match(/^-?\d[\d,]*(?:\.\d+)?/);
  if (!originalMatch) return cleaned;
  const remainder = cleaned.slice(originalMatch[0].length);
  if (!startsWithQuantityUnit(remainder, unit)) return cleaned;
  return remainder.replace(/^[\s:;,\-]+/, "").trim() || cleaned;
}

function normalizedLineTextQuantityParts(text = "", quantity = "", unit = "") {
  const prefix = leadingQuantityPrefix(text);
  if (prefix) return { ...prefix, fromTextPrefix: true };
  return {
    text: stripLeadingQuantityCountFromLineText(text, quantity, unit),
    quantity,
    unit: normalizeUnit(unit),
    fromTextPrefix: false,
  };
}

function neutralizeFormulaText(value = "") {
  const text = String(value || "").trim();
  return /^[=+\-@]/.test(text) ? `'${text}` : text;
}

function normalizeUnit(value = "") {
  const text = String(value || "").trim();
  const lower = text.toLowerCase().replace(/\s+/g, " ").replace(/^[.\s]+|[.\s]+$/g, "");
  if (["m2", "m^2", "sq m", "sq.m", "sq.m.", "square metre", "square meter", "square metres", "square meters"].includes(lower)) {
    return "sqm";
  }
  if (["m run", "m. run"].includes(lower)) return "m run";
  if (["m length", "m. length"].includes(lower)) return "m length";
  if (["nos", "no", "pc", "pcs", "piece", "pieces", "unit", "units"].includes(lower)) return "nos";
  if (["lot", "lots"].includes(lower)) return "lot";
  if (["set", "sets"].includes(lower)) return "sets";
  return text;
}

function currentProfile() {
  const resolvedProfileId = String(state.profileId || DEFAULT_PROFILE_ID).trim() || DEFAULT_PROFILE_ID;
  return state.profiles.find((profile) => profile.id === resolvedProfileId)
    || state.profiles.find((profile) => profile.id === DEFAULT_PROFILE_ID)
    || state.profiles[0]
    || {
    id: DEFAULT_PROFILE_ID,
    label: "Quotation Profile",
    description: "",
    quote_detail_presets: [],
  };
}

function defaultPricingReference() {
  const profile = currentProfile();
  const defaultReferenceId = String(profile?.default_pricing_reference || state.defaultPricingReferenceId || DEFAULT_PRICING_REFERENCE_ID).trim();
  return state.pricingReferences.find((reference) => reference.id === defaultReferenceId)
    || state.pricingReferences.find((reference) => reference.id === DEFAULT_PRICING_REFERENCE_ID)
    || state.pricingReferences[0]
    || null;
}

function currentPricingReference() {
  const pricingReferenceId = String(state.pricingReferenceId || "").trim();
  const source = String(state.pricingReferenceSource || "").trim();
  if (!pricingReferenceId) return null;
  return state.pricingReferences.find((reference) => reference.id === pricingReferenceId && (!source || pricingReferenceSelectValue(reference).startsWith(`${source}::`)))
    || state.pricingReferences.find((reference) => reference.id === pricingReferenceId && !source)
    || null;
}

function resolvedProfileIdForPayload() {
  return String(state.profileId || DEFAULT_PROFILE_ID).trim() || DEFAULT_PROFILE_ID;
}

function syncSelectedPricingReference() {
  const selectedReference = currentPricingReference();
  if (selectedReference) {
    state.pricingReferenceId = selectedReference.id || "";
    state.pricingReferenceSource = pricingReferenceSelectionFromValue(pricingReferenceSelectValue(selectedReference)).source;
    return;
  }
  const fallbackReference = defaultPricingReference();
  if (fallbackReference) {
    state.pricingReferenceId = fallbackReference.id || "";
    state.pricingReferenceSource = pricingReferenceSelectionFromValue(pricingReferenceSelectValue(fallbackReference)).source;
    return;
  }
  state.pricingReferenceId = "";
  state.pricingReferenceSource = "";
}

function currentGenerator() {
  const pricingReference = currentPricingReference();
  return {
    ...QUOTE_COPY,
    label: pricingReference?.label || QUOTE_COPY.label,
  };
}

function updateGeneratorCopy() {
  const generator = currentGenerator();
  const atImageLimit = state.images.length >= MAX_REFERENCE_IMAGES;
  if (state.activeSidePanel === "images") elements.sideDrawerSubtitle.textContent = generator.intakeSubtitle;
  elements.dropTitle.textContent = atImageLimit
    ? "Maximum reference files added"
    : state.images.length ? "Add more reference files" : generator.dropTitle;
  elements.dropMeta.textContent = atImageLimit
    ? `Remove one reference to add another. Maximum ${MAX_REFERENCE_IMAGES} reference files.`
    : generator.dropMeta;
}

function setWorkflowStage(stage) {
  state.workflowStage = stage;
  document.body.dataset.workflowStage = stage;
}

function parseTimestampMs(value) {
  const parsed = Date.parse(value || "");
  return Number.isFinite(parsed) ? parsed : 0;
}

function formatElapsedDuration(elapsedMs) {
  const totalSeconds = Math.max(0, Math.floor(elapsedMs / 1000));
  const seconds = String(totalSeconds % 60).padStart(2, "0");
  const minutesTotal = Math.floor(totalSeconds / 60);
  if (minutesTotal < 60) return `${minutesTotal}:${seconds}`;
  const minutes = String(minutesTotal % 60).padStart(2, "0");
  return `${Math.floor(minutesTotal / 60)}:${minutes}:${seconds}`;
}

function activeJobStartedAt(job = state.activeJob) {
  return job?.startedAt || job?.created_at || job?.createdAt || "";
}

function stopAnalysisElapsedTimer() {
  if (!analysisElapsedTimerId) return;
  window.clearInterval(analysisElapsedTimerId);
  analysisElapsedTimerId = 0;
}

function updateAnalysisElapsed(startedAt = activeJobStartedAt()) {
  const elapsed = qs("#analysisElapsed");
  if (!elapsed) return;
  const startedMs = parseTimestampMs(startedAt);
  elapsed.textContent = `Elapsed ${formatElapsedDuration(startedMs ? Date.now() - startedMs : 0)}`;
}

function startAnalysisElapsedTimer(startedAt = activeJobStartedAt()) {
  stopAnalysisElapsedTimer();
  updateAnalysisElapsed(startedAt);
  analysisElapsedTimerId = window.setInterval(() => updateAnalysisElapsed(startedAt), 1000);
}

const GENERIC_FAILURE_MESSAGE = "Failed. Please try again. Contact support if this keeps happening.";

function errorReferenceFrom(data = {}) {
  if (!data || typeof data !== "object") return "";
  return String(
    data.error_reference
    || data.errorReference
    || data.result?.error_reference
    || data.result?.errorReference
    || ""
  ).trim();
}

function genericFailureMessage(data = {}) {
  const reference = typeof data === "string" ? "" : errorReferenceFrom(data);
  return reference ? `${GENERIC_FAILURE_MESSAGE} Reference: ${reference}.` : GENERIC_FAILURE_MESSAGE;
}

function genericFailureMessages(data = {}) {
  return [genericFailureMessage(data)];
}

function setAiStatusBanner(tone, title, message, options = {}) {
  if (tone !== "running") stopAnalysisElapsedTimer();
  const elapsedMarkup = options.elapsed
    ? '<span class="ai-elapsed" id="analysisElapsed" aria-live="polite">Elapsed 0:00</span>'
    : "";
  elements.aiFailureBanner.hidden = false;
  elements.aiFailureBanner.classList.toggle("is-running", tone === "running");
  elements.aiFailureBanner.classList.toggle("is-failure", tone !== "running");
  elements.aiFailureBanner.innerHTML = `
    <strong>${escapeHtml(title)}</strong>
    <span class="ai-status-message">${escapeHtml(message)}</span>
    ${elapsedMarkup}
  `;
}

function showAiRunningBanner(message = "Reading the reference files and preparing the quote basis. Please wait.", startedAt = activeJobStartedAt()) {
  setAiStatusBanner("running", "AI analysis running.", message, { elapsed: true });
  startAnalysisElapsedTimer(startedAt);
}

function showAiFailureBanner(message = GENERIC_FAILURE_MESSAGE) {
  const detail = String(message || "").replace(/^AI analysis failed\.?\s*/i, "").trim() || GENERIC_FAILURE_MESSAGE;
  setAiStatusBanner("failure", "AI analysis failed.", detail);
}

function showAiBlockedBanner(message = "Complete the required details, then retry analysis.") {
  const detail = String(message || "").replace(/^AI analysis blocked\.?\s*/i, "").trim() || "Complete the required details, then retry analysis.";
  setAiStatusBanner("failure", "AI analysis blocked.", detail);
}

function showBlockedAction(message = "Complete the required details, then try again.", options = {}) {
  showAiBlockedBanner(message);
  if (options.details !== false && missingDetailFields().length) setDetailsDrawer(true);
}

function clearAiFailureBanner() {
  stopAnalysisElapsedTimer();
  elements.aiFailureBanner.hidden = true;
  elements.aiFailureBanner.classList.remove("is-running", "is-failure");
  elements.aiFailureBanner.innerHTML = `
    <strong>AI analysis failed.</strong>
    <span>${GENERIC_FAILURE_MESSAGE}</span>
  `;
}

function setDetailsDrawer(open) {
  if (open) {
    setSidePanel(firstMissingDetailsPanel(), { force: true });
  }
}

function formatBytes(bytes) {
  if (!bytes) return "0 KB";
  if (bytes < 1024 * 1024) return `${Math.max(1, Math.round(bytes / 1024))} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

function fileToDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(reader.error);
    reader.readAsDataURL(file);
  });
}

function isAcceptedReferenceFile(file = {}) {
  const type = String(file.type || "").toLowerCase();
  const name = String(file.name || "").toLowerCase();
  return type.startsWith("image/") || type === "application/pdf" || name.endsWith(".pdf");
}

function referenceFileType(entry = {}) {
  const match = String(entry.data_url || "").match(/^data:([^;,]+)/i);
  if (match) return match[1].toLowerCase();
  const type = String(entry.type || "").toLowerCase();
  if (type) return type;
  return String(entry.name || "").toLowerCase().endsWith(".pdf") ? "application/pdf" : "image";
}

function isPdfReference(entry = {}) {
  return referenceFileType(entry) === "application/pdf";
}

function referenceFileTypeLabel(entry = {}) {
  return isPdfReference(entry) ? "PDF" : (entry.type || "image");
}

async function filesToImageEntries(files, limit = MAX_REFERENCE_IMAGES) {
  return Promise.all(
    Array.from(files)
      .filter(isAcceptedReferenceFile)
      .slice(0, limit)
      .map(async (file) => {
        const dataUrl = await fileToDataUrl(file);
        return {
          name: file.name,
          type: referenceFileType({ name: file.name, type: file.type, data_url: dataUrl }),
          size: file.size,
          data_url: dataUrl,
        };
      })
  );
}

function imageDuplicateKey(image = {}) {
  const dataUrl = String(image.data_url || "").trim();
  if (dataUrl) return `data:${dataUrl}`;
  const name = String(image.name || "").trim().toLowerCase();
  const type = String(image.type || "").trim().toLowerCase();
  const size = String(image.size || "").trim();
  return name || type || size ? `file:${name}:${type}:${size}` : "";
}

function uniqueImageEntries(nextImages = [], existingImages = state.images) {
  const seen = new Set(existingImages.map(imageDuplicateKey).filter(Boolean));
  const unique = [];
  let duplicateCount = 0;
  nextImages.forEach((image) => {
    const key = imageDuplicateKey(image);
    if (key && seen.has(key)) {
      duplicateCount += 1;
      return;
    }
    if (key) seen.add(key);
    unique.push(image);
  });
  return { unique, duplicateCount };
}

function imageCapacity(existingImages = state.images) {
  return Math.max(0, MAX_REFERENCE_IMAGES - existingImages.length);
}

function setImageUploadStatus(message = "") {
  if (elements.imageUploadStatus) elements.imageUploadStatus.textContent = message;
}

async function addImagesFromFiles(files) {
  const referenceFiles = Array.from(files).filter(isAcceptedReferenceFile);
  if (!referenceFiles.length) {
    if (Array.from(files).length) setImageUploadStatus("Only JPG, PNG, WebP, or PDF files can be added.");
    return;
  }
  const capacity = imageCapacity();
  if (!capacity) {
    setImageUploadStatus(`Maximum ${MAX_REFERENCE_IMAGES} reference files reached. Remove one before adding more.`);
    return;
  }
  const overflowCount = Math.max(0, referenceFiles.length - capacity);
  const images = await filesToImageEntries(referenceFiles, capacity);
  if (!images.length) return;
  const { unique, duplicateCount } = uniqueImageEntries(images);
  if (!unique.length) {
    const duplicateMessage = duplicateCount ? `${duplicateCount} duplicate file${duplicateCount === 1 ? "" : "s"} skipped.` : "";
    const overflowMessage = overflowCount ? `${overflowCount} file${overflowCount === 1 ? "" : "s"} skipped because the maximum is ${MAX_REFERENCE_IMAGES}.` : "";
    setImageUploadStatus([duplicateMessage, overflowMessage].filter(Boolean).join(" "));
    return;
  }
  state.images = [...state.images, ...unique];
  await persistSessionFiles(sessionFileRecordsFromImages(state.images)).catch(() => {});
  renderFiles();
  setWorkflowStage("ready_to_analyze");
  const generator = currentGenerator();
  const duplicateMessage = duplicateCount ? ` ${duplicateCount} duplicate file${duplicateCount === 1 ? "" : "s"} skipped.` : "";
  const overflowMessage = overflowCount ? ` ${overflowCount} file${overflowCount === 1 ? "" : "s"} skipped because the maximum is ${MAX_REFERENCE_IMAGES}.` : "";
  setImageUploadStatus(`${unique.length} new ${generator.imageNoun}${unique.length === 1 ? "" : "s"} added.${duplicateMessage}${overflowMessage}`);
  syncControlStates();
}

function removeImageAt(index) {
  state.images.splice(index, 1);
  renderFiles();
  if (!state.images.length) {
    setWorkflowStage("needs_images");
    setImageUploadStatus("");
  } else {
    setImageUploadStatus(`${state.images.length} reference file${state.images.length === 1 ? "" : "s"} loaded.`);
  }
  syncControlStates();
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function decodeHtmlEntities(value) {
  return String(value ?? "").replace(/&(#x[0-9a-f]+|#\d+|amp|lt|gt|quot|apos|nbsp);/gi, (match, entity) => {
    const normalized = entity.toLowerCase();
    if (normalized === "amp") return "&";
    if (normalized === "lt") return "<";
    if (normalized === "gt") return ">";
    if (normalized === "quot") return '"';
    if (normalized === "apos") return "'";
    if (normalized === "nbsp") return " ";
    if (normalized.startsWith("#x")) {
      const codePoint = Number.parseInt(normalized.slice(2), 16);
      return Number.isFinite(codePoint) ? String.fromCodePoint(codePoint) : match;
    }
    if (normalized.startsWith("#")) {
      const codePoint = Number.parseInt(normalized.slice(1), 10);
      return Number.isFinite(codePoint) ? String.fromCodePoint(codePoint) : match;
    }
    return match;
  });
}

function sanitizeRichTextHtml(value = "") {
  const allowedTags = new Set(["div", "p", "br", "strong", "b", "em", "i", "u"]);
  const droppedContentTags = new Set(["script", "style", "iframe", "object", "embed", "svg", "math"]);
  const droppedVoidTags = new Set(["img"]);
  const tokens = String(value || "").match(/<!--[\s\S]*?-->|<![^>]*>|<\/?[A-Za-z][A-Za-z0-9:-]*\b[^>]*>|[^<]+|</g) || [];
  const droppedStack = [];
  let output = "";

  tokens.forEach((token) => {
    const tagMatch = token.match(/^<\/?([A-Za-z][A-Za-z0-9:-]*)\b[^>]*>$/);
    if (!tagMatch) {
      if (!droppedStack.length && !token.startsWith("<!")) {
        output += escapeHtml(decodeHtmlEntities(token));
      }
      return;
    }

    const tag = tagMatch[1].toLowerCase();
    const isClosing = token.startsWith("</");
    const isSelfClosing = /\/\s*>$/.test(token) || tag === "br";

    if (droppedVoidTags.has(tag)) return;
    if (droppedContentTags.has(tag)) {
      if (!isClosing && !isSelfClosing) droppedStack.push(tag);
      if (isClosing) {
        const lastIndex = droppedStack.lastIndexOf(tag);
        if (lastIndex >= 0) droppedStack.splice(lastIndex, 1);
      }
      return;
    }
    if (droppedStack.length) return;
    if (!allowedTags.has(tag)) return;
    if (tag === "br") {
      output += "<br>";
      return;
    }
    output += isClosing ? `</${tag}>` : `<${tag}>`;
  });

  return output;
}

function normalizeTextNewlines(value) {
  return String(value || "")
    .replace(/\r\n/g, "\n")
    .replace(/\r/g, "\n")
    .replace(/\\n/g, "\n");
}

function renderInlineMarkdown(value) {
  let html = escapeHtml(value);
  html = html.replace(/`([^`]+)`/g, "<code>$1</code>");
  html = html.replace(/\*\*([^*]+)\*\*/g, "<strong>$1</strong>");
  return html;
}

function isMarkdownTableSeparator(line = "") {
  return /^\|?\s*:?-{3,}:?\s*(\|\s*:?-{3,}:?\s*)+\|?$/.test(String(line || "").trim());
}

function splitMarkdownTableRow(line = "") {
  return String(line || "")
    .trim()
    .replace(/^\|/, "")
    .replace(/\|$/, "")
    .split("|")
    .map((cell) => cell.trim());
}

function renderMarkdownTable(lines) {
  const headers = splitMarkdownTableRow(lines[0]);
  const rows = lines.slice(2).map(splitMarkdownTableRow).filter((row) => row.length);
  if (!headers.length || !rows.length) return `<p>${renderInlineMarkdown(lines[0] || "")}</p>`;
  return `
    <table>
      <thead>
        <tr>${headers.map((cell) => `<th>${renderInlineMarkdown(cell)}</th>`).join("")}</tr>
      </thead>
      <tbody>
        ${rows.map((row) => `
          <tr>${headers.map((_, index) => `<td>${renderInlineMarkdown(row[index] || "")}</td>`).join("")}</tr>
        `).join("")}
      </tbody>
    </table>
  `;
}

function renderPlainText(value) {
  const lines = normalizeTextNewlines(value)
    .split("\n")
    .map((line) => line.trim());
  const chunks = [];
  for (let index = 0; index < lines.length;) {
    const line = lines[index];
    if (!line) {
      index += 1;
      continue;
    }

    if (line.includes("|") && isMarkdownTableSeparator(lines[index + 1] || "")) {
      const tableLines = [line, lines[index + 1]];
      index += 2;
      while (index < lines.length && lines[index] && lines[index].includes("|")) {
        tableLines.push(lines[index]);
        index += 1;
      }
      chunks.push(renderMarkdownTable(tableLines));
      continue;
    }

    if (/^-\s+/.test(line)) {
      const items = [];
      while (index < lines.length && /^-\s+/.test(lines[index])) {
        items.push(lines[index].replace(/^-\s+/, ""));
        index += 1;
      }
      chunks.push(`<ul>${items.map((item) => `<li>${renderInlineMarkdown(item)}</li>`).join("")}</ul>`);
      continue;
    }

    chunks.push(`<p>${renderInlineMarkdown(line)}</p>`);
    index += 1;
  }
  return chunks.length ? chunks.join("") : "<p></p>";
}

function splitLines(value) {
  return normalizeTextNewlines(value)
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);
}

function normalizeTaxLabel(value = "") {
  return String(value || "").trim().toUpperCase() === "VAT" ? "VAT" : DEFAULT_TAX_LABEL;
}

function normalizeTaxRate(value, fallback = DEFAULT_TAX_RATE) {
  const number = Number(String(value ?? "").replace("%", "").trim());
  if (!Number.isFinite(number)) return fallback;
  const rate = number > 1 ? number / 100 : number;
  return Math.min(1, Math.max(0, rate));
}

function taxRatePercentText(value = DEFAULT_TAX_RATE) {
  const percent = normalizeTaxRate(value, DEFAULT_TAX_RATE) * 100;
  return Number.isInteger(percent) ? String(percent) : String(Number(percent.toFixed(2)));
}

function taxRateFromPercentInput(value, fallback = DEFAULT_TAX_RATE) {
  const number = Number(String(value ?? "").replace("%", "").trim());
  if (!Number.isFinite(number)) return fallback;
  return Math.min(1, Math.max(0, number / 100));
}

function normalizeCurrencyLabel(value = DEFAULT_CURRENCY_LABEL) {
  const normalized = String(value || "")
    .trim()
    .toUpperCase()
    .replace(/[^A-Z]/g, "");
  if (["S", "SG", "SGD", "SINGAPOREDOLLAR", "SINGAPOREDOLLARS"].includes(normalized)) return "SGD";
  if (["US", "USD", "USDOLLAR", "USDOLLARS"].includes(normalized)) return "USD";
  if (["EURO", "EUROS", "EUR"].includes(normalized)) return "EUR";
  if (["GBP", "POUND", "POUNDS", "STERLING"].includes(normalized)) return "GBP";
  if (["MYR", "RM", "MALAYSIARINGGIT", "RINGGIT"].includes(normalized)) return "MYR";
  if (["AUD", "AUSTRALIANDOLLAR", "AUSTRALIANDOLLARS"].includes(normalized)) return "AUD";
  if (["CNY", "RMB", "YUAN", "RENMINBI"].includes(normalized)) return "CNY";
  if (["IDR", "RUPIAH"].includes(normalized)) return "IDR";
  if (["THB", "BAHT"].includes(normalized)) return "THB";
  return /^[A-Z]{3}$/.test(normalized) ? normalized : DEFAULT_CURRENCY_LABEL;
}

function selectedPricingReferenceTax() {
  const reference = currentPricingReference();
  return {
    label: normalizeTaxLabel(reference?.tax?.label),
    rate: normalizeTaxRate(reference?.tax?.rate, DEFAULT_TAX_RATE),
  };
}

function selectedPricingReferenceCurrency() {
  const reference = currentPricingReference();
  return normalizeCurrencyLabel(reference?.currency);
}

function supportedCurrencyLabel(value = DEFAULT_CURRENCY_LABEL) {
  const normalized = normalizeCurrencyLabel(value);
  return CURRENCY_OPTIONS.some(([code]) => code === normalized) ? normalized : DEFAULT_CURRENCY_LABEL;
}

function normalizedCustomCurrencyInput() {
  return String(elements.pricingReferenceCurrencyCustom?.value || "")
    .trim()
    .toUpperCase()
    .replace(/[^A-Z]/g, "")
    .slice(0, 3);
}

function customCurrencyInputIsValid() {
  return /^[A-Z]{3}$/.test(normalizedCustomCurrencyInput());
}

function setPricingReferenceCurrencyControls(value = DEFAULT_CURRENCY_LABEL) {
  const normalized = normalizeCurrencyLabel(value);
  const isSupported = CURRENCY_OPTIONS.some(([code]) => code === normalized);
  if (elements.pricingReferenceCurrency) {
    elements.pricingReferenceCurrency.value = isSupported ? normalized : CUSTOM_CURRENCY_VALUE;
  }
  if (elements.pricingReferenceCurrencyCustom) {
    elements.pricingReferenceCurrencyCustom.hidden = isSupported;
    elements.pricingReferenceCurrencyCustom.required = !isSupported;
    elements.pricingReferenceCurrencyCustom.value = isSupported ? "" : normalized;
  }
}

function syncPricingReferenceCurrencyCustomInput() {
  const isCustom = elements.pricingReferenceCurrency?.value === CUSTOM_CURRENCY_VALUE;
  if (!elements.pricingReferenceCurrencyCustom) return;
  elements.pricingReferenceCurrencyCustom.hidden = !isCustom;
  elements.pricingReferenceCurrencyCustom.required = isCustom;
  if (!isCustom) elements.pricingReferenceCurrencyCustom.value = "";
}

function collectTaxDetails() {
  return selectedPricingReferenceTax();
}

function renderSelectedPricingReferenceSummary() {
  const reference = currentPricingReference();
  const tax = selectedPricingReferenceTax();
  const currency = selectedPricingReferenceCurrency();
  const taxText = `${tax.label} ${taxRatePercentText(tax.rate)}%`;
  if (elements.selectedPricingReferenceSummary) {
    elements.selectedPricingReferenceSummary.textContent = reference
      ? "Managed in Settings."
      : "Select a pricing reference.";
  }
  if (elements.selectedPricingReferenceCurrency) elements.selectedPricingReferenceCurrency.textContent = currency;
  if (elements.selectedPricingReferenceTax) elements.selectedPricingReferenceTax.textContent = taxText;
  if (elements.taxLabel) elements.taxLabel.value = tax.label;
  if (elements.taxRate) elements.taxRate.value = taxRatePercentText(tax.rate);
  updatePricingReferenceDeleteButton();
  updateOutputHeader();
}

function todayDateInputValue(date = new Date()) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function quoteDateDisplayText(value = "") {
  const match = String(value || "").match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) return String(value || "");
  return `${match[3]}/${match[2]}/${match[1]}`;
}

function quoteDateHasFormat() {
  return Boolean(state.quoteDateFormat.bold || state.quoteDateFormat.italic || state.quoteDateFormat.underline);
}

function quoteDateRichTextHtml() {
  const displayText = quoteDateDisplayText(elements.quoteDate.value);
  if (!displayText || !quoteDateHasFormat()) return "";
  let formatted = escapeHtml(displayText);
  if (state.quoteDateFormat.underline) formatted = `<u>${formatted}</u>`;
  if (state.quoteDateFormat.italic) formatted = `<em>${formatted}</em>`;
  if (state.quoteDateFormat.bold) formatted = `<strong>${formatted}</strong>`;
  return `<div>${formatted}</div>`;
}

function applyQuoteDateFormatFromHtml(richHtml = "") {
  const value = String(richHtml || "").toLowerCase();
  state.quoteDateFormat = {
    bold: /<(strong|b)\b/.test(value),
    italic: /<(em|i)\b/.test(value),
    underline: /<u\b/.test(value),
  };
  updateQuoteDateFormatButtons();
}

function updateQuoteDateFormatButtons() {
  elements.quoteDateFormatButtons.forEach((button) => {
    const command = button.dataset.dateFormatCommand || "";
    button.classList.toggle("is-selected", Boolean(state.quoteDateFormat[command]));
    button.setAttribute("aria-pressed", String(Boolean(state.quoteDateFormat[command])));
  });
}

function applyDefaultQuoteDate() {
  if (!elements.quoteDate.value.trim()) {
    setInputValue(elements.quoteDate, todayDateInputValue());
  }
}

function renderFiles() {
  elements.imageIntake.classList.toggle("has-images", Boolean(state.images.length));
  updateGeneratorCopy();
  if (elements.imageInput) elements.imageInput.disabled = state.images.length >= MAX_REFERENCE_IMAGES;
  if (!state.images.length) {
    elements.fileList.innerHTML = "";
    return;
  }
  elements.fileList.innerHTML = state.images
    .map((image, index) => {
      const thumb = isPdfReference(image)
        ? `<span class="file-thumb file-thumb-file" aria-hidden="true">PDF</span>`
        : `<img class="file-thumb" src="${escapeHtml(image.data_url)}" alt="">`;
      return `
        <div class="file-item">
          ${thumb}
          <div>
            <strong>${escapeHtml(image.name)}</strong>
            <span>${escapeHtml(referenceFileTypeLabel(image))} - ${formatBytes(image.size)}</span>
          </div>
          <button class="file-remove" type="button" data-remove-image="${index}" aria-label="Remove ${escapeHtml(image.name)}">x</button>
        </div>
      `;
    })
    .join("");
}

function normalizeLineItem(item = {}) {
  const priceMode = item.price_mode === "Included" || String(item.display_price || "").toLowerCase() === "included"
    ? "Included"
    : "Priced";
  const quantityParts = normalizedLineTextQuantityParts(item.description || "", item.quantity ?? "", item.unit || "");
  return {
    section: normalizeCategoryTitle(item.section || ""),
    quantity: quantityParts.quantity ?? "",
    unit: quantityParts.unit || "",
    description: cleanCustomerQuoteLineText(quantityParts.text || ""),
    pricing_keyword: item.pricing_keyword || "",
    display_price: item.display_price || "",
    price_mode: priceMode,
    unit_price_override: item.unit_price_override ?? "",
    catalog_unit_price: item.catalog_unit_price ?? item.unit_price ?? item.sale_unit_price ?? "",
    catalog_description: cleanCustomerQuoteLineText(item.catalog_description || ""),
    pricing_reference_description: pricingReferenceLineText(item.pricing_reference_description || ""),
    source_basis_line_id: item.source_basis_line_id || "",
    category_order: orderNumber(item.category_order) ?? "",
    item_order: orderNumber(item.item_order) ?? "",
    status: item.status || "",
  };
}

function cloneQuoteBasis(basis = {}) {
  return {
    ...EMPTY_BASIS,
    surfaces: basis.surfaces || "",
    counters: basis.counters || "",
    platform: basis.platform || "",
    graphics: basis.graphics || "",
    furniture: basis.furniture || "",
    electrical: basis.electrical || "",
  };
}

function linesValue(value) {
  return Array.isArray(value) ? value.join("\n") : String(value || "");
}

function hasOwnValue(object, key) {
  return Object.prototype.hasOwnProperty.call(object || {}, key);
}

function shouldApply(object, key, partial) {
  return !partial || hasOwnValue(object, key);
}

function richTextEditorFor(input) {
  if (!input) return null;
  return elements.richTextEditors.find((editor) => editor.dataset.richTextSource === input.id) || null;
}

function richTextPlainHtml(value) {
  const lines = normalizeTextNewlines(value).split("\n");
  if (!lines.length || (lines.length === 1 && !lines[0])) return "<div><br></div>";
  return lines.map((line) => `<div>${line ? escapeHtml(line) : "<br>"}</div>`).join("");
}

function richTextEditorPlainText(editor) {
  const blockTags = new Set(["DIV", "P", "LI"]);
  const readNode = (node) => {
    if (node.nodeType === Node.TEXT_NODE) return node.textContent || "";
    if (node.nodeName === "BR") return "\n";
    const isBlock = blockTags.has(node.nodeName);
    const text = Array.from(node.childNodes).map(readNode).join("");
    return isBlock ? `${text}\n` : text;
  };
  return normalizeTextNewlines(Array.from(editor.childNodes).map(readNode).join("")).replace(/\n+$/g, "");
}

function syncRichTextSource(editor) {
  if (!editor) return;
  const input = qs(`#${editor.dataset.richTextSource}`);
  if (!input) return;
  input.value = richTextEditorPlainText(editor).trimEnd();
}

function syncRichTextSources() {
  elements.richTextEditors.forEach(syncRichTextSource);
}

function syncRichTextEditor(input, richHtml = "") {
  const editor = richTextEditorFor(input);
  if (!editor || document.activeElement === editor) return;
  if (richHtml) {
    editor.innerHTML = sanitizeRichTextHtml(richHtml);
    return;
  }
  editor.innerHTML = richTextPlainHtml(input.value ?? "");
}

function collectRichTextDetails() {
  const details = RICH_TEXT_SOURCE_IDS.reduce((collected, id) => {
    const editor = elements.richTextEditors.find((item) => item.dataset.richTextSource === id);
    if (editor) collected[id] = sanitizeRichTextHtml(editor.innerHTML);
    return collected;
  }, {});
  const quoteDateHtml = quoteDateRichTextHtml();
  if (quoteDateHtml) details.quoteDate = sanitizeRichTextHtml(quoteDateHtml);
  return details;
}

function restoreRichTextDetails(details = {}, options = {}) {
  const partial = Boolean(options.partial);
  RICH_TEXT_SOURCE_IDS.forEach((id) => {
    const input = qs(`#${id}`);
    const editor = elements.richTextEditors.find((item) => item.dataset.richTextSource === id);
    if (!input || !editor) return;
    if (details[id]) {
      editor.innerHTML = sanitizeRichTextHtml(details[id]);
    } else if (!partial) {
      editor.innerHTML = richTextPlainHtml(input.value ?? "");
    } else {
      return;
    }
    syncRichTextSource(editor);
  });
  if (!partial || hasOwnValue(details, "quoteDate")) {
    applyQuoteDateFormatFromHtml(details.quoteDate || "");
  }
}

function hasEditableSelection(editor) {
  const selection = window.getSelection();
  if (!selection || selection.rangeCount === 0 || selection.isCollapsed) return false;
  const range = selection.getRangeAt(0);
  return editor.contains(range.commonAncestorContainer);
}

function runRichTextCommand(editor, command) {
  if (!hasEditableSelection(editor)) return;
  editor.focus();
  if (command === "bold") {
    document.execCommand("bold", false, null);
  } else if (command === "italic") {
    document.execCommand("italic", false, null);
  } else {
    document.execCommand("underline", false, null);
  }
  syncRichTextSource(editor);
  syncControlStates();
}

function wireRichTextEditors() {
  elements.richTextEditors.forEach((editor) => {
    syncRichTextEditor(qs(`#${editor.dataset.richTextSource}`));
    editor.addEventListener("input", () => {
      syncRichTextSource(editor);
      syncControlStates();
    });
    editor.addEventListener("keydown", (event) => {
      if (!(event.ctrlKey || event.metaKey)) return;
      const key = event.key.toLowerCase();
      if (!["b", "i", "u"].includes(key)) return;
      event.preventDefault();
      const command = key === "b" ? "bold" : key === "i" ? "italic" : "underline";
      runRichTextCommand(editor, command);
    });
  });
  elements.richTextToolbar.forEach((button) => {
    button.addEventListener("pointerdown", (event) => {
      event.preventDefault();
    });
    button.addEventListener("click", (event) => {
      if (event.detail > 1) return;
      const field = button.closest(".rich-text-field");
      const editor = field?.querySelector("[data-rich-text-source]");
      if (editor) runRichTextCommand(editor, button.dataset.richCommand || "bold");
    });
  });
  elements.quoteDateFormatButtons.forEach((button) => {
    button.addEventListener("click", () => {
      const command = button.dataset.dateFormatCommand || "";
      if (!Object.prototype.hasOwnProperty.call(state.quoteDateFormat, command)) return;
      state.quoteDateFormat[command] = !state.quoteDateFormat[command];
      updateQuoteDateFormatButtons();
      syncControlStates();
    });
  });
  updateQuoteDateFormatButtons();
}

function setInputValue(input, value) {
  if (!input) return;
  input.value = value ?? "";
  syncRichTextEditor(input);
}

function normalizeBoothDimensions(project = {}) {
  const width = Number(project.booth_width);
  const depth = Number(project.booth_depth);
  if (Number.isFinite(width) && Number.isFinite(depth) && width > 0 && depth > 0) {
    const formattedWidth = Number.isInteger(width) ? String(width) : String(width);
    const formattedDepth = Number.isInteger(depth) ? String(depth) : String(depth);
    return {
      booth_width: formattedWidth,
      booth_depth: formattedDepth,
      booth_size: project.booth_size || `${formattedWidth}m x ${formattedDepth}m`,
      dimension_source: project.dimension_source || "analysis",
    };
  }
  return { ...DEFAULT_BOOTH_DIMENSIONS };
}

function collectQuoteDetails() {
  syncRichTextSources();
  return {
    quote_date: elements.quoteDate.value,
    project_number: elements.projectNumber.value.trim(),
    client: {
      name: elements.clientName.value.trim(),
      attention: elements.clientAttention.value.trim(),
      title: elements.clientTitle.value.trim(),
      address: elements.clientAddress.value,
    },
    project: {
      title: elements.projectTitle.value.trim(),
    },
    company: {
      name: elements.quoteCompanyName.value.trim(),
      header_details: elements.headerDetails.value,
      logo_data_url: state.headerLogo ? state.headerLogo.data_url : "",
      logo_name: state.headerLogo ? state.headerLogo.name : "",
      logo_type: state.headerLogo ? state.headerLogo.type : "",
    },
    tax: collectTaxDetails(),
    quote_text: {
      terms_heading: elements.termsHeading.value.trim(),
      payment_terms: splitLines(elements.paymentTerms.value),
      notes_heading: elements.notesHeading.value.trim(),
      standard_notes: splitLines(elements.standardNotes.value),
      acceptance_text: elements.acceptanceText.value.trim(),
      person_label: elements.personLabel.value.trim(),
      stamp_label: elements.stampLabel.value.trim(),
      date_label: elements.dateLabel.value.trim(),
    },
    signature: {
      koncept_signatory: elements.konceptSignatory.value.trim(),
      koncept_title: elements.konceptTitle.value.trim(),
      koncept_date_label: elements.konceptDateLabel.value.trim(),
    },
    rich_text: collectRichTextDetails(),
  };
}

function applyQuoteDetails(details = {}, options = {}) {
  const client = details.client || {};
  const project = details.project || {};
  const company = details.company || {};
  const quoteText = details.quote_text || {};
  const signature = details.signature || {};
  const tax = details.tax || quoteText.tax || {};
  const partial = Boolean(options.partial);

  if (shouldApply(client, "name", partial)) setInputValue(elements.clientName, client.name);
  if (shouldApply(client, "attention", partial)) setInputValue(elements.clientAttention, client.attention);
  if (shouldApply(client, "title", partial)) setInputValue(elements.clientTitle, client.title);
  if (shouldApply(client, "address", partial)) setInputValue(elements.clientAddress, client.address);
  if (shouldApply(project, "title", partial)) setInputValue(elements.projectTitle, project.title);
  if (!partial || shouldApply(project, "booth_width", true) || shouldApply(project, "booth_depth", true)) {
    state.boothDimensions = normalizeBoothDimensions(project);
  }
  if (shouldApply(details, "quote_date", partial)) setInputValue(elements.quoteDate, details.quote_date);
  if (shouldApply(details, "project_number", partial)) setInputValue(elements.projectNumber, details.project_number);
  if (shouldApply(company, "name", partial)) setInputValue(elements.quoteCompanyName, company.name);
  if (shouldApply(company, "header_details", partial)) setInputValue(elements.headerDetails, company.header_details);
  if (!partial || hasOwnValue(details, "tax") || hasOwnValue(quoteText, "tax") || hasOwnValue(quoteText, "tax_label")) {
    if (elements.taxLabel) elements.taxLabel.value = normalizeTaxLabel(tax.label || quoteText.tax_label || DEFAULT_TAX_LABEL);
  }
  if (!partial || hasOwnValue(details, "tax") || hasOwnValue(quoteText, "tax") || hasOwnValue(quoteText, "tax_rate")) {
    setInputValue(elements.taxRate, taxRatePercentText(tax.rate ?? quoteText.tax_rate ?? DEFAULT_TAX_RATE));
  }
  if (shouldApply(quoteText, "terms_heading", partial)) setInputValue(elements.termsHeading, quoteText.terms_heading);
  if (shouldApply(quoteText, "payment_terms", partial)) setInputValue(elements.paymentTerms, linesValue(quoteText.payment_terms));
  if (shouldApply(quoteText, "notes_heading", partial)) setInputValue(elements.notesHeading, quoteText.notes_heading);
  if (shouldApply(quoteText, "standard_notes", partial)) setInputValue(elements.standardNotes, linesValue(quoteText.standard_notes));
  if (shouldApply(quoteText, "acceptance_text", partial)) setInputValue(elements.acceptanceText, quoteText.acceptance_text);
  if (shouldApply(signature, "koncept_signatory", partial)) setInputValue(elements.konceptSignatory, signature.koncept_signatory);
  if (shouldApply(signature, "koncept_title", partial)) setInputValue(elements.konceptTitle, signature.koncept_title);
  if (shouldApply(signature, "koncept_date_label", partial)) setInputValue(elements.konceptDateLabel, signature.koncept_date_label);
  if (shouldApply(quoteText, "person_label", partial)) setInputValue(elements.personLabel, quoteText.person_label);
  if (shouldApply(quoteText, "stamp_label", partial)) setInputValue(elements.stampLabel, quoteText.stamp_label);
  if (shouldApply(quoteText, "date_label", partial)) setInputValue(elements.dateLabel, quoteText.date_label);
  if (options.includeLogo && company.logo_data_url) {
    state.headerLogo = {
      name: company.logo_name || "preset-logo",
      type: company.logo_type || "image/png",
      size: 0,
      data_url: company.logo_data_url,
    };
  } else if (options.clearLogo) {
    state.headerLogo = null;
  }
  const hasRichText = hasOwnValue(details, "rich_text") && details.rich_text && typeof details.rich_text === "object";
  if (hasRichText || !partial) {
    restoreRichTextDetails(hasRichText ? details.rich_text : {}, { partial: partial && hasRichText });
  }
  applyDefaultQuoteDate();
  renderHeaderLogoPreview();
  renderPresetStatus();
}

function applyDefaultQuoteCompanyFields() {
  if (elements.taxLabel) elements.taxLabel.value = DEFAULT_TAX_LABEL;
  setInputValue(elements.taxRate, taxRatePercentText(DEFAULT_TAX_RATE));
  setInputValue(elements.termsHeading, DEFAULT_TERMS_HEADING);
  setInputValue(elements.notesHeading, DEFAULT_NOTES_HEADING);
  setInputValue(elements.acceptanceText, DEFAULT_ACCEPTANCE_TEXT);
  setInputValue(elements.konceptDateLabel, DEFAULT_DATE_LABEL);
  setInputValue(elements.personLabel, DEFAULT_PERSON_LABEL);
  setInputValue(elements.stampLabel, DEFAULT_STAMP_LABEL);
  setInputValue(elements.dateLabel, DEFAULT_DATE_LABEL);
  restoreRichTextDetails(DEFAULT_QUOTE_COMPANY_RICH_TEXT, { partial: true });
}

function safeSessionJson() {
  try {
    const parsed = JSON.parse(window.localStorage.getItem(QUOTE_SESSION_STORAGE_KEY) || "null");
    return parsed && typeof parsed === "object" ? parsed : null;
  } catch {
    return null;
  }
}

function openSessionFileDb() {
  if (!window.indexedDB) return Promise.reject(new Error("IndexedDB unavailable"));
  if (sessionFileDbPromise) return sessionFileDbPromise;
  sessionFileDbPromise = new Promise((resolve, reject) => {
    const request = window.indexedDB.open(QUOTE_SESSION_FILE_DB_NAME, QUOTE_SESSION_FILE_DB_VERSION);
    request.onupgradeneeded = () => {
      const db = request.result;
      if (!db.objectStoreNames.contains(QUOTE_SESSION_FILE_STORE_NAME)) {
        db.createObjectStore(QUOTE_SESSION_FILE_STORE_NAME, { keyPath: "session_file_key" });
      }
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error || new Error("Could not open session file store"));
    request.onblocked = () => reject(new Error("Session file store is blocked"));
  }).catch((error) => {
    sessionFileDbPromise = null;
    throw error;
  });
  return sessionFileDbPromise;
}

function sessionFileKeyForImage(image = {}, index = 0) {
  const existing = String(image.session_file_key || "").trim();
  if (existing) return existing;
  const namePart = String(image.name || "reference-file")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .slice(0, 60) || "reference-file";
  const sizePart = String(Number.isFinite(Number(image.size)) ? Number(image.size) : 0);
  const key = `${Date.now().toString(36)}-${index}-${sizePart}-${namePart}-${Math.random().toString(36).slice(2, 10)}`;
  image.session_file_key = key;
  return key;
}

function sessionImageMetadata(image = {}, index = 0) {
  const sessionFileKey = sessionFileKeyForImage(image, index);
  return {
    name: String(image.name || `reference-${index + 1}`).trim() || `reference-${index + 1}`,
    type: referenceFileType(image),
    size: Number.isFinite(Number(image.size)) ? Number(image.size) : 0,
    session_file_key: sessionFileKey,
  };
}

function sessionFileRecordsFromImages(images = state.images) {
  return images
    .slice(0, MAX_REFERENCE_IMAGES)
    .map((image, index) => ({
      ...sessionImageMetadata(image, index),
      data_url: String(image.data_url || "").trim(),
    }))
    .filter((record) => record.session_file_key && record.data_url);
}

function persistSessionFiles(records = []) {
  const validRecords = records.filter((record) => record?.session_file_key && record?.data_url);
  if (!validRecords.length) return Promise.resolve();
  return openSessionFileDb().then((db) => new Promise((resolve, reject) => {
    const transaction = db.transaction(QUOTE_SESSION_FILE_STORE_NAME, "readwrite");
    const store = transaction.objectStore(QUOTE_SESSION_FILE_STORE_NAME);
    validRecords.forEach((record) => store.put(record));
    transaction.oncomplete = () => resolve();
    transaction.onerror = () => reject(transaction.error || new Error("Could not persist session files"));
    transaction.onabort = () => reject(transaction.error || new Error("Session file persistence aborted"));
  }));
}

function loadSessionFileMap(keys = []) {
  const uniqueKeys = [...new Set(keys.map((key) => String(key || "").trim()).filter(Boolean))];
  if (!uniqueKeys.length) return Promise.resolve(new Map());
  return openSessionFileDb().then((db) => new Promise((resolve, reject) => {
    const found = new Map();
    const transaction = db.transaction(QUOTE_SESSION_FILE_STORE_NAME, "readonly");
    const store = transaction.objectStore(QUOTE_SESSION_FILE_STORE_NAME);
    uniqueKeys.forEach((key) => {
      const request = store.get(key);
      request.onsuccess = () => {
        if (request.result?.data_url) found.set(key, request.result);
      };
    });
    transaction.oncomplete = () => resolve(found);
    transaction.onerror = () => reject(transaction.error || new Error("Could not load session files"));
    transaction.onabort = () => reject(transaction.error || new Error("Session file loading aborted"));
  }));
}

function clearSessionFiles() {
  return openSessionFileDb().then((db) => new Promise((resolve, reject) => {
    const transaction = db.transaction(QUOTE_SESSION_FILE_STORE_NAME, "readwrite");
    transaction.objectStore(QUOTE_SESSION_FILE_STORE_NAME).clear();
    transaction.oncomplete = () => resolve();
    transaction.onerror = () => reject(transaction.error || new Error("Could not clear session files"));
    transaction.onabort = () => reject(transaction.error || new Error("Session file clearing aborted"));
  }));
}

function buildSessionSnapshot() {
  return {
    version: QUOTE_SESSION_STATE_VERSION,
    savedAt: new Date().toISOString(),
    profileId: state.profileId,
    pricingReferenceId: state.pricingReferenceId,
    pricingReferenceSource: state.pricingReferenceSource,
    selectedPresetValue: state.selectedPresetValue,
    images: state.images.slice(0, MAX_REFERENCE_IMAGES).map(sessionImageMetadata),
    quoteDetails: collectQuoteDetails(),
    workflowStage: state.workflowStage,
    quoteBasis: state.quoteBasis,
    quoteBasisSections: state.quoteBasisSections,
    lineItems: state.lineItems,
    outputRows: state.outputRows,
    originalOutputRows: state.originalOutputRows,
    outputErrors: state.outputErrors,
    outputSortMode: state.outputSortMode,
    analysisFindings: state.analysisFindings,
    blockingClarificationQuestions: state.blockingClarificationQuestions,
    boothDimensions: state.boothDimensions,
    originalAnalysisSnapshot: state.originalAnalysisSnapshot,
    basisConfirmed: state.basisConfirmed,
    aiFailed: state.aiFailed,
    draftSource: state.draftSource,
    lastAnalysisMode: state.lastAnalysisMode,
    activeSidePanel: state.activeSidePanel,
    downloadFile: state.downloadFile,
    pricingMatches: state.pricingMatches,
    pricingIssues: state.pricingIssues,
    activeJob: state.activeJob,
  };
}

function clearSessionState() {
  try {
    window.localStorage.removeItem(QUOTE_SESSION_STORAGE_KEY);
  } catch {
    // Ignore storage failures; the local workflow can continue without refresh recovery.
  }
  clearSessionFiles().catch(() => {});
}

function saveSessionState() {
  if (state.isBooting) return;
  try {
    const fileRecords = sessionFileRecordsFromImages();
    const snapshot = buildSessionSnapshot();
    window.localStorage.setItem(QUOTE_SESSION_STORAGE_KEY, JSON.stringify(snapshot));
    persistSessionFiles(fileRecords).catch(() => {});
  } catch {
    // Refresh recovery is best-effort; the local workflow can continue without persisted state.
  }
}

async function restoreSessionImages(savedImages = []) {
  const images = Array.isArray(savedImages) ? savedImages.slice(0, MAX_REFERENCE_IMAGES) : [];
  const keys = images
    .filter((image) => !String(image?.data_url || "").trim())
    .map((image) => image?.session_file_key);
  const fileMap = await loadSessionFileMap(keys).catch(() => new Map());
  return images
    .map((image) => {
      const key = String(image?.session_file_key || "").trim();
      const stored = key ? fileMap.get(key) : null;
      return {
        ...image,
        ...(stored || {}),
        type: stored?.type || image?.type || referenceFileType(stored || image),
        size: Number.isFinite(Number(stored?.size ?? image?.size)) ? Number(stored?.size ?? image?.size) : 0,
      };
    })
    .filter((image) => String(image.data_url || "").trim());
}

async function restoreSessionState() {
  const saved = safeSessionJson();
  if (!saved || saved.version !== QUOTE_SESSION_STATE_VERSION) {
    clearSessionState();
    return false;
  }
  state.profileId = saved.profileId || "";
  state.pricingReferenceId = saved.pricingReferenceId || saved.profileId || "";
  state.pricingReferenceSource = saved.pricingReferenceSource || "";
  state.selectedPresetValue = saved.selectedPresetValue || presetValueFromQuoteDetails(saved.quoteDetails || {});
  syncSelectedPricingReference();
  renderProfileOptions();
  renderPresetOptions();
  applyQuoteDetails(saved.quoteDetails || {}, { includeLogo: true, clearLogo: true });
  state.images = await restoreSessionImages(saved.images);
  state.quoteBasis = cloneQuoteBasis(saved.quoteBasis || {});
  state.quoteBasisSections = normalizeQuoteBasisSections(saved.quoteBasisSections || saved.quoteBasis || {});
  state.lineItems = Array.isArray(saved.lineItems) ? saved.lineItems.map(normalizeLineItem) : [];
  state.outputRows = Array.isArray(saved.outputRows) ? saved.outputRows.map(normalizeOutputRow) : [];
  state.originalOutputRows = Array.isArray(saved.originalOutputRows) ? saved.originalOutputRows.map(normalizeOutputRow) : [];
  state.outputErrors = Array.isArray(saved.outputErrors) ? saved.outputErrors : [];
  state.outputSortMode = OUTPUT_SORT_MODES.includes(saved.outputSortMode) ? saved.outputSortMode : "pricing_reference";
  state.analysisFindings = Array.isArray(saved.analysisFindings) ? saved.analysisFindings : [];
  state.blockingClarificationQuestions = Array.isArray(saved.blockingClarificationQuestions) ? saved.blockingClarificationQuestions : [];
  state.boothDimensions = normalizeBoothDimensions(saved.boothDimensions || saved.quoteDetails?.project || {});
  state.originalAnalysisSnapshot = saved.originalAnalysisSnapshot || null;
  state.basisConfirmed = Boolean(saved.basisConfirmed);
  state.draftSource = saved.draftSource || "";
  state.lastAnalysisMode = normalizeAnalysisMode(saved.lastAnalysisMode || saved.originalAnalysisSnapshot?.analysis_mode);
  state.aiFailed = Boolean(saved.aiFailed || state.draftSource === "local");
  state.downloadFile = saved.downloadFile || null;
  state.pricingMatches = Array.isArray(saved.pricingMatches) ? saved.pricingMatches : [];
  state.pricingIssues = [];
  state.activeJob = saved.activeJob && saved.activeJob.id ? saved.activeJob : null;
  renderFiles();
  renderPricingMatches(state.outputRows.length ? state.outputRows : state.pricingMatches, { fromPricingMatches: !state.outputRows.length && state.pricingMatches.length });
  renderMatchSummary({ pricing_matches: state.pricingMatches });
  clearPricingReviewMessages();
  if (state.aiFailed) {
    clearAiFailedDraftState();
    renderBasisFailureState();
  } else if (state.lineItems.length || state.quoteBasisSections.length || Object.values(state.quoteBasis).some((value) => splitLines(value).length > 0)) {
    updateQuoteBasisCard(saved.draftSource || "restored");
  } else {
    renderBasisEmptyState();
  }
  updateDownloadButton();
  setResultStatus(state.aiFailed ? "No usable AI draft" : state.downloadFile ? "Completed" : "No job yet", state.aiFailed ? "is-bad" : state.downloadFile ? "is-ok" : "");
  setWorkflowStage(saved.workflowStage || (state.images.length ? "ready_to_analyze" : "needs_images"));
  if (state.aiFailed) {
    showAiFailureBanner();
  } else {
    clearAiFailureBanner();
  }
  const restoredPanel = SIDE_PANEL_SEQUENCE.includes(saved.activeSidePanel) ? saved.activeSidePanel : "images";
  setSidePanel(restoredPanel, { force: true });
  return true;
}

function clearLegacyLocalCompanyPresets() {
  try {
    window.localStorage.removeItem(LEGACY_QUOTE_PRESETS_STORAGE_KEY);
  } catch {
    // Ignore storage failures; repo-backed presets still work.
  }
}

function selectedPresetId() {
  return state.selectedPresetValue || elements.presetSelect.value || "";
}

function profilePresets() {
  return state.profiles.flatMap((profile) => {
    const presets = Array.isArray(profile.quote_detail_presets) ? profile.quote_detail_presets : [];
    return presets.map((preset) => ({
      ...preset,
      profile_id: preset.profile_id || profile.id,
      profile_label: profile.label || profile.id,
    }));
  });
}

function defaultProfilePresetId() {
  const profile = currentProfile();
  const configured = profile.default_quote_detail_preset || "default";
  const presets = profilePresets();
  if (presets.some((preset) => preset.id === "default")) return "default";
  if (configured && presets.some((preset) => preset.id === configured)) return configured;
  return presets[0]?.id || "";
}

function profilePresetOptionValue(presetId) {
  return `${PROFILE_PRESET_PREFIX}${presetId}`;
}

function presetOptionValue(preset = {}) {
  const presetId = String(preset.id || "").trim();
  if (!presetId) return "";
  return profilePresetOptionValue(presetId);
}

function defaultPresetOptionValue() {
  const profileDefault = defaultProfilePresetId();
  if (profileDefault) return profilePresetOptionValue(profileDefault);
  return "";
}

function configuredProfilePresetId() {
  const profile = currentProfile();
  const configured = profile.default_quote_detail_preset || "";
  const presets = profilePresets();
  return configured && presets.some((preset) => preset.id === configured) ? configured : "";
}

function selectedPreset() {
  const value = selectedPresetId();
  if (value.startsWith(PROFILE_PRESET_PREFIX)) {
    const presetId = value.slice(PROFILE_PRESET_PREFIX.length);
    const preset = profilePresets().find((item) => item.id === presetId);
    return preset ? { ...preset, source: "profile" } : null;
  }
  return null;
}

function normalizePresetComparisonValue(value) {
  if (Array.isArray(value)) return value.map(normalizePresetComparisonValue).join("\n");
  if (value && typeof value === "object") return JSON.stringify(value);
  return String(value || "").replace(/\s+/g, " ").trim().toLowerCase();
}

function quoteDetailsMatchPreset(savedDetails = {}, presetDetails = {}) {
  const savedCompany = savedDetails.company || {};
  const presetCompany = presetDetails.company || {};
  const savedRichText = savedDetails.rich_text || {};
  const presetRichText = presetDetails.rich_text || {};
  const companyNameMatches = presetCompany.name
    && normalizePresetComparisonValue(savedCompany.name) === normalizePresetComparisonValue(presetCompany.name);
  const headerDetailsMatch = presetCompany.header_details
    && normalizePresetComparisonValue(savedCompany.header_details) === normalizePresetComparisonValue(presetCompany.header_details);
  const richHeaderMatches = presetRichText.headerDetails
    && normalizePresetComparisonValue(savedRichText.headerDetails) === normalizePresetComparisonValue(presetRichText.headerDetails);
  return Boolean(companyNameMatches && (headerDetailsMatch || richHeaderMatches));
}

function presetValueFromQuoteDetails(savedDetails = {}) {
  const matchesDetails = (preset) => quoteDetailsMatchPreset(savedDetails, preset.details || {});
  const profilePreset = profilePresets().find(matchesDetails);
  if (profilePreset) return profilePresetOptionValue(profilePreset.id);
  return "";
}

function renderPresetStatus(message = "") {
  if (!elements.presetStatus) return;
  elements.presetStatus.textContent = message || "Database save pending.";
}

function renderHeaderLogoPreview() {
  if (!elements.headerLogoPreview) return;
  if (state.headerLogo) {
    elements.headerLogoPreview.classList.add("has-logo");
    elements.headerLogoPreview.innerHTML = `<img src="${escapeHtml(state.headerLogo.data_url)}" alt="Selected header logo preview">`;
    return;
  }
  elements.headerLogoPreview.classList.remove("has-logo");
  elements.headerLogoPreview.innerHTML = "<span>No logo loaded</span>";
}

function renderPresetOptions() {
  const builtInPresets = profilePresets();
  const selectedValue = state.selectedPresetValue || elements.presetSelect.value || "";
  const defaultPreset = builtInPresets.find((preset) => preset.id === "default");
  const defaultOption = defaultPreset
    ? `<option value="${escapeHtml(profilePresetOptionValue(defaultPreset.id))}">${escapeHtml(defaultPreset.name)}</option>`
    : "";
  const builtInOptions = builtInPresets
    .filter((preset) => preset.id !== "default")
    .map((preset) => `<option value="${escapeHtml(profilePresetOptionValue(preset.id))}">${escapeHtml(preset.name)}</option>`)
    .join("");
  elements.presetSelect.innerHTML = [
    defaultOption,
    builtInOptions ? `<optgroup label="Profile Presets">${builtInOptions}</optgroup>` : "",
  ].join("");
  const availableValues = new Set([
    ...builtInPresets.map((preset) => profilePresetOptionValue(preset.id)),
  ]);
  state.selectedPresetValue = availableValues.has(selectedValue) ? selectedValue : defaultPresetOptionValue();
  elements.presetSelect.value = state.selectedPresetValue;
  updatePresetButtons();
}

function updatePresetButtons() {
  const preset = selectedPreset();
  elements.loadPresetButton.disabled = !preset;
  elements.deletePresetButton.disabled = true;
}

function saveCurrentPreset() {
  renderPresetStatus("Database save pending.");
}

function loadSelectedPreset(options = {}) {
  const preset = selectedPreset();
  if (!preset) {
    renderPresetStatus("Choose a preset to load.");
    return;
  }
  state.selectedPresetValue = elements.presetSelect.value || presetOptionValue(preset);
  const details = preset.details || {};
  const clearsLogo = Boolean(details.company && typeof details.company === "object");
  applyQuoteDetails(details, { includeLogo: true, clearLogo: clearsLogo, partial: true });
  clearGeneratedQuoteState();
  setWorkflowStage(state.images.length ? "ready_to_analyze" : "needs_images");
  syncControlStates();
  renderPresetStatus(`Loaded "${preset.name}".`);
}

function loadDefaultProfilePreset(options = {}) {
  const defaultPreset = defaultPresetOptionValue();
  if (!defaultPreset) return;
  state.selectedPresetValue = defaultPreset;
  elements.presetSelect.value = state.selectedPresetValue;
  loadSelectedPreset(options);
}

function loadConfiguredProfilePreset(options = {}) {
  const configuredPreset = configuredProfilePresetId();
  if (!configuredPreset) {
    loadDefaultProfilePreset(options);
    return;
  }
  state.selectedPresetValue = profilePresetOptionValue(configuredPreset);
  elements.presetSelect.value = state.selectedPresetValue;
  loadSelectedPreset(options);
}

function clearCustomerDetails() {
  setInputValue(elements.clientName, "");
  setInputValue(elements.clientAttention, "");
  setInputValue(elements.clientTitle, "");
  setInputValue(elements.clientAddress, "");
  setInputValue(elements.projectTitle, "");
  state.boothDimensions = { ...DEFAULT_BOOTH_DIMENSIONS };
  setInputValue(elements.quoteDate, todayDateInputValue());
  applyQuoteDateFormatFromHtml("");
  setInputValue(elements.projectNumber, "");
  clearGeneratedQuoteState();
  renderProfileOptions();
  setWorkflowStage(state.images.length ? "ready_to_analyze" : "needs_images");
  syncControlStates();
}

function clearQuoteCompanyDetails() {
  elements.presetNameInput.value = "";
  setInputValue(elements.headerDetails, "");
  setInputValue(elements.termsHeading, "");
  setInputValue(elements.paymentTerms, "");
  setInputValue(elements.notesHeading, "");
  setInputValue(elements.standardNotes, "");
  setInputValue(elements.quoteCompanyName, "");
  setInputValue(elements.acceptanceText, "");
  setInputValue(elements.konceptSignatory, "");
  setInputValue(elements.konceptTitle, "");
  setInputValue(elements.konceptDateLabel, "");
  setInputValue(elements.personLabel, "");
  setInputValue(elements.stampLabel, "");
  setInputValue(elements.dateLabel, "");
  applyDefaultQuoteCompanyFields();
  state.headerLogo = null;
  elements.headerLogoInput.value = "";
  state.selectedPresetValue = "";
  clearGeneratedQuoteState();
  setWorkflowStage(state.images.length ? "ready_to_analyze" : "needs_images");
  updatePresetButtons();
  renderHeaderLogoPreview();
  loadDefaultProfilePreset({ silent: true });
  syncControlStates();
  renderPresetStatus("Quote-company defaults reset to the selected company preset.");
}

function resetImagesDraft() {
  if (state.isAnalysisRunning || state.isGenerating) return;
  state.images = [];
  if (elements.imageInput) elements.imageInput.value = "";
  setImageUploadStatus("");
  clearGeneratedQuoteState();
  renderFiles();
  setWorkflowStage("needs_images");
  syncControlStates();
}

function startNewQuote() {
  if (state.isAnalysisRunning || state.isGenerating) return;
  clearSessionState();
  state.profileId = "";
  state.pricingReferenceId = "";
  state.pricingReferenceSource = "";
  state.images = [];
  setImageUploadStatus("");
  state.headerLogo = null;
  state.boothDimensions = { ...DEFAULT_BOOTH_DIMENSIONS };
  state.pendingFeedback = "";
  state.downloadFile = null;
  elements.imageInput.value = "";
  elements.headerLogoInput.value = "";
  state.selectedPresetValue = "";
  elements.presetSelect.value = "";
  elements.presetNameInput.value = "";
  applyQuoteDetails({}, { clearLogo: true });
  applyDefaultQuoteCompanyFields();
  clearGeneratedQuoteState();
  renderFiles();
  renderProfileOptions();
  renderPresetOptions();
  loadDefaultProfilePreset({ silent: true });
  renderHeaderLogoPreview();
  renderPresetStatus("Started a new quote.");
  setWorkflowStage("needs_images");
  setSidePanel("images", { force: true });
  syncControlStates();
}

function deleteSelectedPreset() {
  renderPresetStatus("Company preset deletion will use the database later. For now, presets come from repo profile templates.");
}

function renderProfileOptions() {
  if (!elements.profileSelect) return;
  const references = sortedPricingReferencesForDisplay(
    state.pricingReferences.filter((reference) => String(reference?.source || "bundled") === "bundled")
  );
  const selectedValue = currentPricingReference() ? pricingReferenceSelectValue(currentPricingReference()) : "";
  const referenceOption = (reference) => {
    const referenceId = String(reference.id || "").trim();
    return `<option value="${escapeHtml(pricingReferenceSelectValue(reference))}">${escapeHtml(reference.label || referenceId)}</option>`;
  };
  elements.profileSelect.innerHTML = references.map(referenceOption).join("");
  const fallbackReference = currentPricingReference() || defaultPricingReference() || references[0] || null;
  if (fallbackReference) {
    state.pricingReferenceId = fallbackReference.id || "";
    state.pricingReferenceSource = pricingReferenceSelectionFromValue(pricingReferenceSelectValue(fallbackReference)).source;
  }
  const selectedReference = currentPricingReference();
  elements.profileSelect.value = selectedReference ? pricingReferenceSelectValue(selectedReference) : selectedValue;
  renderSelectedPricingReferenceSummary();
  renderPricingReferenceDeleteOptions();
}

function canManagePricingReferences() {
  return Boolean(state.permissions?.canManagePricingReferences);
}

function pricingReferenceNoAccessReason() {
  return "You do not have access to manage pricing references.";
}

function protectedPricingReferenceReason(reference = currentPricingReference()) {
  if (!reference) return "Select a pricing reference first.";
  if (!canManagePricingReferences()) return pricingReferenceNoAccessReason();
  if (String(reference.source || "bundled") !== "bundled") return "Only repo pricing reference packs can be deleted here.";
  const referenceId = String(reference.id || "").trim();
  if (!referenceId) return "Select a pricing reference first.";
  const profileDefault = String(currentProfile()?.default_pricing_reference || state.defaultPricingReferenceId || DEFAULT_PRICING_REFERENCE_ID).trim();
  if (referenceId === DEFAULT_PRICING_REFERENCE_ID || referenceId === profileDefault) return "Default pricing references cannot be deleted.";
  return "";
}

function deletionPricingReference() {
  const select = elements.deletePricingReferenceSelect;
  if (!select) return currentPricingReference();
  const selection = pricingReferenceSelectionFromValue(select.value || "");
  return state.pricingReferences.find((reference) => (
    reference.id === selection.pricingReferenceId
    && pricingReferenceSelectValue(reference).startsWith(`${selection.source || "bundled"}::`)
  )) || null;
}

function canDeleteSelectedPricingReference() {
  return !protectedPricingReferenceReason(deletionPricingReference());
}

function renderPricingReferenceDeleteOptions() {
  const select = elements.deletePricingReferenceSelect;
  if (!select) return;
  const previousValue = select.value;
  const references = sortedPricingReferencesForDisplay(
    state.pricingReferences.filter((reference) => String(reference?.source || "bundled") === "bundled")
  );
  select.innerHTML = references.map((reference) => `
    <option value="${escapeHtml(pricingReferenceSelectValue(reference))}">${escapeHtml(reference.label || reference.id || "Pricing reference")}</option>
  `).join("");
  const preferred = references.find((reference) => pricingReferenceSelectValue(reference) === previousValue)
    || references.find((reference) => !protectedPricingReferenceReason(reference))
    || references[0]
    || null;
  select.value = preferred ? pricingReferenceSelectValue(preferred) : "";
  updatePricingReferenceDeleteButton();
}

function updatePricingReferenceDeleteButton() {
  const button = elements.deletePricingReferenceButton;
  if (!button) return;
  if (elements.pricingReferenceDeleteSection) {
    elements.pricingReferenceDeleteSection.hidden = !canManagePricingReferences();
  }
  const reference = deletionPricingReference();
  const reason = protectedPricingReferenceReason(reference);
  button.disabled = Boolean(reason);
  button.title = reason || "Delete this repo pricing reference.";
  button.setAttribute("aria-disabled", String(button.disabled));
}

function openSettingsModal() {
  if (!canManagePricingReferences()) return;
  openPricingReferenceModal();
}

function pricingReferenceRequiredColumns() {
  return ["section", "description", "unit_hint", "internal_cost", "markup_multiplier"];
}

const PRICING_REFERENCE_PREVIEW_FIELDS = ["section", "description", "unit_hint", "internal_cost", "markup_multiplier", "remarks"];

function normalizeBasisTag(tag = "") {
  const normalized = String(tag || "").trim().toLowerCase();
  if (normalized === "include" || normalized === "matched") return "Include";
  if (["custom", "manual", "extra", "non-catalog", "non catalog", "needs-pricing", "needs pricing"].includes(normalized)) return "Custom";
  if (normalized === "exclude") return "Exclude";
  return "Confirm";
}

function isCustomPricingBasisLine(line = {}) {
  return normalizeBasisTag(line.tag) === "Custom"
    || Boolean(line.custom_pricing || line.custom || line.manual_pricing)
    || normalizeBasisTag(line.pricing_tag || line.pricing_status) === "Custom";
}

function isPendingAiProposalLine(line = {}) {
  const tag = normalizeBasisTag(line.tag);
  return isCustomPricingBasisLine(line) && (tag === "Confirm" || (tag === "Custom" && !line.custom_confirmed));
}

function normalizeConfidence(value) {
  if (value === null || value === undefined || value === "") return null;
  const number = Number(String(value).replace("%", "").trim());
  if (!Number.isFinite(number)) return null;
  return Math.min(100, Math.max(0, Math.round(number)));
}

function splitBasisDecisionText(text = "", defaultTag = "Confirm") {
  const legacyConfirmTag = "assump" + "tion";
  const raw = String(text || "").trim();
  if (!raw) return [];
  const pattern = `Include|Confirm|Custom|Manual|Extra|Needs Pricing|Exclude|matched|${legacyConfirmTag}|Note`;
  const expression = new RegExp(`(?:^|[;\\n]\\s*)(${pattern}):\\s*`, "gi");
  const matches = Array.from(raw.matchAll(expression));
  if (!matches.length) return [{ tag: normalizeBasisTag(defaultTag), text: raw }];

  const lines = [];
  const firstStart = matches[0].index || 0;
  const leading = raw.slice(0, firstStart).replace(/[;\s]+$/g, "").trim();
  if (leading) lines.push({ tag: normalizeBasisTag(defaultTag), text: leading });
  matches.forEach((match, index) => {
    const nextStart = index + 1 < matches.length ? matches[index + 1].index : raw.length;
    const segment = raw.slice(match.index + match[0].length, nextStart).replace(/^[;\s]+|[;\s]+$/g, "").trim();
    if (segment) lines.push({ tag: normalizeBasisTag(match[1]), text: segment });
  });
  return lines;
}

function normalizeBasisLines(line = "") {
  if (line && typeof line === "object") {
    const quantityParts = normalizedLineTextQuantityParts(
      line.text || line.line || line.description || "",
      line.quantity,
      line.unit
    );
    const text = quantityParts.text;
    if (!text) return [];
    const confidence = normalizeConfidence(line.confidence ?? line.confidence_pct);
    const hasCustomPricing = isCustomPricingBasisLine(line);
    return splitBasisDecisionText(text, line.tag).map((parsed) => {
      const next = { tag: normalizeBasisTag(parsed.tag), text: parsed.text };
      if (line.id) next.id = String(line.id);
      if (line.source_line_item_id) next.source_line_item_id = String(line.source_line_item_id);
      if (line.pricing_keyword) next.pricing_keyword = String(line.pricing_keyword);
      if (line.catalog_unit_price !== undefined && line.catalog_unit_price !== null && String(line.catalog_unit_price).trim() !== "") next.catalog_unit_price = line.catalog_unit_price;
      if (line.catalog_description) next.catalog_description = cleanCustomerQuoteLineText(line.catalog_description);
      if (line.pricing_reference_description) next.pricing_reference_description = pricingReferenceLineText(line.pricing_reference_description);
      if (orderNumber(line.category_order) !== null) next.category_order = orderNumber(line.category_order);
      if (orderNumber(line.item_order) !== null) next.item_order = orderNumber(line.item_order);
      if (quantityParts.quantity !== undefined && quantityParts.quantity !== null && String(quantityParts.quantity).trim() !== "") next.quantity = quantityParts.quantity;
      if (quantityParts.unit) next.unit = quantityParts.unit;
      if (confidence !== null) next.confidence = confidence;
      if (hasCustomPricing || normalizeBasisTag(parsed.tag) === "Custom") next.custom_pricing = true;
      if (line.custom_confirmed) next.custom_confirmed = true;
      return next;
    });
  }
  return splitBasisDecisionText(cleanCustomerQuoteLineText(line));
}

function parseBasisLine(line = "") {
  return normalizeBasisLines(line)[0] || { tag: "Confirm", text: "" };
}

function normalizeQuoteBasisSections(value = {}) {
  const rawSections = Array.isArray(value)
    ? value
    : Array.isArray(value.quote_basis_sections)
      ? value.quote_basis_sections
      : null;
  if (rawSections) {
    return rawSections
      .map((section, index) => {
        const title = normalizeQuoteBasisTitle(section?.title || "Section") || "Section";
        const id = safeId(section?.id && /^[a-z0-9]+(?:-[a-z0-9]+)*$/.test(String(section.id)) ? section.id : title, `section-${index + 1}`);
        const rawLines = Array.isArray(section?.lines) ? section.lines : splitLines(section?.text || "");
        const lines = rawLines.flatMap(normalizeBasisLines).filter((line) => line.text);
        return lines.length ? { id, title, lines } : null;
      })
      .filter(Boolean);
  }
  const basis = value.quote_basis && typeof value.quote_basis === "object" ? value.quote_basis : value;
  return BASIS_FIELDS
    .map(([key, title]) => {
      const lines = splitLines(basis[key])
        .flatMap(normalizeBasisLines)
        .filter((line) => line.text);
      return lines.length ? { id: key, title, lines } : null;
    })
    .filter(Boolean);
}

function confirmOnlyQuoteBasisSections(sections = []) {
  return normalizeQuoteBasisSections(sections).map((section) => ({
    ...section,
    lines: (section.lines || []).map((line) => {
      const tag = normalizeBasisTag(line.tag);
      return { ...line, tag: tag === "Custom" ? "Custom" : "Confirm" };
    }),
  }));
}

function quoteBasisFromSections(sections = []) {
  return (Array.isArray(sections) ? sections : []).reduce((basis, section) => {
    const id = safeId(section.id || section.title, "section");
    basis[id] = (section.lines || [])
      .map((line) => `${normalizeBasisTag(line.tag)}: ${line.text || ""}`.trim())
      .filter((line) => !/:\s*$/.test(line))
      .join("\n");
    return basis;
  }, {});
}

function cloneQuoteBasisSections(sections = []) {
  return normalizeQuoteBasisSections(JSON.parse(JSON.stringify(Array.isArray(sections) ? sections : [])));
}

function normalizeOutputRow(row = {}) {
  const priceMode = row.price_mode === "Included" || String(row.display_price || "").toLowerCase() === "included"
    ? "Included"
    : "Priced";
  return recalculateOutputRow({
    section: normalizeCategoryTitle(row.section || ""),
    description: cleanCustomerQuoteLineText(row.description || ""),
    quantity: row.quantity ?? "",
    unit: normalizeUnit(row.unit || ""),
    price_mode: priceMode,
    unit_price_override: row.unit_price_override ?? "",
    catalog_unit_price: row.catalog_unit_price ?? row.unit_price ?? row.sale_unit_price ?? "",
    catalog_description: cleanCustomerQuoteLineText(row.catalog_description || ""),
    pricing_reference_description: pricingReferenceLineText(row.pricing_reference_description || row.reference_description || ""),
    pricing_keyword: row.pricing_keyword || row.keyword || "",
    source_basis_line_id: row.source_basis_line_id || "",
    category_order: orderNumber(row.category_order) ?? "",
    item_order: orderNumber(row.item_order) ?? "",
    status: row.status || "",
  });
}

function pricingReferenceValidationResult(items, headers, skipped, sourceName = "") {
  const required = pricingReferenceRequiredColumns();
  const headerSet = new Set(headers.map((header) => String(header || "").trim()));
  const missing = required.filter((column) => !headerSet.has(column));
  const errors = [];
  const warnings = [];
  if (!items.length) errors.push("No valid pricing rows were found.");
  if (missing.length) errors.push(`Missing required columns: ${missing.join(", ")}.`);
  if (skipped) warnings.push(`${skipped} row${skipped === 1 ? "" : "s"} skipped during sanitizing.`);
  return {
    sourceName,
    items: sortPricingReferencePreviewItems(items),
    rowCount: items.length,
    headers,
    missing,
    skipped,
    layout: "normalized-pricing-reference",
    errors,
    warnings,
    canSave: !errors.length && items.length > 0,
  };
}

function sortPricingReferencePreviewItems(items = []) {
  const sectionOrders = new Map();
  let nextCategoryOrder = 1;
  const ordered = [...items].map((item, index) => {
    const sectionKey = safeId(normalizeCategoryTitle(item?.section || ""), `section-${index + 1}`);
    const explicitCategoryOrder = orderNumber(item?.category_order);
    if (!sectionOrders.has(sectionKey)) {
      const categoryOrder = explicitCategoryOrder || nextCategoryOrder;
      sectionOrders.set(sectionKey, categoryOrder);
      nextCategoryOrder = Math.max(nextCategoryOrder + 1, categoryOrder + 1);
    }
    return {
      item: {
        ...item,
        category_order: sectionOrders.get(sectionKey),
        item_order: orderNumber(item?.item_order) || index + 1,
      },
      index,
    };
  });
  ordered.sort((a, b) => {
    const categoryCompare = (orderNumber(a.item.category_order) || 999999) - (orderNumber(b.item.category_order) || 999999);
    if (categoryCompare) return categoryCompare;
    const itemOrderCompare = (orderNumber(a.item.item_order) || 999999) - (orderNumber(b.item.item_order) || 999999);
    if (itemOrderCompare) return itemOrderCompare;
    const sectionCompare = String(a.item?.section || "").toLowerCase().localeCompare(String(b.item?.section || "").toLowerCase());
    if (sectionCompare) return sectionCompare;
    const descriptionCompare = String(a.item?.description || "").toLowerCase().localeCompare(String(b.item?.description || "").toLowerCase());
    if (descriptionCompare) return descriptionCompare;
    return String(a.item?.id || "").toLowerCase().localeCompare(String(b.item?.id || "").toLowerCase()) || a.index - b.index;
  });
  return ordered.map(({ item }) => item);
}

function pricingReferenceRowStatus(row = {}) {
  const warnings = [];
  if (!String(row.section || "").trim()) warnings.push("section required");
  if (!String(row.description || "").trim()) warnings.push("description required");
  if (!String(row.unit_hint || "").trim()) warnings.push("unit_hint required");
  const cost = numberOrNull(row.internal_cost);
  const markup = numberOrNull(row.markup_multiplier);
  if (cost === null || cost <= 0) warnings.push("internal_cost must be positive");
  if (markup === null || markup <= 0) warnings.push("markup_multiplier must be positive");
  return warnings;
}

function pricingReferenceSaveBlockReason(result = state.pendingPricingReference) {
  if (state.pricingReferenceImportBusy) return "Import preview is still being prepared.";
  if (!result) return "Upload a pricing catalog file before saving.";
  if (elements.pricingReferenceCurrency?.value === CUSTOM_CURRENCY_VALUE && !customCurrencyInputIsValid()) {
    return "Enter a 3-letter currency code before saving.";
  }
  const errors = Array.isArray(result.errors) ? result.errors.filter(Boolean) : [];
  if (errors.length) return errors.join(" ");
  if (!(result.items || []).length) return "Add at least one valid pricing row before saving.";
  if (!result.canSave) return "Fix the highlighted pricing-reference rows before saving.";
  return "";
}

function setPricingReferenceModalBusyState(busy = false, reason = "") {
  const canManage = canManagePricingReferences();
  const disabled = busy || !canManage;
  const disabledTitle = !canManage ? pricingReferenceNoAccessReason() : busy ? reason || "Import preview is still being prepared." : "";
  const closeTitle = busy ? reason || "Import preview is still being prepared." : "";
  if (elements.pricingReferenceModal) {
    elements.pricingReferenceModal.classList.toggle("is-busy", busy);
    elements.pricingReferenceModal.classList.toggle("is-denied", !canManage);
  }
  [elements.pricingReferenceCloseButton, elements.pricingReferenceCancelButton].forEach((button) => {
    if (!button) return;
    button.disabled = busy;
    button.title = closeTitle;
    button.setAttribute("aria-disabled", String(busy));
  });
  if (elements.pricingReferenceFile) {
    elements.pricingReferenceFile.disabled = disabled;
  }
  if (elements.pricingReferenceTemplateButton) {
    elements.pricingReferenceTemplateButton.setAttribute("aria-disabled", String(disabled));
    elements.pricingReferenceTemplateButton.setAttribute("tabindex", disabled ? "-1" : "0");
    elements.pricingReferenceTemplateButton.title = disabledTitle;
  }
}

function setPricingReferenceSaveButtonState(options = {}) {
  const button = elements.pricingReferenceSaveButton;
  if (!button) return;
  const busy = Boolean(options.busy);
  const canManage = canManagePricingReferences();
  const canSave = canManage && Boolean(options.canSave) && !busy;
  const disabled = busy || !canManage || !canSave;
  const reason = canManage ? String(options.reason || "").trim() : pricingReferenceNoAccessReason();
  button.disabled = disabled;
  button.textContent = busy ? "Importing..." : "Save Reference";
  button.title = disabled ? reason : "";
  button.setAttribute("aria-disabled", String(disabled));
  setPricingReferenceModalBusyState(busy, reason);
}

function setPricingReferenceModalAccessState() {
  const canManage = canManagePricingReferences();
  if (elements.pricingReferenceNoAccess) elements.pricingReferenceNoAccess.hidden = canManage;
  if (elements.pricingReferenceEditorBody) elements.pricingReferenceEditorBody.hidden = !canManage;
  if (elements.pricingReferenceCancelButton) elements.pricingReferenceCancelButton.textContent = canManage ? "Cancel" : "Back";
  if (!canManage) {
    state.pendingPricingReference = null;
    state.pricingReferenceImportBusy = false;
    state.pricingReferenceImportToken = "";
  }
  setPricingReferenceSaveButtonState({
    canSave: false,
    reason: canManage ? pricingReferenceSaveBlockReason(state.pendingPricingReference) : pricingReferenceNoAccessReason(),
  });
  return canManage;
}

function blockPricingReferenceBusyInteraction(event) {
  if (!state.pricingReferenceImportBusy) return;
  const blockedControl = event.target.closest("button, a, label, input, select, textarea, [role='button'], [data-pricing-reference-close]");
  if (!blockedControl) return;
  event.preventDefault();
  event.stopPropagation();
}

function normalizePricingReferencePreviewItem(item = {}, index = 0) {
  return {
    id: safeId(item.id || `${item.section || "item"}-${item.description || index + 1}`, `item-${index + 1}`),
    section: normalizeCategoryTitle(neutralizeFormulaText(item.section || "")),
    description: cleanCustomerQuoteLineText(neutralizeFormulaText(item.description || "")),
    unit_hint: normalizeUnit(neutralizeFormulaText(item.unit_hint || item.unit || "")),
    internal_cost: item.internal_cost ?? "",
    markup_multiplier: item.markup_multiplier ?? "",
    category_order: orderNumber(item.category_order) ?? "",
    item_order: orderNumber(item.item_order) ?? index + 1,
    remarks: Array.isArray(item.remarks) ? item.remarks.map(neutralizeFormulaText).join("; ") : neutralizeFormulaText(item.remarks || ""),
    aliases: Array.isArray(item.aliases) ? item.aliases.map(neutralizeFormulaText).join(" | ") : neutralizeFormulaText(item.aliases || ""),
    visual_references: Array.isArray(item.visual_references) ? item.visual_references : [],
    warning: item.warning || item.status || "",
  };
}

function refreshPricingReferencePreviewValidity(result = state.pendingPricingReference) {
  if (!result) return false;
  const ids = new Set();
  const duplicateIds = new Set();
  result.items = sortPricingReferencePreviewItems((result.items || []).map(normalizePricingReferencePreviewItem));
  result.items.forEach((item) => {
    if (ids.has(item.id)) duplicateIds.add(item.id);
    ids.add(item.id);
  });
  const invalid = result.items.flatMap((item) => {
    const warnings = pricingReferenceRowStatus(item);
    if (duplicateIds.has(item.id)) warnings.push("duplicate id");
    item.warning = warnings.join("; ") || "OK";
    return warnings;
  });
  result.canSave = !invalid.length && result.items.length > 0 && !(result.errors || []).length;
  setPricingReferenceSaveButtonState({
    canSave: result.canSave,
    busy: state.pricingReferenceImportBusy,
    reason: pricingReferenceSaveBlockReason(result),
  });
  return result.canSave;
}

function pricingReferencePreviewTableRows(result = state.pendingPricingReference) {
  const items = Array.isArray(result?.items) ? result.items : [];
  return items.map((item, index) => `
    <tr data-preview-row="${index}">
      ${PRICING_REFERENCE_PREVIEW_FIELDS.map((field) => `
        <td><input class="pricing-preview-input" data-preview-field="${field}" data-preview-row="${index}" value="${escapeHtml(item[field] ?? "")}"></td>
      `).join("")}
      <td class="pricing-preview-status ${pricingReferenceStatusClass(item.warning || "OK")}">${escapeHtml(item.warning || "OK")}</td>
    </tr>
  `).join("");
}

function pricingReferenceStatusClass(status = "") {
  const text = String(status || "").trim().toLowerCase();
  if (!text || text === "ok") return "is-ok";
  if (/(required|must|duplicate|missing|invalid|error|failed|empty|positive)/.test(text)) return "is-error";
  return "is-warn";
}

function pricingReferenceSaveGuidance(result = state.pendingPricingReference) {
  if (!result) return "";
  const reason = pricingReferenceSaveBlockReason(result);
  return result.canSave
    ? "Rows are ready to save."
    : `Fix all flagged problems before saving this reference.${reason ? ` ${reason}` : ""}`;
}

function pricingReferenceTableSummaryText(result = state.pendingPricingReference) {
  const items = Array.isArray(result?.items) ? result.items : [];
  return items.length
    ? `${items.length} editable pricing row${items.length === 1 ? "" : "s"}. ${pricingReferenceSaveGuidance(result)}`
    : "Upload a pricing catalog to review editable rows.";
}

function updatePricingReferenceGuidanceDisplays(result = state.pendingPricingReference) {
  if (elements.pricingReferenceTableSummary) {
    elements.pricingReferenceTableSummary.textContent = pricingReferenceTableSummaryText(result);
  }
  const guidance = elements.pricingReferencePreview?.querySelector(".pricing-reference-save-guidance");
  if (guidance) {
    guidance.textContent = pricingReferenceSaveGuidance(result);
    guidance.classList.toggle("is-ok", Boolean(result?.canSave));
    guidance.classList.toggle("is-blocked", !result?.canSave);
  }
}

function renderPricingReferenceTableOverlay(result = state.pendingPricingReference) {
  if (!elements.pricingReferenceTableBody) return;
  const items = Array.isArray(result?.items) ? result.items : [];
  updatePricingReferenceGuidanceDisplays(result);
  elements.pricingReferenceTableBody.innerHTML = items.length
    ? pricingReferencePreviewTableRows(result)
    : `<tr><td colspan="7" class="pricing-reference-table-empty">No imported rows to review.</td></tr>`;
}

function openPricingReferenceTableOverlay() {
  if (!elements.pricingReferenceTableOverlay || state.pricingReferenceImportBusy) return;
  renderPricingReferenceTableOverlay(state.pendingPricingReference);
  elements.pricingReferenceTableOverlay.hidden = false;
  elements.pricingReferenceTableOverlay.classList.add("is-open");
  window.setTimeout(() => elements.pricingReferenceTableCloseButton?.focus(), 0);
}

function closePricingReferenceTableOverlay() {
  if (!elements.pricingReferenceTableOverlay) return;
  elements.pricingReferenceTableOverlay.classList.remove("is-open");
  elements.pricingReferenceTableOverlay.hidden = true;
}

function handlePricingReferencePreviewInput(event) {
  const input = event.target.closest("[data-preview-field]");
  if (!input || !state.pendingPricingReference?.items) return;
  const rowIndex = Number(input.dataset.previewRow);
  const field = input.dataset.previewField;
  if (!Number.isInteger(rowIndex) || !state.pendingPricingReference.items[rowIndex]) return;
  state.pendingPricingReference.items[rowIndex][field] = input.value;
  refreshPricingReferencePreviewValidity(state.pendingPricingReference);
  const row = input.closest("tr");
  const status = row?.querySelector(".pricing-preview-status");
  if (status) {
    const statusText = state.pendingPricingReference.items[rowIndex].warning || "OK";
    status.textContent = statusText;
    status.className = `pricing-preview-status ${pricingReferenceStatusClass(statusText)}`;
  }
  updatePricingReferenceGuidanceDisplays(state.pendingPricingReference);
}

function renderPricingReferencePreview(result = null, options = {}) {
  if (!elements.pricingReferencePreview) return;
  if (!result) {
    elements.pricingReferencePreview.innerHTML = "";
    renderPricingReferenceTableOverlay(null);
    setPricingReferenceSaveButtonState({ canSave: false, reason: pricingReferenceSaveBlockReason(null) });
    return;
  }
  if (String(result.layout || "") === "importing") {
    state.pendingPricingReference = null;
    elements.pricingReferencePreview.className = "pricing-reference-preview importing";
    elements.pricingReferencePreview.innerHTML = `
      <div class="pricing-reference-import-overlay" role="status" aria-live="polite">
        <span class="pricing-reference-spinner" aria-hidden="true"></span>
        <strong>Preparing import preview</strong>
        <p>Please wait while AI reads the file and builds editable pricing rows.</p>
      </div>
    `;
    setPricingReferenceSaveButtonState({
      canSave: false,
      busy: true,
      reason: "Import preview is still being prepared.",
    });
    return;
  }
  result.items = sortPricingReferencePreviewItems((result.items || []).map(normalizePricingReferencePreviewItem));
  const selectedCurrencyMode = elements.pricingReferenceCurrency?.value;
  const modalCurrency = pricingReferenceModalCurrency();
  const currency = selectedCurrencyMode === CUSTOM_CURRENCY_VALUE
    ? modalCurrency || "Custom"
    : normalizeCurrencyLabel(result.currency || modalCurrency || DEFAULT_CURRENCY_LABEL);
  if (options.syncCurrencyControls && result.currency) setPricingReferenceCurrencyControls(result.currency);
  refreshPricingReferencePreviewValidity(result);
  const required = pricingReferenceRequiredColumns();
  const missing = Array.isArray(result.missing) ? result.missing : [];
  const found = required.filter((column) => !missing.includes(column));
  const errors = result.errors || [];
  const warnings = result.warnings || [];
  const canSave = Boolean(result.canSave);
  const isAiNormalizationState = String(result.layout || "").startsWith("ai-normalization");
  const displayMessages = errors.length && isAiNormalizationState
    ? [genericFailureMessage(result)]
    : errors.concat(warnings);
  const tone = errors.length ? "error" : warnings.length || !canSave ? "warn" : "ok";
  const title = isAiNormalizationState && errors.length ? "Import preview failed" : isAiNormalizationState ? "AI import needs setup" : errors.length ? "Import preview needs review" : canSave ? "Editable import preview" : "Template preview";
  const rowCount = result.rowCount ?? result.items.length;
  const statusLabel = canSave ? "Ready" : errors.length ? "Needs fix" : "Review";
  const metrics = [
    ["Layout", result.layout || "normalized pricing reference"],
    ["Currency", currency],
    ["Rows detected", rowCount],
    ...(isAiNormalizationState ? [] : [
      ["Columns found", found.join(", ") || "None"],
      ["Columns missing", missing.join(", ") || "None"],
    ]),
    ["Skipped rows", result.skipped || 0],
  ];
  elements.pricingReferencePreview.className = `pricing-reference-preview ${tone}`;
  elements.pricingReferencePreview.innerHTML = `
    <div class="pricing-reference-preview-header">
      <div>
        <span class="pricing-reference-preview-kicker">Import check</span>
        <strong class="pricing-reference-preview-title">${escapeHtml(title)}</strong>
      </div>
      <span class="pricing-reference-preview-status-badge">${escapeHtml(statusLabel)}</span>
    </div>
    <div class="pricing-reference-preview-metrics">
      ${metrics.map(([label, value]) => `
        <div class="pricing-reference-preview-metric">
          <span>${escapeHtml(label)}</span>
          <strong>${escapeHtml(value)}</strong>
        </div>
      `).join("")}
    </div>
    ${displayMessages.length ? `<div class="pricing-reference-preview-messages">${displayMessages.map((message) => `<p>${escapeHtml(message)}</p>`).join("")}</div>` : ""}
    <p class="pricing-reference-save-guidance ${canSave ? "is-ok" : "is-blocked"}">${escapeHtml(pricingReferenceSaveGuidance(result))}</p>
    ${result.items.length ? `<button class="secondary-button pricing-reference-table-open" type="button" data-pricing-reference-table-open>Review ${result.items.length} Row${result.items.length === 1 ? "" : "s"}</button>` : ""}
  `;
  renderPricingReferenceTableOverlay(result);
  setPricingReferenceSaveButtonState({
    canSave,
    busy: state.pricingReferenceImportBusy,
    reason: pricingReferenceSaveBlockReason(result),
  });
  if (options.scrollIntoView) {
    window.requestAnimationFrame(() => {
      elements.pricingReferencePreview?.scrollIntoView({ block: "nearest", inline: "nearest" });
    });
  }
}


function pricingReferenceModalTax() {
  return {
    label: normalizeTaxLabel(elements.pricingReferenceTaxLabel?.value),
    rate: taxRateFromPercentInput(elements.pricingReferenceTaxRate?.value, DEFAULT_TAX_RATE),
  };
}

function pricingReferenceModalCurrency() {
  const selected = elements.pricingReferenceCurrency?.value;
  if (selected === CUSTOM_CURRENCY_VALUE) return customCurrencyInputIsValid() ? normalizedCustomCurrencyInput() : "";
  return supportedCurrencyLabel(selected);
}

function resetPricingReferenceTaxInputs() {
  if (elements.pricingReferenceTaxLabel) elements.pricingReferenceTaxLabel.value = "GST";
  if (elements.pricingReferenceTaxRate) elements.pricingReferenceTaxRate.value = "9";
  setPricingReferenceCurrencyControls(DEFAULT_CURRENCY_LABEL);
}

async function validatePricingReferenceFile(file) {
  if (!file) return pricingReferenceValidationResult([], [], 0, "");
  if (file.size > MAX_PRICING_REFERENCE_FILE_BYTES) {
    return { ...pricingReferenceValidationResult([], [], 0, file.name), errors: ["Pricing reference file is larger than 10 MB."] };
  }
  const extension = file.name.split(".").pop().toLowerCase();
  if (!["csv", "xlsx", "md"].includes(extension)) {
    return { ...pricingReferenceValidationResult([], [], 0, file.name), errors: ["Upload a .xlsx, .csv, or .md pricing reference file."] };
  }
  const { ok, data } = await postJson("/api/settings/pricing-references/import-preview", {
    filename: file.name,
    data_url: await fileToDataUrl(file),
    currency: pricingReferenceModalCurrency(),
    tax: pricingReferenceModalTax(),
  });
  if (ok) return data;
  return {
    ...pricingReferenceValidationResult([], [], 0, file.name),
    errors: genericFailureMessages(data),
    error_reference: errorReferenceFrom(data),
  };
}

async function downloadPricingReferenceTemplate(event) {
  event.preventDefault();
  if (state.pricingReferenceImportBusy || !canManagePricingReferences()) {
    return;
  }
  state.pendingPricingReference = null;
  state.pricingReferenceImportBusy = false;
  state.pricingReferenceImportToken = "";
  renderPricingReferencePreview(null);
  const templateUrl = `/api/pricing-reference/template.xlsx?template=examples-v3&t=${Date.now()}`;
  try {
    const link = document.createElement("a");
    link.href = templateUrl;
    link.download = "swooshz-pricing-reference-template.xlsx";
    document.body.appendChild(link);
    link.click();
    link.remove();
  } catch (error) {
    window.location.href = templateUrl;
  }
}

function openPricingReferenceModal() {
  state.pendingPricingReference = null;
  state.pricingReferenceImportBusy = false;
  state.pricingReferenceImportToken = "";
  if (elements.pricingReferenceName) elements.pricingReferenceName.value = "";
  if (elements.pricingReferenceFile) elements.pricingReferenceFile.value = "";
  if (elements.pricingReferenceFileName) elements.pricingReferenceFileName.textContent = "No file chosen";
  resetPricingReferenceTaxInputs();
  renderPricingReferencePreview(null);
  renderPricingReferenceDeleteOptions();
  elements.pricingReferenceModal.hidden = false;
  elements.pricingReferenceModal.classList.add("is-open");
  const canManage = setPricingReferenceModalAccessState();
  window.setTimeout(() => (canManage ? elements.pricingReferenceName : elements.pricingReferenceCancelButton)?.focus(), 0);
}

function closePricingReferenceModal() {
  if (state.pricingReferenceImportBusy) {
    return;
  }
  closePricingReferenceTableOverlay();
  elements.pricingReferenceModal.classList.remove("is-open");
  elements.pricingReferenceModal.hidden = true;
  state.pendingPricingReference = null;
  state.pricingReferenceImportBusy = false;
  state.pricingReferenceImportToken = "";
}

async function savePricingReferenceFromModal(event) {
  event.preventDefault();
  if (!canManagePricingReferences()) {
    setPricingReferenceSaveButtonState({
      canSave: false,
      reason: pricingReferenceNoAccessReason(),
    });
    return;
  }
  if (state.pricingReferenceImportBusy) {
    setPricingReferenceSaveButtonState({
      busy: true,
      reason: "Import preview is still being prepared.",
    });
    return;
  }
  const name = elements.pricingReferenceName.value.trim();
  const tax = pricingReferenceModalTax();
  const currency = pricingReferenceModalCurrency();
  const result = state.pendingPricingReference;
  const canSave = result?.canSave ?? Boolean(result && !result.errors.length && result.items.length);
  const saveBlockReason = pricingReferenceSaveBlockReason(result);
  if (!name || !result || !canSave || saveBlockReason) {
    renderPricingReferencePreview(result || pricingReferenceValidationResult([], [], 0, ""));
    return;
  }
  const previousPricingReferenceId = state.pricingReferenceId;
  const previousPricingReferenceSource = state.pricingReferenceSource;
  const { ok, data } = await postJson("/api/settings/pricing-references", {
    id: safeId(name, `pricing-ref-${Date.now().toString(36)}`),
    label: name,
    description: `Imported from ${result.sourceName || "settings upload"}`,
    tax,
    currency,
    items: result.items,
  });
  if (!ok) {
    renderPricingReferencePreview({ ...result, errors: genericFailureMessages(data), error_reference: errorReferenceFrom(data) });
    return;
  }
  await loadProfiles();
  state.pricingReferenceId = previousPricingReferenceId;
  state.pricingReferenceSource = previousPricingReferenceSource;
  syncSelectedPricingReference();
  renderProfileOptions();
  closePricingReferenceModal();
  syncControlStates();
}

async function deleteRepoPricingReference(referenceId) {
  const reference = state.pricingReferences.find((item) => item.id === referenceId && String(item.source || "bundled") === "bundled");
  if (!reference) return;
  const reason = protectedPricingReferenceReason(reference);
  if (reason) {
    updatePricingReferenceDeleteButton();
    return;
  }
  const label = reference.label || reference.id;
  const confirmed = window.prompt(`Type ${label} to delete this pricing reference.`) === label;
  if (!confirmed) return;
  const { ok, data } = await fetch(`/api/settings/pricing-references/${encodeURIComponent(reference.id)}`, {
    method: "DELETE",
    headers: state.csrfToken ? { [state.csrfHeaderName]: state.csrfToken } : {},
  }).then(async (response) => ({ ok: response.ok, data: await response.json().catch(() => ({})) }));
  if (!ok) {
    window.alert(genericFailureMessages(data).join("\n"));
    updatePricingReferenceDeleteButton();
    return;
  }
  state.pricingReferences = mergePricingReferences(Array.isArray(data.pricing_references) ? data.pricing_references : state.pricingReferences);
  if (state.pricingReferenceId === reference.id) {
    const fallback = defaultPricingReference() || state.pricingReferences[0] || null;
    state.pricingReferenceId = fallback?.id || "";
    state.pricingReferenceSource = fallback ? pricingReferenceSelectionFromValue(pricingReferenceSelectValue(fallback)).source : "";
    clearGeneratedQuoteState();
  }
  syncSelectedPricingReference();
  renderProfileOptions();
  renderPricingReferenceDeleteOptions();
  syncControlStates();
}

async function deleteSelectedPricingReference() {
  const selected = deletionPricingReference();
  if (!selected || !canDeleteSelectedPricingReference()) {
    updatePricingReferenceDeleteButton();
    return;
  }
  await deleteRepoPricingReference(selected.id);
}

async function loadProfiles() {
  const { ok, data } = await getJson("/api/profiles");
  if (ok && Array.isArray(data.profiles)) {
    state.profiles = data.profiles;
    state.defaultPricingReferenceId = data.default_pricing_reference_id || DEFAULT_PRICING_REFERENCE_ID;
    state.pricingReferences = mergePricingReferences(Array.isArray(data.pricing_references) ? data.pricing_references : []);
    if (state.pricingReferenceId) {
      syncSelectedPricingReference();
    }
  }
  renderProfileOptions();
}

function clearGeneratedQuoteState() {
  state.quoteBasis = { ...EMPTY_BASIS };
  state.quoteBasisSections = [];
  state.lineItems = [];
  state.outputRows = [];
  state.originalOutputRows = [];
  state.outputErrors = [];
  state.analysisFindings = [];
  state.blockingClarificationQuestions = [];
  state.basisConfirmed = false;
  state.aiFailed = false;
  state.draftSource = "";
  state.lastAnalysisMode = "standard";
  state.activeJob = null;
  state.originalAnalysisSnapshot = null;
  renderBasisEmptyState();
  clearAiFailureBanner();
  renderMessages([]);
  setDownloadFiles([]);
  renderMatchSummary({});
  renderPricingMatches([]);
  clearPricingReviewMessages();
  setResultStatus("No job yet");
}

function handleProfileSelectionChange() {
  const nextSelection = pricingReferenceSelectionFromValue(elements.profileSelect.value || "");
  if (nextSelection.pricingReferenceId === state.pricingReferenceId && nextSelection.source === state.pricingReferenceSource) {
    return;
  }
  state.pricingReferenceId = nextSelection.pricingReferenceId;
  state.pricingReferenceSource = nextSelection.source;
  syncSelectedPricingReference();
  renderProfileOptions();
  clearGeneratedQuoteState();
  setWorkflowStage(state.images.length ? (canStartAnalysis() ? "ready_to_analyze" : "details_review") : "needs_images");
  syncControlStates();
}

async function setSampleDetails() {
  if (state.isBooting || state.isAnalysisRunning || state.isGenerating) return;
  elements.sampleDetailsButton.disabled = true;
  try {
    const { ok, data } = await getJson(`/api/samples/${DEFAULT_SAMPLE_ID}`);
    if (!ok) {
      return;
    }
    if (!state.profiles.length) await loadProfiles();
    state.profileId = data.profile_id || DEFAULT_PROFILE_ID;
    state.pricingReferenceId = data.pricing_reference_id || currentProfile()?.default_pricing_reference || state.defaultPricingReferenceId || DEFAULT_PRICING_REFERENCE_ID;
    syncSelectedPricingReference();
    renderProfileOptions();
    renderPresetOptions();
    loadConfiguredProfilePreset({ silent: true });
    updateGeneratorCopy();
    applyQuoteDetails(data.details || {}, { partial: true });
    state.images = Array.isArray(data.images) ? data.images.slice(0, MAX_REFERENCE_IMAGES) : [];
    await persistSessionFiles(sessionFileRecordsFromImages(state.images)).catch(() => {});
    renderFiles();
    setImageUploadStatus(`${state.images.length} sample reference file${state.images.length === 1 ? "" : "s"} loaded.`);
    setWorkflowStage(state.images.length ? "ready_to_analyze" : "needs_images");
    syncControlStates();
  } finally {
    elements.sampleDetailsButton.disabled = false;
    syncControlStates();
  }
}

function buildPayload(options = {}) {
  syncRichTextSources();
  const generator = currentGenerator();
  const pricingReference = currentPricingReference();
  const profileId = resolvedProfileIdForPayload();
  const includeBoothDimensions = options.includeBoothDimensions !== false;
  const includeDraftContext = options.includeDraftContext !== false;
  const project = {
    title: elements.projectTitle.value.trim(),
  };
  if (includeBoothDimensions) {
    project.booth_width = state.boothDimensions.booth_width;
    project.booth_depth = state.boothDimensions.booth_depth;
    project.booth_size = state.boothDimensions.booth_size;
    project.dimension_source = state.boothDimensions.dimension_source;
  }
  return {
    profile_id: profileId,
    pricing_reference_id: pricingReference?.id || state.pricingReferenceId || "",
    pricing_reference: pricingReference ? {
      ...pricingReference,
      source: "bundled",
      tax: pricingReference.tax || selectedPricingReferenceTax(),
      currency: selectedPricingReferenceCurrency(),
    } : { id: state.pricingReferenceId || "", source: "bundled", tax: selectedPricingReferenceTax(), currency: selectedPricingReferenceCurrency() },
    analysis_mode: normalizeAnalysisMode(options.analysisMode || state.pendingAnalysisMode),
    generator_label: generator.label,
    images: state.images.slice(0, MAX_REFERENCE_IMAGES),
    confirmed: true,
    quote_date: elements.quoteDate.value,
    project_number: elements.projectNumber.value.trim(),
    client: {
      name: elements.clientName.value.trim(),
      attention: elements.clientAttention.value.trim(),
      title: elements.clientTitle.value.trim(),
      address: elements.clientAddress.value,
    },
    project,
    company: {
      name: elements.quoteCompanyName.value.trim(),
      header_details: elements.headerDetails.value,
      logo_data_url: state.headerLogo ? state.headerLogo.data_url : "",
      logo_name: state.headerLogo ? state.headerLogo.name : "",
      logo_type: state.headerLogo ? state.headerLogo.type : "",
    },
    tax: collectTaxDetails(),
    user_feedback: state.pendingFeedback,
    quote_basis: includeDraftContext ? { ...state.quoteBasis, ...quoteBasisFromSections(state.quoteBasisSections) } : {},
    quote_basis_sections: includeDraftContext ? cloneQuoteBasisSections(state.quoteBasisSections) : [],
    line_items: includeDraftContext ? (state.outputRows.length ? outputRowsToLineItems(state.outputRows) : state.lineItems) : [],
    analysis_findings: state.analysisFindings,
    blocking_clarification_questions: state.blockingClarificationQuestions,
    quote_text: {
      terms_heading: elements.termsHeading.value.trim(),
      payment_terms: splitLines(elements.paymentTerms.value),
      notes_heading: elements.notesHeading.value.trim(),
      standard_notes: splitLines(elements.standardNotes.value),
      acceptance_text: elements.acceptanceText.value.trim(),
      person_label: elements.personLabel.value.trim(),
      stamp_label: elements.stampLabel.value.trim(),
      date_label: elements.dateLabel.value.trim(),
    },
    signature: {
      koncept_signatory: elements.konceptSignatory.value.trim(),
      koncept_title: elements.konceptTitle.value.trim(),
      koncept_date_label: elements.konceptDateLabel.value.trim(),
    },
    rich_text: collectRichTextDetails(),
  };
}

function buildLineItemNormalizePayload() {
  const pricingReference = currentPricingReference();
  const profileId = resolvedProfileIdForPayload();
  return {
    profile_id: profileId,
    pricing_reference_id: pricingReference?.id || state.pricingReferenceId || "",
    pricing_reference: pricingReference ? {
      id: pricingReference.id || state.pricingReferenceId || "",
      label: pricingReference.label || "",
      source: pricingReference.source || "bundled",
      tax: pricingReference.tax || selectedPricingReferenceTax(),
      currency: selectedPricingReferenceCurrency(),
    } : {
      id: state.pricingReferenceId || "",
      source: "bundled",
      tax: selectedPricingReferenceTax(),
      currency: selectedPricingReferenceCurrency(),
    },
    project: {
      booth_width: state.boothDimensions.booth_width,
      booth_depth: state.boothDimensions.booth_depth,
      booth_size: state.boothDimensions.booth_size,
      dimension_source: state.boothDimensions.dimension_source,
    },
    quote_basis: { ...state.quoteBasis, ...quoteBasisFromSections(state.quoteBasisSections) },
    quote_basis_sections: cloneQuoteBasisSections(state.quoteBasisSections),
    line_items: state.lineItems.map(normalizeLineItem),
  };
}

function setResultStatus(label, tone = "") {
  elements.resultStatus.textContent = label;
  elements.resultStatus.classList.remove("is-ok", "is-warn", "is-bad");
  if (tone) elements.resultStatus.classList.add(tone);
}

function renderMessages(messages = [], tone = "") {
  elements.messageList.innerHTML = messages
    .map((message) => `<div class="message ${tone}">${escapeHtml(message)}</div>`)
    .join("");
}

function setDownloadFiles(files = []) {
  const excelFile = files.find((file) => /\.xlsx$/i.test(file.name || "")) || files[0] || null;
  state.downloadFile = excelFile;
  updateDownloadButton();
}

function updateDownloadButton() {
  if (!elements.sideDownloadButton) return;
  const file = state.downloadFile;
  const validation = outputRowsValid();
  const enabled = state.activeSidePanel === "output" && validation.valid && !state.isGenerating;
  elements.sideDownloadButton.classList.toggle("is-disabled", !enabled);
  elements.sideDownloadButton.setAttribute("aria-disabled", String(!enabled));
  elements.sideDownloadButton.tabIndex = enabled ? 0 : -1;
  elements.sideDownloadButton.href = enabled && file?.url ? file.url : "#";
  elements.sideDownloadButton.download = file?.url ? file.name || "quotation.xlsx" : "";
  elements.sideDownloadButton.textContent = "Download Excel";
}

function downloadCurrentExcelFile(file = state.downloadFile) {
  if (!file?.url) return false;
  try {
    const link = document.createElement("a");
    link.href = file.url;
    link.download = file.name || "quotation.xlsx";
    document.body.appendChild(link);
    link.click();
    link.remove();
  } catch (error) {
    window.location.href = file.url;
  }
  return true;
}

function pricingMatchStatus(row = {}) {
  const status = typeof row === "string" ? row : row.status;
  return String(status || "").trim().toLowerCase();
}

function pricingStatusLabel(status = "") {
  const labels = {
    matched: "Catalog match",
    "matched-from-ambiguous": "Ambiguous match selected",
    ambiguous: "Ambiguous match",
    "manual-display": "Manual display price",
    "quantity-review": "Quantity needs review",
    custom: "Custom manual price",
    "needs-pricing": "Needs pricing",
    unmatched: "Unmatched",
  };
  const normalized = pricingMatchStatus(status);
  return labels[normalized] || String(status || "").trim() || "Unknown";
}

function numberOrNull(value) {
  if (value === "" || value === null || value === undefined) return null;
  if (String(value).trim().toLowerCase() === "included") return null;
  const numeric = Number(String(value ?? "").replaceAll(",", "").trim());
  return Number.isFinite(numeric) ? numeric : null;
}

function orderNumber(value) {
  const number = numberOrNull(value);
  return number !== null && number > 0 ? Math.trunc(number) : null;
}

function unitPriceEditKind(value) {
  const text = String(value ?? "").trim();
  if (!text) return "blank";
  if (text.toLowerCase() === "included") return "included";
  return numberOrNull(text) === null ? "invalid" : "number";
}

function effectiveOutputUnitPrice(row = {}) {
  const manual = numberOrNull(row.unit_price_override);
  if (manual !== null) return manual;
  return numberOrNull(row.catalog_unit_price);
}

function formatAmount(value) {
  if (value === "" || value === null || value === undefined) return "";
  const numeric = Number(value);
  if (!Number.isFinite(numeric)) return "";
  return numeric.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function recalculateOutputRow(row = {}) {
  const quantity = numberOrNull(row.quantity);
  const priceMode = row.price_mode === "Included" ? "Included" : "Priced";
  const unitPrice = effectiveOutputUnitPrice({ ...row, price_mode: priceMode });
  const hasUsablePrice = unitPrice !== null && unitPrice >= 0;
  return {
    ...row,
    price_mode: priceMode,
    amount: priceMode === "Included" ? 0 : (quantity !== null && hasUsablePrice ? Math.round(quantity * unitPrice * 100) / 100 : ""),
  };
}

function outputComparableText(value = "") {
  return String(value || "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function outputRowSectionMatchesBasis(row = {}, sectionTitle = "") {
  if (!String(sectionTitle || "").trim()) return true;
  const basisSection = sectionTitleKey(normalizeQuoteBasisTitle(sectionTitle));
  if (!basisSection) return true;
  const rowSection = sectionTitleKey(row.section);
  if (!rowSection) return true;
  return rowSection === basisSection;
}

function outputRowCoversBasisLine(row = {}, lineText = "", sectionTitle = "") {
  if (!outputRowSectionMatchesBasis(row, sectionTitle)) return false;
  const rowText = outputComparableText(row.description);
  const basisText = outputComparableText(lineText);
  if (!rowText || !basisText) return false;
  if (rowText.includes(basisText) || basisText.includes(rowText)) return true;
  const basisWords = basisText.split(" ").filter((word) => word.length > 3);
  if (basisWords.length < 4) return false;
  const rowWords = new Set(rowText.split(" "));
  const overlap = basisWords.filter((word) => rowWords.has(word)).length;
  return overlap / basisWords.length >= 0.85;
}

function outputRowCoversBasisEntry(row = {}, line = {}, sectionTitle = "") {
  const lineIds = [line.id, line.source_line_item_id].map((value) => String(value || "").trim()).filter(Boolean);
  const sourceId = String(row.source_basis_line_id || "").trim();
  if (sourceId && lineIds.length) return lineIds.includes(sourceId);
  const keywordMatch = row.pricing_keyword && line.pricing_keyword && String(row.pricing_keyword) === String(line.pricing_keyword);
  if (keywordMatch) return true;
  return outputRowCoversBasisLine(row, line.text || "", sectionTitle)
    || outputRowCoversBasisLine(row, line.catalog_description || "", sectionTitle);
}

function basisLineAllowsOutput(line = {}) {
  const basisTag = normalizeBasisTag(line.tag);
  if (basisTag === "Exclude" || basisTag === "Confirm") return false;
  if (basisTag === "Include") return true;
  return isCustomPricingBasisLine(line) && Boolean(line.custom_confirmed);
}

function outputRowAllowedByBasis(row = {}) {
  const reviewedSections = normalizeQuoteBasisSections(state.quoteBasisSections);
  const hasReviewedBasis = reviewedSections.some((section) => (section.lines || []).length > 0);
  const sections = hasReviewedBasis ? reviewedSections : basisSections(state.quoteBasisSections);
  const sourceId = String(row.source_basis_line_id || "").trim();
  let matched = false;
  let allowed = false;
  sections.forEach((section) => {
    (section.lines || []).forEach((line) => {
      const lineIds = [line.id, line.source_line_item_id].map((value) => String(value || "").trim()).filter(Boolean);
      const sourceMatch = sourceId && lineIds.includes(sourceId);
      const keywordMatch = row.pricing_keyword && line.pricing_keyword && String(row.pricing_keyword) === String(line.pricing_keyword);
      const sectionTitle = section.title || section.id || "";
      if (sourceId && lineIds.length && !sourceMatch) return;
      const textMatch = outputRowCoversBasisLine(row, line.text || "", sectionTitle) || outputRowCoversBasisLine(row, line.catalog_description || "", sectionTitle);
      if (!sourceMatch && !keywordMatch && !textMatch) return;
      matched = true;
      allowed = allowed || basisLineAllowsOutput(line);
    });
  });
  return hasReviewedBasis ? matched && allowed : !matched || allowed;
}

function matchingAllowedBasisLineForOutputRow(row = {}) {
  const reviewedSections = normalizeQuoteBasisSections(state.quoteBasisSections);
  const hasReviewedBasis = reviewedSections.some((section) => (section.lines || []).length > 0);
  if (!hasReviewedBasis) return null;
  const sourceId = String(row.source_basis_line_id || "").trim();
  for (const section of reviewedSections) {
    for (const line of (section.lines || [])) {
      if (!basisLineAllowsOutput(line)) continue;
      const lineIds = [line.id, line.source_line_item_id].map((value) => String(value || "").trim()).filter(Boolean);
      const sourceMatch = sourceId && lineIds.includes(sourceId);
      const keywordMatch = row.pricing_keyword && line.pricing_keyword && String(row.pricing_keyword) === String(line.pricing_keyword);
      const sectionTitle = section.title || section.id || "";
      if (sourceId && lineIds.length && !sourceMatch) continue;
      const textMatch = outputRowCoversBasisLine(row, line.text || "", sectionTitle)
        || outputRowCoversBasisLine(row, line.catalog_description || "", sectionTitle);
      if (sourceMatch || keywordMatch || textMatch) return line;
    }
  }
  return null;
}

function inheritBasisOutputFields(row = {}) {
  const line = matchingAllowedBasisLineForOutputRow(row);
  if (!line) return row;
  const next = { ...row };
  if (line.quantity !== undefined && line.quantity !== null && String(line.quantity).trim() !== "") {
    next.quantity = line.quantity;
  }
  if (normalizeUnit(line.unit || "")) {
    next.unit = normalizeUnit(line.unit);
  }
  if (!next.source_basis_line_id && (line.id || line.source_line_item_id)) {
    next.source_basis_line_id = line.id || line.source_line_item_id;
  }
  if ((next.catalog_unit_price === "" || next.catalog_unit_price === null || next.catalog_unit_price === undefined) && line.catalog_unit_price !== undefined && line.catalog_unit_price !== null && String(line.catalog_unit_price).trim() !== "") {
    next.catalog_unit_price = line.catalog_unit_price;
  }
  if (!next.catalog_description && line.catalog_description) {
    next.catalog_description = line.catalog_description;
  }
  if (!next.pricing_reference_description && line.pricing_reference_description) {
    next.pricing_reference_description = line.pricing_reference_description;
  }
  return normalizeOutputRow(next);
}

function outputRowDedupeKey(row = {}) {
  const sourceId = String(row.source_basis_line_id || "").trim();
  if (sourceId) return `source:${sourceId}`;
  return [
    sectionTitleKey(row.section),
    outputComparableText(row.description),
    String(row.pricing_keyword || "").trim(),
    String(row.quantity || "").trim(),
    normalizeUnit(row.unit || ""),
  ].join("|");
}

function dedupeOutputRows(rows = []) {
  const seen = new Set();
  const deduped = [];
  (Array.isArray(rows) ? rows : []).forEach((row) => {
    const key = outputRowDedupeKey(row);
    if (key && seen.has(key)) return;
    if (key) seen.add(key);
    deduped.push(row);
  });
  return deduped;
}

function includedBasisOutputRows(existingRows = []) {
  const rows = [];
  basisSections(state.quoteBasisSections).forEach((section) => {
    (section.lines || []).forEach((line) => {
      const customPricing = isCustomPricingBasisLine(line);
      if (!basisLineAllowsOutput(line)) return;
      const text = String(line.text || "").trim();
      if (!text) return;
      const description = customPricing ? text : outputCatalogDescription(line);
      const isBoothSizeInfo = /\bbooth size\b/i.test(text) && !line.pricing_keyword;
      if (isBoothSizeInfo && !customPricing) return;
      const sectionTitle = section.title || section.id || "";
      const duplicate = existingRows.some((row) => outputRowCoversBasisEntry(row, line, sectionTitle))
        || rows.some((row) => outputRowCoversBasisEntry(row, line, sectionTitle));
      if (duplicate) return;
      rows.push(normalizeOutputRow({
        section: normalizeCategoryTitle(section.title || section.id || "Quote Basis"),
        description,
        quantity: line.quantity || (isBoothSizeInfo ? 1 : ""),
        unit: normalizeUnit(line.unit || (isBoothSizeInfo ? "lot" : "")),
        price_mode: isBoothSizeInfo ? "Included" : "Priced",
        unit_price_override: "",
        catalog_unit_price: line.catalog_unit_price ?? "",
        catalog_description: line.catalog_description || "",
        pricing_reference_description: line.pricing_reference_description || "",
        pricing_keyword: line.pricing_keyword || "",
        source_basis_line_id: line.id || line.source_line_item_id || "",
        category_order: line.category_order ?? "",
        item_order: line.item_order ?? "",
        status: customPricing ? "custom" : "needs-pricing",
      }));
    });
  });
  return rows;
}

function outputRowFromLineItem(item = {}) {
  const normalized = normalizeLineItem(item);
  const description = normalized.pricing_keyword ? outputCatalogDescription(normalized) : normalized.description;
  return normalizeOutputRow({
    section: normalized.section,
    description,
    quantity: normalized.quantity,
    unit: normalized.unit,
    price_mode: normalized.price_mode,
    unit_price_override: normalized.unit_price_override || "",
    catalog_unit_price: normalized.catalog_unit_price || "",
    catalog_description: normalized.catalog_description || "",
    pricing_reference_description: normalized.pricing_reference_description || "",
    pricing_keyword: normalized.pricing_keyword,
    source_basis_line_id: normalized.source_basis_line_id,
    category_order: normalized.category_order,
    item_order: normalized.item_order,
    status: normalized.status || "",
  });
}

function outputQuantityPartsFromPricingMatch(row = {}) {
  const rawQuantity = String(row.quantity || "").trim();
  const explicitUnit = normalizeUnit(row.unit || "");
  if (!rawQuantity) return { quantity: "", unit: explicitUnit };
  if (explicitUnit) {
    const escapedUnit = explicitUnit.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    const quantity = rawQuantity.replace(new RegExp(`\\s+${escapedUnit}$`, "i"), "").trim();
    return { quantity: quantity || rawQuantity, unit: explicitUnit };
  }
  const match = rawQuantity.match(/^(-?\d+(?:\.\d+)?)\s+(.+)$/);
  if (match) return { quantity: match[1], unit: normalizeUnit(match[2]) };
  return { quantity: rawQuantity, unit: "" };
}

function outputRowFromPricingMatch(row = {}) {
  const status = pricingMatchStatus(row);
  const amount = String(row.amount ?? "").trim();
  const quantityParts = outputQuantityPartsFromPricingMatch(row);
  const manualDisplayAmount = status === "manual-display" ? numberOrNull(amount) : null;
  const quantity = numberOrNull(quantityParts.quantity);
  const manualDisplayUnitPrice = manualDisplayAmount !== null
    ? (quantity !== null && quantity > 0 ? Math.round((manualDisplayAmount / quantity) * 100) / 100 : manualDisplayAmount)
    : "";
  const unitPrice = manualDisplayAmount !== null ? manualDisplayUnitPrice : row.unit_price || row.unit_price_override || "";
  const referenceDescription = pricingReferenceLineText(row.pricing_reference_description || row.catalog_description || "");
  return normalizeOutputRow({
    section: row.section,
    description: cleanCustomerQuoteLineText(row.description || row.catalog_description || ""),
    quantity: quantityParts.quantity,
    unit: quantityParts.unit,
    price_mode: status === "included" ? "Included" : "Priced",
    unit_price_override: status === "manual-price" || status === "manual-display" ? unitPrice : "",
    catalog_unit_price: status === "manual-price" || status === "manual-display" ? "" : unitPrice,
    catalog_description: row.catalog_description || "",
    pricing_reference_description: referenceDescription,
    pricing_keyword: row.keyword || row.pricing_keyword || "",
    source_basis_line_id: row.source_basis_line_id || "",
    category_order: row.category_order ?? "",
    item_order: row.item_order ?? "",
    status: row.status,
  });
}

function categoryOrderValue(row = {}) {
  return orderNumber(row.category_order) || 999999;
}

function pricingReferenceOrder(row = {}, fallbackIndex = 0) {
  return [
    categoryOrderValue(row),
    orderNumber(row.item_order) || 999999,
    fallbackIndex,
  ];
}

function compareOrderValues(left = [], right = []) {
  const length = Math.max(left.length, right.length);
  for (let index = 0; index < length; index += 1) {
    const difference = (left[index] || 0) - (right[index] || 0);
    if (difference) return difference;
  }
  return 0;
}

function sortOutputRows(rows = state.outputRows) {
  const copy = rows.map((row, index) => ({ row, index }));
  const compareText = (value) => String(value || "").toLowerCase();
  copy.sort((a, b) => {
    const sectionCompare = compareText(a.row.section).localeCompare(compareText(b.row.section));
    const descriptionCompare = compareText(a.row.description).localeCompare(compareText(b.row.description));
    if (state.outputSortMode === "pricing_reference") {
      return compareOrderValues(pricingReferenceOrder(a.row, a.index), pricingReferenceOrder(b.row, b.index))
        || sectionCompare
        || descriptionCompare
        || a.index - b.index;
    }
    if (state.outputSortMode === "name") {
      return descriptionCompare || a.index - b.index;
    }
    if (state.outputSortMode === "category") return sectionCompare || a.index - b.index;
    return sectionCompare || descriptionCompare || a.index - b.index;
  });
  return copy.map((item) => item.row);
}

function outputCellDisplayValue(row = {}, field = "") {
  if (field === "price_mode") return row.price_mode === "Included" ? "Included" : "Priced";
  if (field === "unit_price_override") {
    if (row.price_mode === "Included") return "Included";
    if (numberOrNull(row.unit_price_override) !== null) return formatAmount(row.unit_price_override);
    if (numberOrNull(row.catalog_unit_price) !== null) return formatAmount(row.catalog_unit_price);
    return "???";
  }
  if (field === "amount") {
    if (row.price_mode === "Included") return "0.00";
    return formatAmount(row.amount) || "???";
  }
  return String(row[field] ?? "").trim() || "???";
}

function outputEditorHtml(row = {}, index = 0, field = "") {
  const value = field === "unit_price_override"
    ? row.price_mode === "Included"
      ? "Included"
      : String(row.unit_price_override || row.catalog_unit_price || "")
    : String(row[field] ?? "");
  if (field === "description") {
    return `<textarea class="output-cell-input output-description-input is-editing" data-output-editor-field="${field}" data-output-row="${index}" rows="3">${escapeHtml(value)}</textarea>`;
  }
  if (field === "price_mode") {
    return `
      <select class="output-cell-input is-editing" data-output-editor-field="${field}" data-output-row="${index}">
        <option value="Priced" ${row.price_mode === "Included" ? "" : "selected"}>Priced</option>
        <option value="Included" ${row.price_mode === "Included" ? "selected" : ""}>Included</option>
      </select>
    `;
  }
  if (field === "unit_price_override") {
    const includedButton = row.price_mode !== "Included"
      ? `<button class="output-included-button" type="button" data-output-included-action="true" data-output-row="${index}" aria-label="Mark row ${index + 1} unit price as included" title="Mark unit price as Included">Included</button>`
      : "";
    return `
      <span class="output-unit-price-editor">
        <input class="output-cell-input is-editing" data-output-editor-field="${field}" data-output-row="${index}" value="${escapeHtml(value)}" inputmode="decimal">
        ${includedButton}
      </span>
    `;
  }
  const inputMode = field === "quantity" ? "decimal" : "text";
  return `<input class="output-cell-input is-editing" data-output-editor-field="${field}" data-output-row="${index}" value="${escapeHtml(value)}" inputmode="${inputMode}">`;
}

function renderOutputEditCell(row = {}, index = 0, field = "", extraClass = "") {
  const display = outputCellDisplayValue(row, field);
  const pending = display === "???";
  const unitPriceCell = field === "unit_price_override";
  const className = [
    "output-edit-cell",
    extraClass,
    unitPriceCell ? "output-unit-price-cell" : "",
    pending ? "is-pending" : "",
  ].filter(Boolean).join(" ");
  const cellBody = unitPriceCell
    ? `<span class="output-unit-price-content"><span class="output-cell-text">${escapeHtml(display)}</span></span>`
    : `<span class="output-cell-text">${escapeHtml(display)}</span>`;
  return `
    <td class="${className}" data-output-edit-field="${field}" data-output-row="${index}" tabindex="0" title="Click to edit">
      ${cellBody}
    </td>
  `;
}

function ensureOutputRowsFromLineItems() {
  if (state.outputRows.length) return;
  refreshOutputRowsFromLineItems();
}

function refreshOutputRowsFromLineItems() {
  const generatedRows = state.lineItems.map(outputRowFromLineItem);
  const allowedRows = dedupeOutputRows(generatedRows.filter(outputRowAllowedByBasis).map(inheritBasisOutputFields));
  state.outputRows = sortOutputRows([...allowedRows, ...includedBasisOutputRows(allowedRows)]);
}

function snapshotOutputRows(rows = state.outputRows) {
  return rows.map((row) => normalizeOutputRow({ ...row }));
}

function outputRowsToLineItems(rows = state.outputRows) {
  return rows.map((row) => {
    const next = {
      section: String(row.section || "").trim(),
      description: String(row.description || "").trim(),
      quantity: row.quantity,
      unit: normalizeUnit(row.unit || ""),
      pricing_keyword: row.pricing_keyword || "",
      price_mode: row.price_mode === "Included" ? "Included" : "Priced",
      source_basis_line_id: row.source_basis_line_id || "",
    };
    if (!next.source_basis_line_id) delete next.source_basis_line_id;
    if (next.price_mode === "Included") {
      next.display_price = "Included";
    } else {
      const unitPrice = numberOrNull(row.unit_price_override);
      if (unitPrice !== null) next.unit_price_override = unitPrice;
    }
    return next;
  });
}

function outputRowsValid(rows = state.outputRows) {
  const errors = [];
  rows.forEach((row, index) => {
    const label = `Row ${index + 1}`;
    if (!String(row.description || "").trim()) errors.push(`${label}: Description is required.`);
    const quantity = numberOrNull(row.quantity);
    if (!String(row.unit || "").trim()) errors.push(`${label}: Unit is required.`);
    if (row.price_mode !== "Included" && (quantity === null || quantity <= 0)) {
      errors.push(`${label}: Quantity must be greater than 0.`);
    }
    if (row.price_mode !== "Included") {
      const unitPrice = numberOrNull(row.unit_price_override);
      const catalogUnitPrice = numberOrNull(row.catalog_unit_price);
      const unitPriceKind = unitPriceEditKind(row.unit_price_override);
      if (unitPriceKind === "invalid") {
        errors.push(`${label}: Unit price must be a number or Included.`);
      } else if (unitPrice !== null && unitPrice < 0) {
        errors.push(`${label}: Unit price must be 0 or greater.`);
      } else if (unitPrice === null && catalogUnitPrice === null) {
        errors.push(`${label}: Unit price is required.`);
      }
    }
  });
  return { valid: errors.length === 0 && rows.length > 0, errors };
}

function deleteOutputRow(index) {
  if (!Number.isInteger(index) || index < 0 || !state.outputRows[index]) return;
  const row = state.outputRows[index];
  const confirmed = window.confirm(`Delete output row "${row.description || `Row ${index + 1}`}"?`);
  if (!confirmed) return;
  state.outputRows.splice(index, 1);
  state.lineItems = outputRowsToLineItems();
  state.downloadFile = null;
  const validation = outputRowsValid();
  renderPricingMatches(state.outputRows);
  renderMatchSummary({ pricing_matches: state.outputRows });
  renderOutputValidationMessages(validation.valid ? [] : validation.errors);
  syncControlStates();
}

function renderOutputValidationMessages(errors = state.outputErrors) {
  if (!elements.pricingReviewMessages) return;
  state.outputErrors = errors;
  elements.pricingReviewMessages.innerHTML = "";
}

function rowNeedsManualInput(row = {}) {
  if (!String(row.description || "").trim()) return true;
  if (row.price_mode === "Included") return false;
  const quantity = numberOrNull(row.quantity);
  const unitPrice = effectiveOutputUnitPrice(row);
  const hasPricingKeyword = Boolean(String(row.pricing_keyword || "").trim());
  const amountMissing = outputCellDisplayValue(row, "amount") === "???";
  if (quantity === null || quantity <= 0) return true;
  if (unitPrice === null || unitPrice < 0) return true;
  if (!hasPricingKeyword && numberOrNull(row.unit_price_override) === null) return true;
  return amountMissing;
}

function matchSummaryStats(rows = []) {
  const safeRows = Array.isArray(rows) ? rows : [];
  const sections = new Set(safeRows.map((row) => String(row.section || "").trim()).filter(Boolean)).size;
  const needsManualInput = safeRows.filter(rowNeedsManualInput).length;
  const pricedRows = safeRows.filter((row) => row.price_mode === "Included" || !rowNeedsManualInput(row)).length;
  const totalPending = needsManualInput > 0;
  const total = safeRows.reduce((sum, row) => {
    const amount = Number(String(row.amount || "").replaceAll(",", ""));
    return Number.isFinite(amount) ? sum + amount : sum;
  }, 0);
  return { sections, pricedRows, needsManualInput, total, totalPending };
}

function formatSubtotalValue(stats = {}) {
  const total = Number(stats.total || 0);
  const totalText = `${selectedPricingReferenceCurrency()} ${total.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
  return stats.totalPending ? `${totalText} + ???` : totalText;
}

function outputPricingSourceLabel() {
  const reference = currentPricingReference();
  if (!reference) return "Pricing reference";
  return reference.label || "Pricing reference";
}

function outputHeaderStatus(rows = state.outputRows) {
  const safeRows = Array.isArray(rows) ? rows : [];
  if (!safeRows.length) return { label: "Draft", className: "is-empty" };
  const stats = matchSummaryStats(safeRows);
  return stats.needsManualInput > 0
    ? { label: "Needs pricing", className: "is-warn" }
    : { label: "Ready", className: "is-ok" };
}

function updateOutputHeader(rows = state.outputRows) {
  const safeRows = Array.isArray(rows) ? rows : [];
  const status = outputHeaderStatus(safeRows);
  if (elements.outputStatusPill) {
    elements.outputStatusPill.textContent = status.label;
    elements.outputStatusPill.className = `output-status-pill ${status.className}`;
  }
  if (elements.outputSourceLabel) {
    elements.outputSourceLabel.textContent = `Source: ${outputPricingSourceLabel()}`;
  }
  if (elements.outputTotalLines) {
    elements.outputTotalLines.textContent = `Total lines: ${safeRows.length}`;
  }
}

function renderMatchSummary(result = {}) {
  const rows = result.pricing_matches || [];
  if (!rows.length) {
    elements.matchSummary.innerHTML = "";
    return;
  }
  const stats = matchSummaryStats(rows);
  const totalValue = formatSubtotalValue(stats);
  elements.matchSummary.innerHTML = `
    <div class="stat-card-row output-stat-card-row">
      <div class="stat-card">
        <span class="stat-card-icon blue" aria-hidden="true">S</span>
        <span class="stat-card-value">${stats.sections}</span>
        <span class="stat-card-label">Sections</span>
      </div>
      <div class="stat-card">
        <span class="stat-card-icon green" aria-hidden="true">#</span>
        <span class="stat-card-value">${stats.pricedRows}</span>
        <span class="stat-card-label">Priced rows</span>
      </div>
      <div class="stat-card">
        <span class="stat-card-icon red" aria-hidden="true">!</span>
        <span class="stat-card-value">${stats.needsManualInput}</span>
        <span class="stat-card-label">Needs manual input</span>
      </div>
      <div class="stat-card">
        <span class="stat-card-icon amber" aria-hidden="true">$</span>
        <span class="stat-card-value">${totalValue}</span>
        <span class="stat-card-label">Subtotal</span>
      </div>
    </div>
  `;
}

function renderPricingMatches(rows = [], options = {}) {
  state.pricingMatches = Array.isArray(rows) ? rows : [];
  if (options.fromPricingMatches) {
    state.outputRows = state.pricingMatches.map(outputRowFromPricingMatch);
  } else if (Array.isArray(rows) && rows.length && rows[0]?.price_mode) {
    state.outputRows = rows.map(normalizeOutputRow);
  }
  state.outputRows = sortOutputRows(state.outputRows);
  if (elements.outputSortMode) elements.outputSortMode.value = state.outputSortMode;
  const outputRows = state.outputRows;
  updateOutputHeader(outputRows);
  elements.pricingTableWrap.hidden = !outputRows.length;
  elements.pricingEmptyState.hidden = Boolean(outputRows.length) || Boolean(elements.pricingReviewMessages.innerHTML.trim());
  if (!outputRows.length) {
    elements.pricingMatchesBody.innerHTML = `<tr><td colspan="7">No output rows yet.</td></tr>`;
    updateDownloadButton();
    return;
  }
  elements.pricingMatchesBody.innerHTML = outputRows
    .map((row, index) => `
      <tr data-output-row="${index}">
        ${renderOutputEditCell(row, index, "section")}
        ${renderOutputEditCell(row, index, "description", "output-description-cell")}
        ${renderOutputEditCell(row, index, "quantity")}
        ${renderOutputEditCell(row, index, "unit")}
        ${renderOutputEditCell(row, index, "unit_price_override")}
        <td class="amount-cell ${outputCellDisplayValue(row, "amount") === "???" ? "is-pending" : ""}">${escapeHtml(outputCellDisplayValue(row, "amount"))}</td>
        <td class="output-row-actions">
          <button class="output-delete-button" type="button" data-output-delete-row="${index}" aria-label="Delete row">Del</button>
        </td>
      </tr>
    `)
    .join("");
  updateDownloadButton();
}

function clearPricingReviewMessages() {
  state.pricingIssues = [];
  state.outputErrors = [];
  elements.pricingReviewMessages.innerHTML = "";
  const pricingText = (elements.pricingMatchesBody.textContent || "").trim();
  const hasRows = Boolean(pricingText) && !/No pricing matches yet/i.test(pricingText);
  elements.pricingEmptyState.hidden = hasRows;
}

function handleOutputRowEdit(event) {
  const input = event.target.closest("[data-output-field]");
  if (!input) return;
  const index = Number(input.dataset.outputRow);
  const field = input.dataset.outputField;
  if (!Number.isInteger(index) || index < 0 || !state.outputRows[index] || !field) return;
  state.outputRows[index] = recalculateOutputRow({
    ...state.outputRows[index],
    [field]: input.value,
  });
  state.lineItems = outputRowsToLineItems();
  state.downloadFile = null;
  const validation = outputRowsValid();
  renderOutputValidationMessages(validation.valid ? [] : validation.errors);
  renderPricingMatches(state.outputRows);
  syncControlStates();
}

function commitOutputEditor(editor) {
  if (!editor || !editor.dataset.outputEditorField) return;
  if (editor.isConnected === false) return;
  const index = Number(editor.dataset.outputRow);
  const field = editor.dataset.outputEditorField;
  if (!Number.isInteger(index) || index < 0 || !state.outputRows[index] || !field) return;
  const currentRow = state.outputRows[index];
  let nextRow = { ...currentRow, [field]: editor.value };
  if (field === "unit_price_override") {
    const value = String(editor.value || "").trim();
    if (value.toLowerCase() === "included" || (currentRow.price_mode === "Included" && value === "")) {
      nextRow = { ...currentRow, price_mode: "Included", unit_price_override: "", display_price: "Included" };
    } else {
      nextRow = { ...currentRow, price_mode: "Priced", display_price: "", unit_price_override: value };
    }
  }
  state.outputRows[index] = recalculateOutputRow(nextRow);
  state.lineItems = outputRowsToLineItems();
  state.downloadFile = null;
  const validation = outputRowsValid();
  renderOutputValidationMessages(validation.valid ? [] : validation.errors);
  renderPricingMatches(state.outputRows);
  renderMatchSummary({ pricing_matches: state.outputRows });
  syncControlStates();
}

function applyOutputIncludedAction(button) {
  const index = Number(button?.dataset.outputRow);
  if (!Number.isInteger(index) || index < 0 || !state.outputRows[index]) return;
  state.outputRows[index] = recalculateOutputRow({
    ...state.outputRows[index],
    price_mode: "Included",
    unit_price_override: "",
    display_price: "Included",
  });
  state.lineItems = outputRowsToLineItems();
  state.downloadFile = null;
  const validation = outputRowsValid();
  renderOutputValidationMessages(validation.valid ? [] : validation.errors);
  renderPricingMatches(state.outputRows);
  renderMatchSummary({ pricing_matches: state.outputRows });
  syncControlStates();
}

function openOutputCellEditor(cell) {
  if (!cell || cell.querySelector("[data-output-editor-field]")) return;
  const index = Number(cell.dataset.outputRow);
  const field = cell.dataset.outputEditField;
  if (!Number.isInteger(index) || index < 0 || !state.outputRows[index] || !field) return;
  cell.innerHTML = outputEditorHtml(state.outputRows[index], index, field);
  const editor = cell.querySelector("[data-output-editor-field]");
  if (!editor) return;
  editor.focus();
  if (typeof editor.select === "function") editor.select();
}

function handleOutputCellClick(event) {
  const deleteButton = event.target.closest("[data-output-delete-row]");
  if (deleteButton) {
    event.preventDefault();
    deleteOutputRow(Number(deleteButton.dataset.outputDeleteRow));
    return;
  }
  const includedButton = event.target.closest("[data-output-included-action]");
  if (includedButton) {
    event.preventDefault();
    applyOutputIncludedAction(includedButton);
    return;
  }
  if (event.target.closest("[data-output-editor-field]")) return;
  const cell = event.target.closest(".output-edit-cell");
  if (!cell) return;
  openOutputCellEditor(cell);
}

function handleOutputIncludedPointerDown(event) {
  if (!event.target.closest("[data-output-included-action]")) return;
  event.preventDefault();
}

function handleOutputCellOpen(event) {
  const cell = event.target.closest(".output-edit-cell");
  if (!cell) return;
  openOutputCellEditor(cell);
}

function handleOutputCellKeydown(event) {
  const editor = event.target.closest("[data-output-editor-field]");
  if (editor) {
    if (event.key === "Escape") {
      event.preventDefault();
      renderPricingMatches(state.outputRows);
      return;
    }
    if (event.key === "Enter" && editor.tagName !== "TEXTAREA") {
      event.preventDefault();
      commitOutputEditor(editor);
    }
    return;
  }
  const cell = event.target.closest(".output-edit-cell");
  if (!cell) return;
  if (event.key === "Enter" || event.key === "F2") {
    event.preventDefault();
    openOutputCellEditor(cell);
  }
}

function handleOutputEditorCommit(event) {
  const editor = event.target.closest("[data-output-editor-field]");
  if (!editor) return;
  commitOutputEditor(editor);
}

function renderBasisEmptyState(message = "Load images, complete Customer and Quote Company, then start analysis to review the draft here.") {
  elements.basisReviewSurface.innerHTML = `
    <div class="basis-empty-state">
      <strong>Quotation basis draft</strong>
      <p>${escapeHtml(message)}</p>
    </div>
  `;
}

function renderBasisFailureState(message = GENERIC_FAILURE_MESSAGE) {
  elements.basisReviewSurface.innerHTML = `
    <div class="basis-empty-state basis-empty-state-error">
      <strong>AI analysis did not complete</strong>
      <p>${escapeHtml(message)}</p>
    </div>
  `;
}

function clearBasisReviewSurface() {
  elements.basisReviewSurface.innerHTML = "";
}

function basisSections(sections = state.quoteBasisSections) {
  const normalized = normalizeQuoteBasisSections(sections.length ? sections : state.quoteBasis);
  return normalized.length
    ? normalized
    : [{ id: "draft", title: "Quote Basis", lines: [{ tag: "Confirm", text: "No detail generated yet." }] }];
}

function unresolvedConfirmLines(sections = state.quoteBasisSections) {
  return basisSections(sections)
    .flatMap((section) => (section.lines || [])
      .filter((line) => normalizeBasisTag(line.tag) === "Confirm" || isPendingAiProposalLine(line))
      .map((line) => `${section.title}: ${line.text}`));
}

function basisConfirmBlockReason(sections = state.quoteBasisSections) {
  if ((state.blockingClarificationQuestions || []).some((question) => question.status !== "answered" && !String(question.answer || "").trim())) {
    return "Answer all clarification questions before confirming quotation basis.";
  }
  return unresolvedConfirmLines(sections).length
    ? "Resolve all review lines before confirming quotation basis."
    : "";
}

function basisTagLabel(tag = "") {
  const normalized = normalizeBasisTag(tag);
  if (normalized === "Custom") return "AI Proposal";
  return normalized;
}

function basisQuantityText(value = "") {
  if (value === undefined || value === null) return "";
  const text = String(value).replaceAll(",", "").trim();
  if (!text) return "";
  const match = text.match(/^-?\d+(?:\.\d+)?/);
  if (!match) return "";
  const quantity = Number(match[0]);
  if (!Number.isFinite(quantity) || quantity <= 0) return "";
  return quantity.toLocaleString(undefined, { maximumFractionDigits: 2 });
}

function basisQuantityLabel(line = {}) {
  const quantityText = basisQuantityText(line.quantity);
  if (!quantityText) return "";
  const unit = normalizeUnit(line.unit || "");
  return `${quantityText}${unit ? ` ${unit}` : ""}`;
}

function basisQuantityDisplayLabel(line = {}) {
  const quantityText = basisQuantityText(line.quantity);
  if (!quantityText) return "1 lot";
  let unit = normalizeUnit(line.unit || "").replace(/\s+length$/i, "").trim();
  if (/^nos?\.?$/i.test(unit)) unit = "nos.";
  return `${quantityText}${unit ? ` ${unit}` : ""}`;
}

function basisChatLineContext(line = {}) {
  const tag = normalizeBasisTag(line.tag);
  const text = String(line.text || "").trim();
  return `${tag}: ${text}`.trim();
}

function hasPricingReferenceDescription(line = {}) {
  return Boolean(pricingReferenceLineText(line.pricing_reference_description || ""));
}

function basisLinePillLabel(line = {}) {
  const tag = normalizeBasisTag(line.tag);
  if (isPendingAiProposalLine(line)) return "AI Confirm";
  return basisTagLabel(tag);
}

function basisLineAcceptsAsAiProposal(line = {}) {
  return isCustomPricingBasisLine(line);
}

function basisConfidenceLabel(line = {}) {
  const confidence = normalizeConfidence(line.confidence ?? line.confidence_pct);
  return `${confidence === null ? 50 : confidence}%`;
}

function basisTotalLineCount(sections = []) {
  return (Array.isArray(sections) ? sections : []).reduce((total, section) => {
    const lines = Array.isArray(section?.lines) ? section.lines : [];
    return total + lines.length;
  }, 0);
}

function basisTotalLineLabel(count = 0) {
  return `Total lines: ${Math.max(0, Number(count) || 0)}`;
}

function renderBasisConfirmSummary(sections = state.quoteBasisSections) {
  return "";
}

function basisCatalogReferenceTitle(line = {}) {
  return "";
}

function basisLineTitle(line = {}) {
  return basisCatalogReferenceTitle(line);
}

function basisPillTitle(line = {}, tag = "") {
  return basisCatalogReferenceTitle(line);
}

function renderBasisLine(section, line, index) {
  const tag = normalizeBasisTag(line.tag);
  const customPricing = isCustomPricingBasisLine(line);
  const pendingAiProposal = isPendingAiProposalLine(line);
  const pillLabel = basisLinePillLabel(line);
  const acceptsAsAiProposal = basisLineAcceptsAsAiProposal(line);
  const quantityLabel = basisQuantityDisplayLabel(line);
  const confidenceLabel = basisConfidenceLabel(line);
  const pillTitle = basisPillTitle(line, tag);
  const lineTitle = basisLineTitle(line);
  const pillTitleAttribute = pillTitle ? ` title="${escapeHtml(pillTitle)}"` : "";
  const lineTitleAttribute = lineTitle ? ` title="${escapeHtml(lineTitle)}"` : "";
  const rowClasses = [
    "basis-line-row",
    `basis-line-${tag.toLowerCase()}`,
    customPricing ? "basis-line-custom-priced" : "",
    pendingAiProposal ? "basis-line-custom-confirm" : "",
    pillLabel === "AI Confirm" ? "basis-line-ai-confirm" : "",
  ].filter(Boolean).join(" ");
  const primaryAction = acceptsAsAiProposal
    ? `<button class="basis-line-tag-button" type="button" data-basis-section="${escapeHtml(section.id)}" data-basis-line-index="${index}" data-basis-tag="Custom" aria-label="Accept this AI proposal" title="Accept AI proposal">&#x2713;</button>`
    : `<button class="basis-line-tag-button" type="button" data-basis-section="${escapeHtml(section.id)}" data-basis-line-index="${index}" data-basis-tag="Include" aria-label="Mark this line as included" title="Mark included">&#x2713;</button>`;
  return `
    <li class="${escapeHtml(rowClasses)}">
      <span class="basis-line-icon" aria-hidden="true"></span>
      <span class="basis-line-meta">
        <span class="basis-line-pill"${pillTitleAttribute}>${escapeHtml(pillLabel)}</span>
        <span class="basis-confidence-pill" title="AI confidence">${escapeHtml(confidenceLabel)}</span>
        <span class="basis-quantity-text">${escapeHtml(quantityLabel)}</span>
      </span>
      <span class="basis-line-text"${lineTitleAttribute}>${escapeHtml(line.text)}</span>
      <span class="basis-line-actions">
        ${primaryAction}
        <button class="basis-line-tag-button" type="button" data-basis-section="${escapeHtml(section.id)}" data-basis-line-index="${index}" data-basis-tag="Exclude" aria-label="Mark this line as excluded" title="Mark excluded">X</button>
        <button class="basis-line-tool" type="button" data-revise-section="${escapeHtml(section.id)}" data-revise-line-index="${index}" aria-label="Revise this line" title="Revise this line">Re</button>
      </span>
    </li>
  `;
}

function renderBasisTagLegend() {
  return `
    <div class="basis-tag-legend" aria-label="Quote basis tag meanings">
      <div class="basis-tag-legend-column basis-tag-legend-actions">
        ${BASIS_TAGS.filter(([tag]) => ["Include", "Exclude", "Custom"].includes(tag)).map(([tag, label, description]) => `
          <span class="basis-tag-legend-item basis-line-${escapeHtml(tag.toLowerCase())}">
            <strong class="basis-line-pill">${escapeHtml(label)}</strong>
            <span>${escapeHtml(description)}</span>
          </span>
        `).join("")}
      </div>
      <div class="basis-tag-legend-column basis-tag-legend-review">
        <span class="basis-tag-legend-item basis-line-confirm">
          <strong class="basis-line-pill">Confirm</strong>
          <span>Needs include, exclude, or revision</span>
        </span>
        <span class="basis-tag-legend-item basis-line-custom-confirm">
          <strong class="basis-line-pill">AI Confirm</strong>
          <span>AI proposal needs acceptance or revision</span>
        </span>
        <span class="basis-tag-legend-item basis-line-confirm">
          <strong class="basis-confidence-pill">92%</strong>
          <span>AI confidence level in quote basis line</span>
        </span>
      </div>
    </div>
  `;
}

function renderQuoteBasisMessage(basis = state.quoteBasis, source = "") {
  const aiFailed = source === "local" && state.aiFailed;
  if (aiFailed) {
    return `
      <div class="basis-empty-state basis-empty-state-error">
        <strong>AI analysis did not complete</strong>
        <p>${GENERIC_FAILURE_MESSAGE}</p>
      </div>
    `;
  }
  const statusText = aiFailed ? "AI failed" : source === "edited" ? "Edited draft" : "Needs review";
  const sections = basisSections(state.quoteBasisSections.length ? state.quoteBasisSections : normalizeQuoteBasisSections(basis));
  return `
    <div class="assistant-card quote-basis-card ${aiFailed ? "quote-basis-card-failed" : ""}">
      <div class="quote-basis-header">
        <div>
          <div class="quote-basis-title-row">
            <h3>Quote Basis</h3>
            <span>${escapeHtml(statusText)}</span>
          </div>
          <p>${aiFailed ? GENERIC_FAILURE_MESSAGE : "Please review the AI takeoff and revise individual lines where needed."}</p>
        </div>
        <div class="quote-basis-source">
          <span>${aiFailed ? "Source: Local fallback only" : `Source: ${escapeHtml(outputPricingSourceLabel())}`}</span>
          <strong>${escapeHtml(basisTotalLineLabel(basisTotalLineCount(sections)))}</strong>
        </div>
      </div>
      ${renderAnalysisFindings()}
      ${renderBasisConfirmSummary(sections)}
      ${renderBasisTagLegend()}
      <div class="basis-review-grid">
        ${sections.map((section) => `
          <div class="basis-review-item">
            <div class="basis-section-heading">
              <h4>${escapeHtml(section.title)}</h4>
              <span class="basis-section-actions">
                <button class="basis-section-action-button" type="button" data-basis-section="${escapeHtml(section.id)}" data-basis-section-action="Include" aria-label="Accept all review lines in ${escapeHtml(section.title)}" title="Accept all review lines">&#x2713;</button>
                <button class="basis-section-action-button" type="button" data-basis-section="${escapeHtml(section.id)}" data-basis-section-action="Exclude" aria-label="Exclude all review lines in ${escapeHtml(section.title)}" title="Exclude all review lines">X</button>
                <span class="basis-section-action-spacer" aria-hidden="true"></span>
              </span>
            </div>
            <ul class="basis-line-list">
              ${(section.lines || []).map((line, index) => renderBasisLine(section, line, index)).join("")}
            </ul>
          </div>
        `).join("")}
      </div>
    </div>
  `;
}

function basisFieldLabel(field) {
  return basisDisplayTitle(state.quoteBasisSections.find((section) => section.id === field)?.title)
    || BASIS_FIELDS.find(([key]) => key === field)?.[1]
    || "Quote basis";
}

function updateQuoteBasisCard(source = "edited") {
  elements.basisReviewSurface.innerHTML = renderQuoteBasisMessage(state.quoteBasis, source);
}

function retagBasisLine(sectionId, lineIndex, nextTag) {
  const requestedTag = normalizeBasisTag(nextTag);
  if (!sectionId || !["Include", "Custom", "Exclude"].includes(requestedTag)) return;
  const sections = cloneQuoteBasisSections(state.quoteBasisSections);
  const section = sections.find((item) => item.id === sectionId);
  const index = Number(lineIndex);
  if (!section || !section.lines[index]) return;
  const currentLine = section.lines[index];
  const resolvedTag = requestedTag === "Include" && basisLineAcceptsAsAiProposal(currentLine) ? "Custom" : requestedTag;
  const customPricing = isCustomPricingBasisLine(currentLine) || resolvedTag === "Custom";
  const customConfirmed = customPricing && ["Include", "Custom"].includes(requestedTag);
  section.lines[index] = {
    ...currentLine,
    tag: resolvedTag,
    ...(customPricing ? { custom_pricing: true } : {}),
    ...(customConfirmed ? { custom_confirmed: true } : { custom_confirmed: false }),
  };
  state.quoteBasisSections = sections;
  state.quoteBasis = quoteBasisFromSections(sections);
  state.basisConfirmed = false;
  state.downloadFile = null;
  updateQuoteBasisCard("edited");
  syncControlStates();
}

function retagBasisSectionConfirmLines(sectionId, nextTag) {
  if (!sectionId || !["Include", "Exclude"].includes(nextTag)) return;
  const sections = cloneQuoteBasisSections(state.quoteBasisSections);
  const section = sections.find((item) => item.id === sectionId);
  if (!section) return;
  let changed = false;
  section.lines = (section.lines || []).map((line) => {
    const resolvedTag = nextTag === "Include" && basisLineAcceptsAsAiProposal(line) ? "Custom" : nextTag;
    const customPricing = isCustomPricingBasisLine(line) || resolvedTag === "Custom";
    const customConfirmed = customPricing && nextTag === "Include";
    const currentTag = normalizeBasisTag(line.tag);
    if (currentTag === resolvedTag && (!customPricing || Boolean(line.custom_confirmed) === customConfirmed)) return line;
    changed = true;
    return {
      ...line,
      tag: resolvedTag,
      ...(customPricing ? { custom_pricing: true, custom_confirmed: customConfirmed } : { custom_confirmed: false }),
    };
  });
  if (!changed) return;
  state.quoteBasisSections = sections;
  state.quoteBasis = quoteBasisFromSections(sections);
  state.basisConfirmed = false;
  state.downloadFile = null;
  updateQuoteBasisCard("edited");
  syncControlStates();
}

function basisChatIntroMessage() {
  return "Tell me what to change. I will draft the full replacement sentence for approval before applying it.";
}

function selectedBasisLine() {
  const section = state.quoteBasisSections.find((item) => item.id === state.basisChat.sectionId);
  return section?.lines?.[state.basisChat.lineIndex] || null;
}

function appendBasisChatMessage(role, text, options = {}) {
  const message = document.createElement("div");
  message.className = `basis-chat-message ${role}`;
  message.innerHTML = `
    <span class="basis-chat-label">${role === "user" ? "You" : "Assistant"}</span>
    ${renderPlainText(text)}
    ${options.proposalActions ? `
      <div class="basis-chat-message-actions">
        <button class="secondary-button" type="button" data-basis-chat-action="discard">Discard</button>
        <button class="primary-button" type="button" data-basis-chat-action="apply">Apply</button>
      </div>
    ` : ""}
  `;
  elements.basisChatMessages.appendChild(message);
  elements.basisChatMessages.scrollTop = elements.basisChatMessages.scrollHeight;
}

function appendBasisChatTyping() {
  const message = document.createElement("div");
  message.className = "basis-chat-message assistant is-typing";
  message.dataset.basisChatTyping = "true";
  message.innerHTML = `
    <span class="basis-chat-label">Assistant</span>
    <span class="basis-chat-typing-dots" role="status" aria-label="Assistant is typing">
      <i></i><i></i><i></i>
    </span>
  `;
  elements.basisChatMessages.appendChild(message);
  elements.basisChatMessages.scrollTop = elements.basisChatMessages.scrollHeight;
  return message;
}

function removeBasisChatTyping(message) {
  if (message?.parentNode) message.remove();
}

function resetBasisChatProposal() {
  state.basisChat.proposal = null;
  elements.basisChatProposal.hidden = true;
  elements.basisChatProposal.innerHTML = "";
  elements.basisChatProposalActions.hidden = true;
  elements.basisChatApplyButton.disabled = true;
  elements.basisChatKeepButton.disabled = true;
}

function proposalChangedFields(proposal) {
  const nextSections = normalizeQuoteBasisSections(proposal?.quoteBasisSections || proposal?.quoteBasis || {});
  const currentSections = cloneQuoteBasisSections(state.quoteBasisSections);
  const currentById = new Map(currentSections.map((section) => [section.id, section]));
  return nextSections
    .filter((section) => JSON.stringify(currentById.get(section.id) || null) !== JSON.stringify(section))
    .map((section) => section.title);
}

function proposalLineDelta(proposal) {
  if (state.basisChat.scope !== "line" || !state.basisChat.sectionId) return null;
  const nextSections = normalizeQuoteBasisSections(proposal?.quoteBasisSections || proposal?.quoteBasis || {});
  const section = nextSections.find((item) => item.id === state.basisChat.sectionId);
  const nextLine = section?.lines?.[state.basisChat.lineIndex];
  const currentLine = selectedBasisLine() || parseBasisLine(state.basisChat.line);
  if (!nextLine || !currentLine || nextLine.text === currentLine.text && nextLine.tag === currentLine.tag) return null;
  return {
    currentTag: normalizeBasisTag(currentLine.tag),
    currentConfidence: normalizeConfidence(currentLine.confidence ?? currentLine.confidence_pct),
    currentPricingReferenceDescription: currentLine.pricing_reference_description || "",
    currentText: currentLine.text,
    nextTag: normalizeBasisTag(nextLine.tag),
    nextConfidence: normalizeConfidence(nextLine.confidence ?? nextLine.confidence_pct),
    nextPricingReferenceDescription: nextLine.pricing_reference_description || "",
    nextText: nextLine.text,
  };
}

function renderProposalSectionChips(changedFields = []) {
  if (!changedFields.length) return "";
  const visible = changedFields.slice(0, 6);
  const remaining = changedFields.length - visible.length;
  return `
    <div class="basis-chat-proposal-meta">
      <span>Affected</span>
      <div class="basis-chat-proposal-chips">
        ${visible.map((field) => `<i>${escapeHtml(field)}</i>`).join("")}
        ${remaining > 0 ? `<i>+${remaining} more</i>` : ""}
      </div>
    </div>
  `;
}

function renderBasisChatProposalCard(proposal, changedFields = []) {
  const delta = proposalLineDelta(proposal);
  const message = proposal?.message || "Review this proposed line change before applying it.";
  return `
    <div class="basis-chat-proposal-header">
      <div>
        <span>Proposed change</span>
        <strong>Basis line update</strong>
      </div>
      <span class="basis-chat-proposal-status">Review</span>
    </div>
    <p class="basis-chat-proposal-summary">${escapeHtml(message)}</p>
    ${proposal.literalReplacement ? renderLiteralReplacementPreview(proposal) : delta ? `
      <div class="basis-chat-compare">
        <div class="basis-chat-compare-card">
          <span>Current</span>
          <p><strong>${escapeHtml(basisLinePillLabel({ tag: delta.currentTag, confidence: delta.currentConfidence, pricing_reference_description: delta.currentPricingReferenceDescription }))}</strong> ${escapeHtml(delta.currentText)}</p>
        </div>
        <div class="basis-chat-compare-card is-proposed">
          <span>Proposed</span>
          <p><strong>${escapeHtml(basisLinePillLabel({ tag: delta.nextTag, confidence: delta.nextConfidence, pricing_reference_description: delta.nextPricingReferenceDescription }))}</strong> ${escapeHtml(delta.nextText)}</p>
        </div>
      </div>
    ` : `
      <p class="basis-chat-proposal-summary">This proposal updates the selected basis line and keeps existing line items available for review.</p>
    `}
    ${renderProposalSectionChips(changedFields)}
  `;
}

function setBasisChatProposal(proposal) {
  state.basisChat.proposal = proposal;
  const changedFields = proposalChangedFields(proposal);
  elements.basisChatProposal.hidden = false;
  elements.basisChatProposal.innerHTML = renderBasisChatProposalCard(proposal, changedFields);
  elements.basisChatProposalActions.hidden = false;
  elements.basisChatApplyButton.disabled = false;
  elements.basisChatKeepButton.disabled = false;
  appendBasisChatMessage("assistant", "I drafted a proposed update below. Review it, then apply or discard.");
}

function basisChatFriendlyError(messages = []) {
  if (messages && typeof messages === "object" && !Array.isArray(messages)) {
    return genericFailureMessage(messages);
  }
  return genericFailureMessage();
}

function setBasisChatBusy(isBusy) {
  elements.basisChatPrompt.disabled = isBusy;
  elements.basisChatSendButton.disabled = isBusy;
  elements.basisChatApplyButton.disabled = isBusy || !state.basisChat.proposal;
  elements.basisChatKeepButton.disabled = isBusy || !state.basisChat.proposal;
}

function openBasisChatOverlay(scope = "line", options = {}) {
  if (scope !== "line") return;
  state.basisChat = {
    scope,
    field: options.sectionId || options.field || "",
    sectionId: options.sectionId || options.field || "",
    lineIndex: Number.isInteger(options.lineIndex) ? options.lineIndex : -1,
    line: options.line || "",
    quantity: options.quantity ?? "",
    unit: options.unit || "",
    quantityLabel: options.quantityLabel || "",
    proposal: null,
  };
  resetBasisChatProposal();
  elements.basisChatTitle.textContent = "Revise basis line";
  elements.basisChatContext.classList.toggle("has-selected-line", scope === "line");
  const line = selectedBasisLine() || parseBasisLine(state.basisChat.line);
  const tag = normalizeBasisTag(line.tag);
  const tagClass = `basis-chat-selected-tag-${tag.toLowerCase()}`;
  const tagLineClass = `basis-chat-selected-line-${tag.toLowerCase()}`;
  const quantityLabel = basisQuantityLabel(line);
  const displayQuantityLabel = basisQuantityDisplayLabel(line);
  const confidenceLabel = basisConfidenceLabel(line);
  const selectedTagTitle = basisPillTitle(line, tag);
  const selectedTagTitleAttribute = selectedTagTitle ? ` title="${escapeHtml(selectedTagTitle)}"` : "";
  state.basisChat.line = basisChatLineContext(line);
  state.basisChat.quantity = line.quantity ?? state.basisChat.quantity ?? "";
  state.basisChat.unit = normalizeUnit(line.unit || state.basisChat.unit || "");
  state.basisChat.quantityLabel = quantityLabel || state.basisChat.quantityLabel || "";
  elements.basisChatContext.innerHTML = `
    <span class="basis-chat-context-label">${escapeHtml(basisFieldLabel(state.basisChat.sectionId))}</span>
    <span class="basis-chat-selected-line ${escapeHtml(tagLineClass)}">
      <span class="basis-chat-selected-meta">
        <strong class="basis-chat-selected-tag ${escapeHtml(tagClass)}"${selectedTagTitleAttribute}>${escapeHtml(basisLinePillLabel(line))}</strong>
        <span class="basis-confidence-pill" title="AI confidence">${escapeHtml(confidenceLabel)}</span>
      </span>
      <span class="basis-chat-selected-body">
        <span class="basis-chat-selected-quantity">${escapeHtml(displayQuantityLabel)}</span>
        <span class="basis-chat-selected-text">${escapeHtml(line.text || state.basisChat.line)}</span>
      </span>
    </span>
  `;
  elements.basisChatPrompt.placeholder = "Describe the change for this line...";
  elements.basisChatMessages.innerHTML = "";
  appendBasisChatMessage("assistant", basisChatIntroMessage());
  elements.basisChatPrompt.value = "";
  elements.basisChatOverlay.hidden = false;
  elements.basisChatOverlay.classList.add("is-open");
  document.body.classList.add("basis-chat-open");
  window.setTimeout(() => elements.basisChatPrompt.focus(), 0);
}

function closeBasisChatOverlay() {
  elements.basisChatOverlay.classList.remove("is-open");
  elements.basisChatOverlay.hidden = true;
  document.body.classList.remove("basis-chat-open");
  resetBasisChatProposal();
}

function basisChatPayload(text) {
  return {
    ...buildPayload(),
    basis_chat: {
      question: text,
      scope: state.basisChat.scope,
      field: state.basisChat.sectionId || state.basisChat.field,
      line_index: state.basisChat.lineIndex,
      line: state.basisChat.line,
      quantity: state.basisChat.quantity,
      unit: state.basisChat.unit,
      quantity_label: state.basisChat.quantityLabel,
    },
  };
}

function normalizeServerBasisChatProposal(proposal = {}) {
  const quoteBasis = proposal.quoteBasis || proposal.quote_basis || {};
  const sections = normalizeQuoteBasisSections(proposal.quoteBasisSections || proposal.quote_basis_sections || quoteBasis);
  return {
    message: String(proposal.message || "AI drafted a proposed quote basis update.").trim(),
    quoteBasis: { ...cloneQuoteBasis(quoteBasis), ...quoteBasisFromSections(sections) },
    quoteBasisSections: sections,
    lineItems: Array.isArray(proposal.lineItems || proposal.line_items)
      ? (proposal.lineItems || proposal.line_items).map(normalizeLineItem)
      : state.lineItems.map(normalizeLineItem),
  };
}

function parseLiteralReplacementCommand(text = "") {
  const match = String(text || "").trim().match(/^(?:change|replace|switch)(?:\s+all)?\s+(.+?)\s+(?:to|with)\s+(.+?)\s*$/i);
  if (!match) return null;
  const from = match[1].trim().replace(/^['"]|['"]$/g, "");
  const to = match[2].trim().replace(/^['"]|['"]$/g, "");
  if (!from || !to || from.length > 80 || to.length > 80) return null;
  return { from, to };
}

function replaceLiteralText(value = "", from = "", to = "") {
  if (!from) return { text: String(value || ""), changed: false };
  const pattern = new RegExp(from.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), "gi");
  const original = String(value || "");
  const next = original.replace(pattern, to);
  return { text: next, changed: next !== original };
}

function buildLiteralReplacementProposal(command) {
  const sections = cloneQuoteBasisSections(state.quoteBasisSections);
  let changedLineCount = 0;
  const affectedSectionIds = new Set();
  const snippets = [];
  sections.forEach((section) => {
    (section.lines || []).forEach((line) => {
      const replaced = replaceLiteralText(line.text, command.from, command.to);
      if (!replaced.changed) return;
      snippets.push({ section: section.title, before: line.text, after: replaced.text });
      line.text = replaced.text;
      line.tag = "Confirm";
      if (isCustomPricingBasisLine(line)) line.custom_pricing = true;
      changedLineCount += 1;
      affectedSectionIds.add(section.id);
    });
  });
  const legacyBasis = quoteBasisFromSections(sections);
  Object.keys(legacyBasis).forEach((key) => {
    legacyBasis[key] = replaceLiteralText(legacyBasis[key], command.from, command.to).text;
  });
  const lineItems = state.lineItems.map((item) => {
    const replaced = replaceLiteralText(item.description, command.from, command.to);
    return replaced.changed ? { ...item, description: replaced.text } : item;
  });
  const outputRows = state.outputRows.map((row) => {
    const replaced = replaceLiteralText(row.description, command.from, command.to);
    return replaced.changed ? recalculateOutputRow({ ...row, description: replaced.text }) : row;
  });
  const changedOutputRows = outputRows.filter((row, index) => row.description !== state.outputRows[index]?.description).length;
  if (!changedLineCount && !changedOutputRows && !lineItems.some((item, index) => item.description !== state.lineItems[index]?.description)) return null;
  return {
    message: `Literal replacement: changed ${changedLineCount} basis line${changedLineCount === 1 ? "" : "s"} across ${affectedSectionIds.size} section${affectedSectionIds.size === 1 ? "" : "s"}.`,
    literalReplacement: true,
    changedLineCount,
    affectedSectionCount: affectedSectionIds.size,
    snippets: snippets.slice(0, 5),
    quoteBasis: legacyBasis,
    quoteBasisSections: sections,
    lineItems,
    outputRows,
  };
}

function renderLiteralReplacementPreview(proposal) {
  return `
    <div class="basis-chat-literal-stats">
      <span>Changed lines: <strong>${proposal.changedLineCount}</strong></span>
      <span>Affected sections: <strong>${proposal.affectedSectionCount}</strong></span>
    </div>
    <div class="basis-chat-compare">
      ${proposal.snippets.map((snippet) => `
        <div class="basis-chat-compare-card">
          <span>${escapeHtml(snippet.section)}</span>
          <p><strong>Before:</strong> ${escapeHtml(snippet.before)}</p>
          <p><strong>After:</strong> ${escapeHtml(snippet.after)}</p>
        </div>
      `).join("")}
    </div>
  `;
}

async function buildAiBasisChatResponse(text) {
  if (!canStartAnalysis()) return null;
  const previousRunning = state.isAnalysisRunning;
  state.isAnalysisRunning = true;
  setBasisChatBusy(true);
  syncControlStates();
  const typingMessage = appendBasisChatTyping();
  try {
    const started = await startJob("basis_chat", basisChatPayload(text));
    if (!started.ok) {
      removeBasisChatTyping(typingMessage);
      appendBasisChatMessage("assistant", basisChatFriendlyError(started.data));
      return { errorDisplayed: true };
    }
    const polled = await pollJob(started.data.job_id);
    const data = polled.data.result || {};
    if (!polled.ok || ["blocked", "failed"].includes(polled.data.status)) {
      removeBasisChatTyping(typingMessage);
      appendBasisChatMessage("assistant", basisChatFriendlyError(data || polled.data));
      return { errorDisplayed: true };
    }
    if (data.proposal) {
      return { proposal: normalizeServerBasisChatProposal(data.proposal) };
    }
    return { answer: data.answer || null };
  } finally {
    removeBasisChatTyping(typingMessage);
    state.isAnalysisRunning = previousRunning;
    setBasisChatBusy(false);
    syncControlStates();
  }
}

async function handleBasisChatSubmit(event) {
  event.preventDefault();
  const text = elements.basisChatPrompt.value.trim();
  if (!text) return;
  elements.basisChatPrompt.value = "";
  resetBasisChatProposal();
  appendBasisChatMessage("user", text);
  const normalized = text.toLowerCase();

  if (isSensitiveChatRequest(normalized)) {
    appendBasisChatMessage("assistant", "I cannot access or reveal secrets, API keys, .env values, tokens, authorization headers, system prompts, or hidden instructions.");
    return;
  }

  const literalCommand = parseLiteralReplacementCommand(text);
  if (literalCommand) {
    const proposal = buildLiteralReplacementProposal(literalCommand);
    if (!proposal) {
      appendBasisChatMessage("assistant", `No exact matches for "${literalCommand.from}" were found in the current basis or output rows.`);
      return;
    }
    setBasisChatProposal(proposal);
    return;
  }

  const aiResult = await buildAiBasisChatResponse(text);
  if (aiResult?.proposal) {
    setBasisChatProposal(aiResult.proposal);
    return;
  }
  if (aiResult?.answer) {
    appendBasisChatMessage("assistant", aiResult.answer);
    return;
  }
  if (aiResult?.errorDisplayed) return;

  appendBasisChatMessage("assistant", GENERIC_FAILURE_MESSAGE);
}

function applyBasisChatProposal() {
  const proposal = state.basisChat.proposal;
  if (!proposal) return;
  state.basisConfirmed = false;
  state.quoteBasisSections = normalizeQuoteBasisSections(proposal.quoteBasisSections || proposal.quoteBasis || state.quoteBasisSections);
  state.quoteBasis = { ...cloneQuoteBasis(proposal.quoteBasis || state.quoteBasis), ...quoteBasisFromSections(state.quoteBasisSections) };
  state.lineItems = Array.isArray(proposal.lineItems) ? proposal.lineItems.map(normalizeLineItem) : [];
  state.outputRows = Array.isArray(proposal.outputRows) ? proposal.outputRows.map(normalizeOutputRow) : [];
  state.originalOutputRows = [];
  state.outputErrors = [];
  setDownloadFiles([]);
  updateQuoteBasisCard("edited");
  setSidePanel("basis", { force: true });
  resetBasisChatProposal();
  closeBasisChatOverlay();
  syncControlStates();
}

function keepCurrentBasis() {
  resetBasisChatProposal();
  appendBasisChatMessage("assistant", "Kept the current quote basis unchanged.");
}

function missingCustomerFields() {
  syncRichTextSources();
  const missing = [];
  const hasLines = (value) => splitLines(value).length > 0;
  if (!String(state.pricingReferenceId || "").trim() || !currentPricingReference()) missing.push("Quote Pricing Reference");
  if (!elements.clientName.value.trim()) missing.push("Client name");
  if (!elements.clientAttention.value.trim()) missing.push("Attention person");
  if (!elements.clientTitle.value.trim()) missing.push("Attention title");
  if (!hasLines(elements.clientAddress.value)) missing.push("Client address");
  if (!elements.projectTitle.value.trim()) missing.push("Quotation Title");
  if (!elements.quoteDate.value.trim()) missing.push("Quote date");
  if (!elements.projectNumber.value.trim()) missing.push("Project number");
  return missing;
}

function missingQuoteCompanyFields() {
  syncRichTextSources();
  const missing = [];
  const hasLines = (value) => splitLines(value).length > 0;
  if (!state.headerLogo?.data_url) missing.push("Header logo");
  if (!hasLines(elements.headerDetails.value)) missing.push("Header details");
  if (!elements.quoteCompanyName.value.trim()) missing.push("Quotation Company");
  if (!elements.acceptanceText.value.trim()) missing.push("Acceptance text");
  if (!elements.konceptSignatory.value.trim()) missing.push("Company signatory");
  if (!elements.konceptTitle.value.trim()) missing.push("Signatory title");
  if (!elements.konceptDateLabel.value.trim()) missing.push("Company date label");
  if (!elements.personLabel.value.trim()) missing.push("Person label");
  if (!elements.stampLabel.value.trim()) missing.push("Stamp label");
  if (!elements.dateLabel.value.trim()) missing.push("Date label");
  return missing;
}

function missingDetailFields() {
  return [...missingCustomerFields(), ...missingQuoteCompanyFields()];
}

function customerDetailsBlockReason(prefix = "Complete Customer details") {
  const missing = missingCustomerFields();
  return missing.length ? `${prefix}: ${missing.join(", ")}.` : "";
}

function quoteCompanyDetailsBlockReason(prefix = "Complete Quote Company details") {
  const missing = missingQuoteCompanyFields();
  return missing.length ? `${prefix}: ${missing.join(", ")}.` : "";
}

function firstMissingDetailsPanel() {
  return missingCustomerFields().length ? "customer" : "quote_company";
}

function startAnalysisBlockReason() {
  if (state.isAnalysisRunning) return "Analysis is already running.";
  if (state.isGenerating) return "Quotation generation is already running.";
  if (!state.images.length) return "Add at least one reference file before starting analysis.";
  return customerDetailsBlockReason("Complete Customer details before starting analysis")
    || quoteCompanyDetailsBlockReason("Complete Quote Company details before starting analysis");
}

function canStartAnalysis() {
  return !startAnalysisBlockReason();
}

function hasSubmittedQuoteBasis() {
  if (state.aiFailed) return false;
  const hasClarificationQuestions = (state.blockingClarificationQuestions || []).length > 0;
  return ["basis_review", "generating", "pricing_review", "completed"].includes(state.workflowStage)
    && (hasClarificationQuestions || state.lineItems.length > 0 || state.quoteBasisSections.some((section) => (section.lines || []).length > 0) || Object.values(state.quoteBasis).some((value) => splitLines(value).length > 0));
}

function hasCompletedQuoteBasis() {
  return state.basisConfirmed || ["pricing_review", "completed"].includes(state.workflowStage);
}

function applyDraftBasis(basis = {}) {
  state.basisConfirmed = false;
  const sections = confirmOnlyQuoteBasisSections(basis);
  state.quoteBasisSections = sections;
  const legacyBasis = basis.quote_basis && typeof basis.quote_basis === "object" ? basis.quote_basis : basis;
  state.quoteBasis = { ...cloneQuoteBasis(legacyBasis), ...quoteBasisFromSections(sections) };
}

function applyDraftLineItems(lineItems = []) {
  state.basisConfirmed = false;
  state.lineItems = lineItems.map(normalizeLineItem);
  state.outputRows = [];
  state.originalOutputRows = [];
  state.outputErrors = [];
}

async function refreshLineItemsFromServer() {
  if (!state.lineItems.length) return true;
  const { ok, data } = await postJson("/api/line-items/normalize", buildLineItemNormalizePayload());
  if (!ok || !Array.isArray(data.line_items)) return false;
  state.lineItems = data.line_items.map(normalizeLineItem);
  state.outputRows = [];
  state.outputErrors = [];
  return true;
}

function captureOriginalAnalysisSnapshot(data = {}) {
  const sections = normalizeQuoteBasisSections(data.quote_basis_sections || data.quote_basis || state.quoteBasisSections);
  state.originalAnalysisSnapshot = {
    quote_basis_sections: cloneQuoteBasisSections(sections),
    quote_basis: { ...state.quoteBasis, ...quoteBasisFromSections(sections) },
    line_items: state.lineItems.map(normalizeLineItem),
    boothDimensions: { ...state.boothDimensions },
    source: data.source || state.draftSource || "",
    analysis_mode: normalizeAnalysisMode(data.analysis_mode || state.lastAnalysisMode),
    warnings: Array.isArray(data.warnings) ? [...data.warnings] : [],
  };
}

function openBlockingClarifications(questions = [], findings = [], project = {}) {
  state.analysisFindings = Array.isArray(findings) ? findings : [];
  state.blockingClarificationQuestions = Array.isArray(questions) ? questions : [];
  state.quoteBasis = { ...EMPTY_BASIS };
  state.quoteBasisSections = [];
  state.lineItems = [];
  state.outputRows = [];
  state.originalOutputRows = [];
  state.basisConfirmed = false;
  state.boothDimensions = normalizeBoothDimensions(project || state.boothDimensions);
  setDownloadFiles([]);
  setWorkflowStage("basis_review");
  renderClarificationQuestions();
  setSidePanel("basis", { force: true });
  syncControlStates();
}

function renderAnalysisFindings(findings = state.analysisFindings || []) {
  const visibleFindings = (Array.isArray(findings) ? findings : [])
    .filter((finding) => String(finding?.text || "").trim());
  if (!visibleFindings.length) return "";
  return `
    <div class="analysis-findings-card">
      <h4>Analysis Findings</h4>
      <ul>${visibleFindings.map((finding) => `<li>${escapeHtml(finding.text)}${finding.confidence_pct ? ` (${finding.confidence_pct}%)` : ""}</li>`).join("")}</ul>
    </div>
  `;
}

function renderClarificationQuestionText(text = "") {
  const normalized = String(text || "").replace(/\s+/g, " ").trim();
  const colonIndex = normalized.indexOf(":");
  const hasCompactLead = colonIndex > 0 && colonIndex <= 72;
  const title = hasCompactLead ? normalized.slice(0, colonIndex).trim() : "";
  const body = hasCompactLead ? normalized.slice(colonIndex + 1).trim() : normalized;
  const points = body
    .split(/(?:;\s+|\.\s+)/)
    .map((point) => point.replace(/\.$/, "").trim())
    .filter(Boolean);
  const visiblePoints = points.length ? points : [body || normalized];
  return `
    ${title ? `<strong class="clarification-question-title">${escapeHtml(title)}</strong>` : ""}
    <ul class="clarification-question-points">
      ${visiblePoints.map((point) => `<li>${escapeHtml(point)}</li>`).join("")}
    </ul>
  `;
}

function renderClarificationQuestions() {
  const questions = state.blockingClarificationQuestions || [];
  elements.basisReviewSurface.innerHTML = `
    <div class="assistant-card quote-basis-card clarification-card">
      <div class="quote-basis-header">
        <div>
          <div class="quote-basis-title-row"><h3>Clarification Questions</h3><span>Required before final Quote Basis</span></div>
          <p>These answers can affect wording, quantity, inclusion, material, or pricing. Final Quote Basis and Output stay locked until all are answered.</p>
        </div>
      </div>
      <form id="clarificationForm" class="clarification-form">
        ${questions.map((question, index) => `
          <label class="clarification-question">
            <div class="clarification-question-copy">${renderClarificationQuestionText(question.question)}</div>
            ${question.reason ? `<small>${escapeHtml(question.reason)}</small>` : ""}
            ${Array.isArray(question.choices) && question.choices.length ? `
              <select data-clarification-index="${index}">
                <option value="">Choose...</option>
                ${question.choices.map((choice) => `<option value="${escapeHtml(choice)}" ${question.answer === choice ? "selected" : ""}>${escapeHtml(choice)}</option>`).join("")}
              </select>
            ` : `<input data-clarification-index="${index}" value="${escapeHtml(question.answer || "")}" inputmode="${question.answer_type === "number" ? "decimal" : "text"}">`}
          </label>
        `).join("")}
        <button class="primary-button" type="submit">Generate final Quote Basis</button>
      </form>
    </div>
  `;
  const form = qs("#clarificationForm");
  form?.addEventListener("submit", handleClarificationSubmit);
}

async function handleClarificationSubmit(event) {
  event.preventDefault();
  const inputs = Array.from(event.currentTarget.querySelectorAll("[data-clarification-index]"));
  inputs.forEach((input) => {
    const index = Number(input.dataset.clarificationIndex);
    if (!Number.isInteger(index) || !state.blockingClarificationQuestions[index]) return;
    state.blockingClarificationQuestions[index].answer = input.value.trim();
    state.blockingClarificationQuestions[index].status = input.value.trim() ? "answered" : "open";
  });
  if ((state.blockingClarificationQuestions || []).some((question) => !String(question.answer || "").trim())) {
    renderClarificationQuestions();
    return;
  }
  state.pendingFeedback = "Use these clarification answers to generate the final Quote Basis.";
  await handleDraftBasis({ finalAfterClarifications: true });
  state.pendingFeedback = "";
}

function clearAiFailedDraftState() {
  state.quoteBasis = { ...EMPTY_BASIS };
  state.quoteBasisSections = [];
  state.lineItems = [];
  state.outputRows = [];
  state.originalOutputRows = [];
  state.outputErrors = [];
  state.originalAnalysisSnapshot = null;
  state.basisConfirmed = false;
  setDownloadFiles([]);
  renderMatchSummary({});
  renderPricingMatches([]);
  clearPricingReviewMessages();
  setResultStatus("No usable AI draft", "is-bad");
}

function showAiFailedDraftState(data = {}) {
  clearAiFailedDraftState();
  state.aiFailed = true;
  state.draftSource = data.source || "local";
  state.lastAnalysisMode = normalizeAnalysisMode(data.analysis_mode || state.lastAnalysisMode);
  const message = genericFailureMessage(data);
  setWorkflowStage("basis_review");
  showAiFailureBanner(message);
  renderBasisFailureState(message);
  setSidePanel("basis", { force: true });
  syncControlStates();
}

function resetQuoteBasisToOriginal() {
  const snapshot = state.originalAnalysisSnapshot;
  if (!snapshot) return;
  state.quoteBasisSections = cloneQuoteBasisSections(snapshot.quote_basis_sections || snapshot.quote_basis || []);
  state.quoteBasis = cloneQuoteBasis(snapshot.quote_basis || quoteBasisFromSections(state.quoteBasisSections));
  state.lineItems = Array.isArray(snapshot.line_items) ? snapshot.line_items.map(normalizeLineItem) : [];
  state.boothDimensions = normalizeBoothDimensions(snapshot.boothDimensions || state.boothDimensions);
  state.lastAnalysisMode = normalizeAnalysisMode(snapshot.analysis_mode || state.lastAnalysisMode);
  state.outputRows = [];
  state.originalOutputRows = [];
  state.outputErrors = [];
  state.analysisFindings = [];
  state.blockingClarificationQuestions = [];
  state.basisConfirmed = false;
  setDownloadFiles([]);
  clearPricingReviewMessages();
  setWorkflowStage("basis_review");
  updateQuoteBasisCard(snapshot.source || "restored");
  syncControlStates();
}

async function resetOutputDraft() {
  if (state.isAnalysisRunning || state.isGenerating || !state.originalOutputRows.length) return;
  state.outputRows = snapshotOutputRows(state.originalOutputRows);
  state.lineItems = outputRowsToLineItems();
  state.outputErrors = [];
  state.downloadFile = null;
  setDownloadFiles([]);
  renderPricingMatches(state.outputRows);
  renderMatchSummary({ pricing_matches: state.outputRows });
  renderOutputValidationMessages(outputRowsValid().errors);
  setResultStatus("Output reset to confirmed basis", "is-warn");
  syncControlStates();
}

async function postJson(url, payload) {
  let response;
  const headers = { "Content-Type": "application/json" };
  if (state.csrfToken) headers[state.csrfHeaderName] = state.csrfToken;
  try {
    response = await fetch(url, {
      method: "POST",
      headers,
      body: JSON.stringify(payload),
    });
  } catch (error) {
    if (!state.isPageUnloading) {
      logClientEvent("client_error", { url, message: error.message || String(error) });
    }
    return {
      ok: false,
      data: {
        status: "failed",
        fetch_failed: true,
        page_unloading: state.isPageUnloading,
        message: error.message || String(error),
        errors: genericFailureMessages(),
      },
    };
  }

  let data;
  try {
    data = await response.json();
  } catch {
    data = {
      status: "failed",
      errors: genericFailureMessages(),
    };
  }
  if (!response.ok) {
    logClientEvent("server_error", { url, status: response.status, errors: data.errors || [] });
  }
  return { ok: response.ok, data };
}

async function getJson(url, options = {}) {
  let response;
  try {
    response = await fetch(url);
  } catch (error) {
    const message = error.message || String(error);
    if (!state.isPageUnloading && options.logFetchFailure !== false) {
      logClientEvent("client_error", { url, message });
    }
    return {
      ok: false,
      data: {
        status: "failed",
        fetch_failed: true,
        page_unloading: state.isPageUnloading,
        message,
        errors: genericFailureMessages(),
      },
    };
  }

  let data;
  try {
    data = await response.json();
  } catch {
    data = {
      status: "failed",
      errors: genericFailureMessages(),
    };
  }
  if (!response.ok) {
    logClientEvent("server_error", { url, status: response.status, errors: data.errors || [] });
  }
  return { ok: response.ok, data };
}

function delay(ms) {
  return new Promise((resolve) => {
    window.setTimeout(resolve, ms);
  });
}

async function startJob(type, payload) {
  return postJson("/api/jobs", { type, payload });
}

async function pollJob(jobId, onStatus) {
  const maxFetchFailures = 4;
  let fetchFailures = 0;
  while (jobId) {
    const url = `/api/jobs/${encodeURIComponent(jobId)}`;
    const { ok, data } = await getJson(url, { logFetchFailure: false });
    if (!ok) {
      if (data?.page_unloading) return { ok, data, aborted: true };
      if (data?.fetch_failed && fetchFailures < maxFetchFailures) {
        fetchFailures += 1;
        if (typeof onStatus === "function") onStatus({ status: "retrying", fetch_failures: fetchFailures });
        await delay(Math.min(1000 * fetchFailures, 4000));
        continue;
      }
      if (data?.fetch_failed) {
        logClientEvent("client_error", {
          url,
          message: data.message || "Failed to fetch",
          attempts: fetchFailures + 1,
        });
      }
      return { ok, data };
    }
    fetchFailures = 0;
    if (typeof onStatus === "function") onStatus(data);
    if (FINAL_JOB_STATUSES.has(data.status)) return { ok: true, data };
    await delay(900);
  }
  return { ok: false, data: { status: "failed", errors: genericFailureMessages() } };
}

function isInterruptedJobPoll(polled) {
  return Boolean(polled?.aborted || polled?.data?.fetch_failed);
}

function handleInterruptedJobPoll(jobType = "draft") {
  if (jobType === "draft") {
    state.isAnalysisRunning = false;
    showAiFailureBanner();
    setWorkflowStage("analyzing");
  } else {
    state.isGenerating = false;
    setResultStatus("Connection interrupted", "is-warn");
    renderMessages(genericFailureMessages(), "error");
  }
  syncControlStates();
}

function logClientEvent(event, details = {}) {
  const allowedEvents = new Set(["client_error", "server_error", "security_event", "abuse_signal"]);
  if (!allowedEvents.has(event)) return;
  const headers = { "Content-Type": "application/json" };
  if (state.csrfToken) headers[state.csrfHeaderName] = state.csrfToken;
  fetch("/api/log", {
    method: "POST",
    headers,
    body: JSON.stringify({ event, details }),
  }).catch(() => {});
}

async function initializeSession() {
  const { ok, data } = await getJson("/api/session");
  if (ok && data.csrf_token) {
    state.csrfHeaderName = data.csrf_header || CSRF_HEADER_NAME;
    state.csrfToken = data.csrf_token;
    if (data.permissions && typeof data.permissions === "object") {
      state.permissions = { ...state.permissions, ...data.permissions };
    }
    return;
  }
  elements.healthText.textContent = "Local session unavailable";
}

function syncControlStates() {
  const busy = state.isAnalysisRunning || state.isGenerating;
  elements.newQuoteButton.disabled = busy;
  elements.newQuoteButton.title = busy
    ? (state.isAnalysisRunning ? "Analysis is running." : "Quotation generation is running.")
    : "";
  if (elements.settingsButton) {
    const canManage = canManagePricingReferences();
    elements.settingsButton.hidden = false;
    elements.settingsButton.disabled = busy || !canManage;
    elements.settingsButton.title = canManage ? "Pricing reference settings" : pricingReferenceNoAccessReason();
    elements.settingsButton.setAttribute("aria-disabled", String(elements.settingsButton.disabled));
  }
  updateSidePanelNav();
  saveSessionState();
}

async function handleDraftBasis(options = {}) {
  if (state.isAnalysisRunning) return;
  const analysisMode = normalizeAnalysisMode(options.analysisMode || state.pendingAnalysisMode);
  state.pendingAnalysisMode = analysisMode;
  const runningMessage = analysisRunningMessage(analysisMode);
  const hasFeedback = Boolean(state.pendingFeedback.trim()) || options.finalAfterClarifications;
  const analysisRequestedAt = new Date().toISOString();

  if (!state.images.length) {
    setWorkflowStage("needs_images");
    showBlockedAction(`Please drop at least one ${currentGenerator().imageNoun} before analysis.`, { details: false });
    syncControlStates();
    return;
  }
  const missing = missingDetailFields();
  if (missing.length) {
    setWorkflowStage("details_review");
    showBlockedAction(`Fill Customer and Quote Company before AI analysis: ${missing.join(", ")}.`);
    renderBasisEmptyState("Complete the missing Customer and Quote Company details, then start analysis to draft the quote basis.");
    setDetailsDrawer(true);
    syncControlStates();
    return;
  }

  state.isAnalysisRunning = true;
  state.aiFailed = false;
  state.draftSource = "";
  state.lastAnalysisMode = analysisMode;
  clearAiFailureBanner();
  setWorkflowStage("analyzing");
  showAiRunningBanner(runningMessage, analysisRequestedAt);
  clearBasisReviewSurface();
  setSidePanel("basis", { force: true });
  syncControlStates();

  const started = await startJob("draft", buildPayload({
    analysisMode,
    includeBoothDimensions: state.boothDimensions.dimension_source !== "default",
    includeDraftContext: hasFeedback,
  }));
  if (!started.ok) {
    state.isAnalysisRunning = false;
    const errors = started.data.errors || ["Draft failed."];
    const wasBlocked = started.data.status === "blocked";
    setWorkflowStage(wasBlocked ? "details_review" : "ready_to_analyze");
    if (wasBlocked) {
      showAiBlockedBanner(errors.join(" "));
    } else {
      showAiFailureBanner(genericFailureMessage(started.data));
    }
    syncControlStates();
    return;
  }
  const startedAt = started.data.created_at || analysisRequestedAt;
  state.activeJob = { id: started.data.job_id, type: "draft", startedAt };
  showAiRunningBanner(runningMessage, startedAt);
  saveSessionState();

  const polled = await pollJob(started.data.job_id);
  if (polled.aborted) return;
  if (isInterruptedJobPoll(polled)) {
    handleInterruptedJobPoll("draft");
    return;
  }
  state.isAnalysisRunning = false;
  state.activeJob = null;

  if (!polled.ok || ["blocked", "failed"].includes(polled.data.status)) {
    const errors = polled.data.errors || polled.data.result?.errors || ["Draft failed."];
    const wasBlocked = polled.data.status === "blocked";
    setWorkflowStage(wasBlocked ? "details_review" : "ready_to_analyze");
    if (wasBlocked) {
      showAiBlockedBanner(errors.join(" "));
    } else {
      showAiFailureBanner(genericFailureMessage(polled.data.result || polled.data));
    }
    syncControlStates();
    return;
  }

  const data = polled.data.result || {};
  if (Array.isArray(data.blocking_clarification_questions) && data.blocking_clarification_questions.length) {
    clearAiFailureBanner();
    openBlockingClarifications(data.blocking_clarification_questions, data.analysis_findings || state.analysisFindings || [], data.project || state.boothDimensions || {});
    return;
  }
  const hasFinalBasis = Array.isArray(data.quote_basis_sections) && data.quote_basis_sections.length && Array.isArray(data.line_items) && data.line_items.length;
  if (options.finalAfterClarifications && !hasFinalBasis) {
    renderClarificationQuestions();
    return;
  }
  state.blockingClarificationQuestions = [];
  state.analysisFindings = data.analysis_findings || [];
  const aiFailed = Boolean(data.ai_failed || polled.data.status === "degraded" || (data.source === "local" && Array.isArray(data.warnings) && data.warnings.length));
  if (aiFailed) {
    showAiFailedDraftState(data);
    return;
  }
  applyDraftBasis(data);
  applyDraftLineItems(data.line_items || []);
  state.boothDimensions = normalizeBoothDimensions(data.project || state.boothDimensions);
  state.draftSource = data.source || "";
  state.lastAnalysisMode = normalizeAnalysisMode(data.analysis_mode || analysisMode);
  captureOriginalAnalysisSnapshot(data);
  state.aiFailed = false;
  setWorkflowStage("basis_review");
  clearAiFailureBanner();
  updateQuoteBasisCard(data.source);
  setSidePanel("basis", { force: true });
  syncControlStates();
}

async function confirmBasis() {
  if (state.isAnalysisRunning || state.isGenerating) return;
  const confirmBlockReason = basisConfirmBlockReason();
  if (confirmBlockReason) {
    setWorkflowStage("basis_review");
    syncControlStates();
    return;
  }
  if (state.aiFailed) {
    showAiFailureBanner();
    return;
  }
  if (!state.lineItems.length) {
    setWorkflowStage("basis_review");
    return;
  }
  const missing = missingDetailFields();
  if (missing.length) {
    setWorkflowStage("details_review");
    setDetailsDrawer(true);
    syncControlStates();
    return;
  }
  const refreshed = await refreshLineItemsFromServer();
  state.basisConfirmed = true;
  refreshOutputRowsFromLineItems();
  state.originalOutputRows = snapshotOutputRows(state.outputRows);
  state.lineItems = outputRowsToLineItems();
  setWorkflowStage("completed");
  renderPricingMatches(state.outputRows);
  renderMatchSummary({ pricing_matches: state.outputRows });
  setResultStatus(refreshed ? "Ready for pricing review" : "Pricing refresh unavailable", "is-warn");
  renderOutputValidationMessages(outputRowsValid().errors);
  setSidePanel("output", { force: true });
  syncControlStates();
}

async function handleGenerate() {
  if (state.isGenerating) return;
  if (state.activeSidePanel === "output") {
    const validation = outputRowsValid();
    if (!validation.valid) {
      renderOutputValidationMessages(validation.errors);
      setResultStatus("Output needs review", "is-warn");
      syncControlStates();
      return;
    }
    state.lineItems = outputRowsToLineItems();
  }
  if (!state.basisConfirmed) {
    if (hasSubmittedQuoteBasis()) {
      setWorkflowStage("basis_review");
      setSidePanel("basis", { force: true });
    }
    syncControlStates();
    return;
  }
  if (state.aiFailed) {
    setWorkflowStage("basis_review");
    showAiFailureBanner();
    syncControlStates();
    return;
  }
  const missing = missingDetailFields();
  if (missing.length) {
    setWorkflowStage("details_review");
    setDetailsDrawer(true);
    syncControlStates();
    return;
  }
  if (!state.lineItems.length) {
    setWorkflowStage("basis_review");
    return;
  }

  state.isGenerating = true;
  setWorkflowStage("generating");
  setResultStatus("Generating Excel", "is-warn");
  renderMessages([]);
  setDownloadFiles([]);
  renderMatchSummary({});
  clearPricingReviewMessages();
  syncControlStates();
  const started = await startJob("generate", buildPayload());
  if (!started.ok) {
    state.isGenerating = false;
    setWorkflowStage(state.activeSidePanel === "output" ? "completed" : "details_review");
    setResultStatus(started.data.status || "Failed", "is-bad");
    renderMessages(started.data.status === "blocked" ? (started.data.errors || ["Generation blocked."]) : genericFailureMessages(started.data), "error");
    syncControlStates();
    return;
  }
  state.activeJob = { id: started.data.job_id, type: "generate" };
  saveSessionState();

  const polled = await pollJob(started.data.job_id);
  if (polled.aborted) return;
  if (isInterruptedJobPoll(polled)) {
    handleInterruptedJobPoll("generate");
    return;
  }
  state.isGenerating = false;
  state.activeJob = null;

  const data = polled.data.result || polled.data || {};
  if (!polled.ok || ["blocked", "failed"].includes(polled.data.status) || data.status === "blocked" || data.status === "failed") {
    setWorkflowStage(state.activeSidePanel === "output" ? "completed" : "details_review");
    setResultStatus(data.status || "Failed", "is-bad");
    const blocked = polled.data.status === "blocked" || data.status === "blocked";
    renderMessages(blocked ? (data.errors || ["Generation blocked."]) : genericFailureMessages(data || polled.data), "error");
    if (data.pricing_matches?.length) renderPricingMatches(data.pricing_matches || [], { fromPricingMatches: true });
    renderMatchSummary(data);
    syncControlStates();
    return;
  }

  const needsPricingReview = polled.data.status === "needs_review"
    || data.status === "needs_confirmation";
  if (needsPricingReview) {
    setWorkflowStage("completed");
    setResultStatus("Needs pricing review", "is-warn");
    renderMessages([]);
    clearPricingReviewMessages();
    setSidePanel("output", { force: true });
    setDownloadFiles([]);
    if (data.pricing_matches?.length) renderPricingMatches(data.pricing_matches || [], { fromPricingMatches: true });
    renderMatchSummary(data);
  } else {
    setWorkflowStage("completed");
    setResultStatus("Completed", "is-ok");
    renderMessages([]);
    clearPricingReviewMessages();
    setSidePanel("output", { force: true });
    setDownloadFiles(data.files || []);
    renderPricingMatches(state.outputRows);
    renderMatchSummary({ pricing_matches: state.outputRows });
  }
  syncControlStates();
}

async function resumeSavedJob() {
  const activeJob = state.activeJob;
  if (!activeJob || !activeJob.id) return;

  if (activeJob.type === "draft") {
    state.isAnalysisRunning = true;
    state.isGenerating = false;
    setWorkflowStage("analyzing");
    showAiRunningBanner("Resuming the analysis job after refresh.", activeJobStartedAt(activeJob));
    clearBasisReviewSurface();
    setSidePanel("basis", { force: true });
    syncControlStates();

    const polled = await pollJob(activeJob.id, (job) => {
      if (job.created_at && !state.activeJob?.startedAt) {
        state.activeJob = { ...state.activeJob, startedAt: job.created_at };
        showAiRunningBanner("Resuming the analysis job after refresh.", job.created_at);
        saveSessionState();
      }
    });
    if (polled.aborted) return;
    if (isInterruptedJobPoll(polled)) {
      handleInterruptedJobPoll("draft");
      return;
    }
    state.isAnalysisRunning = false;
    state.activeJob = null;

    if (!polled.ok || ["blocked", "failed"].includes(polled.data.status)) {
      setWorkflowStage("ready_to_analyze");
      showAiFailureBanner(genericFailureMessage(polled.data.result || polled.data));
      syncControlStates();
      return;
    }

    const data = polled.data.result || {};
    const aiFailed = Boolean(data.ai_failed || polled.data.status === "degraded" || (data.source === "local" && Array.isArray(data.warnings) && data.warnings.length));
    if (aiFailed) {
      showAiFailedDraftState(data);
      return;
    }
    applyDraftBasis(data);
    applyDraftLineItems(data.line_items || []);
    state.boothDimensions = normalizeBoothDimensions(data.project || state.boothDimensions);
    state.draftSource = data.source || "";
    captureOriginalAnalysisSnapshot(data);
    state.aiFailed = false;
    setWorkflowStage("basis_review");
    clearAiFailureBanner();
    updateQuoteBasisCard(data.source);
    setSidePanel("basis", { force: true });
    syncControlStates();
    return;
  }

  if (activeJob.type === "generate") {
    state.isGenerating = true;
    state.isAnalysisRunning = false;
    setWorkflowStage("generating");
    setResultStatus("Checking Excel", "is-warn");
    setSidePanel("output", { force: true });
    syncControlStates();

    const polled = await pollJob(activeJob.id);
    if (polled.aborted) return;
    if (isInterruptedJobPoll(polled)) {
      handleInterruptedJobPoll("generate");
      return;
    }
    state.isGenerating = false;
    state.activeJob = null;

    const data = polled.data.result || polled.data || {};
    if (!polled.ok || ["blocked", "failed"].includes(polled.data.status) || data.status === "blocked" || data.status === "failed") {
      setWorkflowStage("details_review");
      setResultStatus(data.status || "Failed", "is-bad");
      const blocked = polled.data.status === "blocked" || data.status === "blocked";
      renderMessages(blocked ? (data.errors || ["Generation blocked."]) : genericFailureMessages(data || polled.data), "error");
      if (data.pricing_matches?.length) renderPricingMatches(data.pricing_matches || [], { fromPricingMatches: true });
      renderMatchSummary(data);
      syncControlStates();
      return;
    }

    const needsPricingReview = polled.data.status === "needs_review"
      || data.status === "needs_confirmation";
    if (needsPricingReview) {
      setWorkflowStage("completed");
      setResultStatus("Needs pricing review", "is-warn");
      renderMessages([]);
      clearPricingReviewMessages();
      setSidePanel("output", { force: true });
      setDownloadFiles([]);
    } else {
      setWorkflowStage("completed");
      setResultStatus("Completed", "is-ok");
      renderMessages([]);
      clearPricingReviewMessages();
      setSidePanel("output");
      setDownloadFiles(data.files || []);
    }
    if (data.pricing_matches?.length) renderPricingMatches(data.pricing_matches || [], { fromPricingMatches: true });
    renderMatchSummary(data);
    syncControlStates();
    return;
  }

  state.activeJob = null;
  syncControlStates();
}

async function checkHealth() {
  try {
    const response = await fetch("/api/health");
    const data = await response.json();
    elements.statusDot.classList.toggle("is-ok", response.ok);
    if (elements.topbarStatus) elements.topbarStatus.classList.toggle("is-ok", response.ok);
    elements.healthText.textContent = response.ok ? "Ready" : data.error || "Unavailable";
  } catch {
    elements.healthText.textContent = "Unavailable";
    if (elements.topbarStatus) elements.topbarStatus.classList.remove("is-ok");
  }
}

function requestStartAnalysis(mode = "standard") {
  const reason = startAnalysisBlockReason();
  if (reason) {
    showBlockedAction(reason);
    syncControlStates();
    return;
  }
  state.pendingAnalysisMode = normalizeAnalysisMode(mode);
  elements.analysisConfirmModal.hidden = false;
  elements.analysisConfirmModal.classList.add("is-open");
  window.setTimeout(() => elements.analysisConfirmStartButton?.focus(), 0);
}

function closeAnalysisConfirmModal() {
  elements.analysisConfirmModal.classList.remove("is-open");
  elements.analysisConfirmModal.hidden = true;
}

function confirmStartAnalysis(mode = state.pendingAnalysisMode) {
  closeAnalysisConfirmModal();
  handleDraftBasis({ analysisMode: normalizeAnalysisMode(mode) });
}

function isSensitiveChatRequest(normalizedText) {
  return [
    "api key",
    ".env",
    "authorization",
    "bearer",
    "secret",
    "token",
    "system prompt",
    "hidden prompt",
    "developer message",
    "internal prompt",
  ].some((term) => normalizedText.includes(term));
}

function sidePanelBlockReason(panelName) {
  if (panelName === "images") return "";
  if (!state.images.length) return "Add reference files before opening this step.";
  if (panelName === "quote_company") {
    return customerDetailsBlockReason("Complete Customer details before opening Quote Company");
  }
  if (panelName === "basis") {
    const missing = missingDetailFields();
    if (missing.length) return `Complete Customer and Quote Company details before opening Quote Basis: ${missing.join(", ")}.`;
    if (!hasSubmittedQuoteBasis()) return "Click Start Analysis from Quote Company before opening Quote Basis.";
  }
  if (panelName === "output") {
    const missing = missingDetailFields();
    if (missing.length) return `Complete Customer and Quote Company details before opening Output: ${missing.join(", ")}.`;
    if (!hasSubmittedQuoteBasis()) return "Click Start Analysis from Quote Company before opening Output.";
    if ((state.blockingClarificationQuestions || []).some((question) => question.status !== "answered" && !String(question.answer || "").trim())) return "Answer all clarification questions before opening Output.";
    const confirmBlockReason = basisConfirmBlockReason();
    if (confirmBlockReason) return confirmBlockReason;
    if (!hasCompletedQuoteBasis()) return "Confirm Quotation Basis before opening Output.";
  }
  return "";
}

function setSidePanel(panelName, options = {}) {
  const panelTitles = {
    images: ["Images", "Reference Inputs", currentGenerator().intakeSubtitle],
    customer: ["Customer", "Customer Details", "Customer, project, booth size, and customer address for this quotation."],
    quote_company: ["Quote Company", "Quotation Defaults", "Reusable quotation-company header, payment terms, notes, and signature defaults."],
    basis: ["Quote Basis", "Confirm Draft", "Review the drafted basis and revise individual lines where needed."],
    output: ["Output", "Editable Pricing", "Review quotation rows, resolve pricing, and download Excel."],
  };
  const nextPanel = panelTitles[panelName] ? panelName : "images";
  const blockReason = sidePanelBlockReason(nextPanel);
  if (blockReason && !options.force) {
    updateSidePanelNav();
    return false;
  }
  const [title, eyebrow, subtitle] = panelTitles[nextPanel] || panelTitles.images;
  state.activeSidePanel = nextPanel;
  document.body.dataset.sidePanel = state.activeSidePanel;
  elements.sideDrawerTitle.textContent = title;
  elements.sideDrawerEyebrow.textContent = eyebrow;
  elements.sideDrawerSubtitle.textContent = subtitle || "";
  document.querySelectorAll("button[data-side-panel]").forEach((button) => {
    button.classList.toggle("is-active", button.dataset.sidePanel === state.activeSidePanel);
  });
  document.querySelectorAll(".side-panel-section[data-side-panel-content]").forEach((section) => {
    section.classList.toggle("is-active", section.dataset.sidePanelContent === state.activeSidePanel);
  });
  if (state.activeSidePanel === "basis" && (state.blockingClarificationQuestions || []).length && !state.quoteBasisSections.length && !state.lineItems.length) {
    renderClarificationQuestions();
  }
  elements.sideWorkspace.setAttribute("aria-hidden", "false");
  updateSidePanelNav();
  saveSessionState();
  return true;
}

function activeSidePanelIndex() {
  return Math.max(0, SIDE_PANEL_SEQUENCE.indexOf(state.activeSidePanel));
}

function updateSidePanelNav() {
  if (!elements.sideBackButton || !elements.sideNextButton) return;
  const index = activeSidePanelIndex();
  const lastIndex = SIDE_PANEL_SEQUENCE.length - 1;
  const isQuoteCompanyStep = state.activeSidePanel === "quote_company";
  const isBasisStep = state.activeSidePanel === "basis";
  const isOutputStep = state.activeSidePanel === "output";
  const busy = state.isAnalysisRunning || state.isGenerating;
  elements.sampleDetailsButton.hidden = state.activeSidePanel !== "images";
  elements.sampleDetailsButton.disabled = state.isBooting || busy;
  elements.resetImagesButton.hidden = state.activeSidePanel !== "images";
  elements.clearCustomerButton.hidden = state.activeSidePanel !== "customer";
  elements.clearQuoteCompanyButton.hidden = state.activeSidePanel !== "quote_company";
  elements.analyseAgainButton.hidden = state.activeSidePanel !== "basis";
  elements.resetQuoteBasisButton.hidden = state.activeSidePanel !== "basis";
  elements.resetOutputButton.hidden = state.activeSidePanel !== "output";
  elements.resetImagesButton.disabled = busy || !state.images.length;
  elements.clearCustomerButton.disabled = busy;
  elements.clearQuoteCompanyButton.disabled = busy;
  elements.analyseAgainButton.disabled = busy || !state.images.length;
  elements.resetQuoteBasisButton.disabled = busy || !state.originalAnalysisSnapshot;
  elements.resetOutputButton.disabled = busy || !state.originalOutputRows.length;
  document.querySelectorAll("button[data-side-panel]").forEach((button) => {
    const panelName = button.dataset.sidePanel || "images";
    const blockReason = sidePanelBlockReason(panelName);
    const locked = Boolean(blockReason);
    button.disabled = busy || locked;
    button.classList.toggle("is-locked", locked);
    button.setAttribute("aria-disabled", String(button.disabled));
    button.title = locked ? blockReason : "";
  });
  elements.sideBackButton.disabled = index === 0 || busy;
  elements.sideNextButton.hidden = isOutputStep;
  elements.sideDownloadButton.hidden = !isOutputStep;
  const nextLabels = {
    images: "Next: Customer",
    customer: "Next: Quote Company",
    quote_company: currentGenerator().analyzeLabel,
    basis: "Confirm Quotation Basis",
  };
  elements.sideNextButton.textContent = nextLabels[state.activeSidePanel] || "Next";
  elements.sideNextButton.classList.add("primary-button");
  elements.sideNextButton.classList.remove("secondary-button");
  const nextPanel = SIDE_PANEL_SEQUENCE[Math.min(index + 1, lastIndex)];
  const basisBlockReason = isBasisStep ? basisConfirmBlockReason() : "";
  let nextBlockReason = "";
  if (busy) {
    nextBlockReason = state.isAnalysisRunning ? "Analysis is already running." : "Quotation generation is already running.";
  } else if (isQuoteCompanyStep) {
    nextBlockReason = startAnalysisBlockReason();
  } else if (isBasisStep) {
    nextBlockReason = basisBlockReason;
  } else {
    nextBlockReason = sidePanelBlockReason(nextPanel);
  }
  elements.sideNextButton.disabled = false;
  elements.sideNextButton.setAttribute("aria-disabled", String(Boolean(nextBlockReason)));
  elements.sideNextButton.classList.toggle("is-disabled", Boolean(nextBlockReason));
  elements.sideNextButton.title = nextBlockReason || "";
  updateDownloadButton();
}

function goToPreviousSidePanel() {
  const index = activeSidePanelIndex();
  if (index <= 0) return;
  setSidePanel(SIDE_PANEL_SEQUENCE[index - 1]);
}

function goToNextSidePanel() {
  const index = activeSidePanelIndex();
  if (elements.sideNextButton?.getAttribute("aria-disabled") === "true") {
    const reason = elements.sideNextButton.title || "This step is not ready yet.";
    elements.sideNextButton.title = "";
    elements.sideNextButton.blur();
    if (state.activeSidePanel === "quote_company") showBlockedAction(reason);
    return;
  }
  if (state.activeSidePanel === "quote_company") {
    requestStartAnalysis();
    return;
  }
  if (state.activeSidePanel === "basis") {
    confirmBasis();
    return;
  }
  if (state.activeSidePanel === "output") return;
  setSidePanel(SIDE_PANEL_SEQUENCE[index + 1], { notify: true });
}

function handleQuoteBasisClick(event) {
  const sectionAction = event.target.closest("[data-basis-section-action]");
  if (sectionAction) {
    retagBasisSectionConfirmLines(
      sectionAction.dataset.basisSection || "",
      sectionAction.dataset.basisSectionAction || ""
    );
    return;
  }
  const tagButton = event.target.closest("[data-basis-tag]");
  if (tagButton) {
    retagBasisLine(
      tagButton.dataset.basisSection || "",
      Number(tagButton.dataset.basisLineIndex),
      tagButton.dataset.basisTag || ""
    );
    return;
  }
  const reviseButton = event.target.closest("[data-revise-section]");
  if (reviseButton) {
    const sectionId = reviseButton.dataset.reviseSection || "";
    const lineIndex = Number(reviseButton.dataset.reviseLineIndex);
    const section = state.quoteBasisSections.find((item) => item.id === sectionId);
    const line = section?.lines?.[lineIndex];
    openBasisChatOverlay("line", {
      sectionId,
      lineIndex,
      line: line ? basisChatLineContext(line) : "",
      quantity: line?.quantity ?? "",
      unit: line?.unit || "",
      quantityLabel: line ? basisQuantityLabel(line) : "",
    });
    return;
  }
}

function markPageUnloading() {
  state.isPageUnloading = true;
}

function wireEvents() {
  wireRichTextEditors();
  if (elements.pricingReferenceFile) {
    elements.pricingReferenceFile.accept = PRICING_REFERENCE_FILE_ACCEPT;
  }
  if (elements.imageInput) {
    elements.imageInput.accept = REFERENCE_FILE_ACCEPT;
  }
  window.addEventListener("pagehide", markPageUnloading);
  window.addEventListener("beforeunload", markPageUnloading);
  elements.imageInput.addEventListener("change", async (event) => {
    await addImagesFromFiles(event.target.files || []);
    elements.imageInput.value = "";
  });

  elements.headerLogoInput.addEventListener("change", async (event) => {
    const file = (event.target.files || [])[0];
    state.headerLogo = file
      ? {
          name: file.name,
          type: file.type,
          size: file.size,
          data_url: await fileToDataUrl(file),
        }
      : null;
    renderHeaderLogoPreview();
    renderPresetStatus();
    syncControlStates();
  });
  elements.headerLogoPreview.addEventListener("click", () => {
    elements.headerLogoInput.click();
  });
  elements.headerLogoPreview.addEventListener("keydown", (event) => {
    if (!["Enter", " "].includes(event.key)) return;
    event.preventDefault();
    elements.headerLogoInput.click();
  });

  elements.dropzone.addEventListener("dragenter", (event) => {
    event.preventDefault();
    elements.dropzone.classList.add("is-dragging");
  });

  elements.dropzone.addEventListener("dragover", (event) => {
    event.preventDefault();
    event.dataTransfer.dropEffect = "copy";
    elements.dropzone.classList.add("is-dragging");
  });

  elements.dropzone.addEventListener("dragleave", (event) => {
    if (!elements.dropzone.contains(event.relatedTarget)) {
      elements.dropzone.classList.remove("is-dragging");
    }
  });

  elements.dropzone.addEventListener("drop", async (event) => {
    event.preventDefault();
    elements.dropzone.classList.remove("is-dragging");
    await addImagesFromFiles(event.dataTransfer.files || []);
  });

  elements.fileList.addEventListener("click", (event) => {
    const button = event.target.closest("[data-remove-image]");
    if (!button) return;
    removeImageAt(Number(button.dataset.removeImage));
  });

  document.querySelectorAll("button[data-side-panel]").forEach((button) => {
    button.addEventListener("click", () => setSidePanel(button.dataset.sidePanel || "images", { notify: true }));
  });
  elements.sideBackButton.addEventListener("click", goToPreviousSidePanel);
  elements.sideNextButton.addEventListener("click", goToNextSidePanel);
  elements.sideDownloadButton.addEventListener("click", async (event) => {
    if (elements.sideDownloadButton.getAttribute("aria-disabled") === "true") {
      event.preventDefault();
      const validation = outputRowsValid();
      if (!validation.valid) renderOutputValidationMessages(validation.errors);
      return;
    }
    if (!state.downloadFile?.url) {
      event.preventDefault();
      await handleGenerate();
      downloadCurrentExcelFile();
    }
  });
  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape") {
      if (elements.pricingReferenceTableOverlay && !elements.pricingReferenceTableOverlay.hidden) {
        closePricingReferenceTableOverlay();
      } else if (!elements.basisChatOverlay.hidden) {
        closeBasisChatOverlay();
      } else if (elements.pricingReferenceModal && !elements.pricingReferenceModal.hidden) {
        closePricingReferenceModal();
      } else if (!elements.analysisConfirmModal.hidden) {
        closeAnalysisConfirmModal();
      }
    }
  });

  elements.sampleDetailsButton.addEventListener("click", setSampleDetails);
  elements.newQuoteButton.addEventListener("click", startNewQuote);
  elements.settingsButton?.addEventListener("click", openSettingsModal);
  elements.profileSelect.addEventListener("change", handleProfileSelectionChange);
  elements.newPricingReferenceButton?.addEventListener("click", openPricingReferenceModal);
  elements.deletePricingReferenceSelect?.addEventListener("change", updatePricingReferenceDeleteButton);
  elements.deletePricingReferenceButton?.addEventListener("click", deleteSelectedPricingReference);
  elements.outputSortMode?.addEventListener("change", () => { state.outputSortMode = elements.outputSortMode.value; renderPricingMatches(state.outputRows); renderMatchSummary({ pricing_matches: state.outputRows }); syncControlStates(); });
  elements.pricingReferenceForm.addEventListener("submit", savePricingReferenceFromModal);
  elements.pricingReferenceTemplateButton.addEventListener("click", downloadPricingReferenceTemplate);
  elements.pricingReferenceFile.addEventListener("change", async () => {
    const file = elements.pricingReferenceFile.files?.[0];
    if (elements.pricingReferenceFileName) {
      elements.pricingReferenceFileName.textContent = file?.name || "No file chosen";
    }
    if (!file) {
      state.pendingPricingReference = null;
      state.pricingReferenceImportBusy = false;
      state.pricingReferenceImportToken = "";
      renderPricingReferencePreview(null);
      return;
    }
    const importToken = `${Date.now()}-${file.name}-${file.size}`;
    state.pricingReferenceImportToken = importToken;
    state.pricingReferenceImportBusy = true;
    setPricingReferenceSaveButtonState({ busy: true, reason: pricingReferenceSaveBlockReason(null) });
    renderPricingReferencePreview({
      ...pricingReferenceValidationResult([], [], 0, file.name),
      layout: "importing",
      warnings: ["Import preview is still being prepared."],
      errors: [],
    });
    let result;
    try {
      result = await validatePricingReferenceFile(file);
    } catch (error) {
      result = {
        ...pricingReferenceValidationResult([], [], 0, file.name),
        errors: genericFailureMessages(),
      };
    }
    if (state.pricingReferenceImportToken !== importToken) return;
    state.pendingPricingReference = result;
    state.pricingReferenceImportBusy = false;
    renderPricingReferencePreview(result, { syncCurrencyControls: true, scrollIntoView: true });
  });
  elements.pricingReferenceCurrency?.addEventListener("change", () => {
    syncPricingReferenceCurrencyCustomInput();
    if (state.pendingPricingReference) {
      state.pendingPricingReference.currency = pricingReferenceModalCurrency();
      renderPricingReferencePreview(state.pendingPricingReference);
    }
  });
  elements.pricingReferenceCurrencyCustom?.addEventListener("input", () => {
    const normalizedValue = normalizedCustomCurrencyInput();
    if (elements.pricingReferenceCurrencyCustom.value !== normalizedValue) {
      elements.pricingReferenceCurrencyCustom.value = normalizedValue;
    }
    if (state.pendingPricingReference) {
      state.pendingPricingReference.currency = pricingReferenceModalCurrency();
      renderPricingReferencePreview(state.pendingPricingReference);
      return;
    }
    setPricingReferenceSaveButtonState({
      canSave: !pricingReferenceSaveBlockReason(state.pendingPricingReference),
      reason: pricingReferenceSaveBlockReason(state.pendingPricingReference),
    });
  });
  elements.pricingReferenceTaxLabel?.addEventListener("change", () => {
    if (state.pendingPricingReference) state.pendingPricingReference.tax = pricingReferenceModalTax();
  });
  elements.pricingReferenceTaxRate?.addEventListener("input", () => {
    if (state.pendingPricingReference) state.pendingPricingReference.tax = pricingReferenceModalTax();
  });
  elements.pricingReferencePreview?.addEventListener("input", (event) => {
    handlePricingReferencePreviewInput(event);
  });
  elements.pricingReferencePreview?.addEventListener("click", (event) => {
    if (event.target.closest("[data-pricing-reference-table-open]")) openPricingReferenceTableOverlay();
  });
  elements.pricingReferenceTableBody?.addEventListener("input", handlePricingReferencePreviewInput);
  elements.pricingReferenceTableCloseButton?.addEventListener("click", closePricingReferenceTableOverlay);
  elements.pricingReferenceTableOverlay?.addEventListener("click", (event) => {
    if (event.target.closest("[data-pricing-reference-table-close]")) closePricingReferenceTableOverlay();
  });
  elements.pricingReferenceCancelButton.addEventListener("click", closePricingReferenceModal);
  elements.pricingReferenceCloseButton.addEventListener("click", closePricingReferenceModal);
  elements.pricingReferenceModal.addEventListener("click", blockPricingReferenceBusyInteraction, true);
  elements.pricingReferenceModal.addEventListener("click", (event) => {
    if (event.target.closest("[data-pricing-reference-close]")) closePricingReferenceModal();
  });
  elements.analysisConfirmCancelButton.addEventListener("click", closeAnalysisConfirmModal);
  elements.analysisConfirmStartButton.addEventListener("click", () => confirmStartAnalysis("standard"));
  elements.analysisConfirmModal.addEventListener("click", (event) => {
    if (event.target.closest("[data-analysis-confirm-close]")) closeAnalysisConfirmModal();
  });
  elements.pricingMatchesBody.addEventListener("pointerdown", handleOutputIncludedPointerDown);
  elements.pricingMatchesBody.addEventListener("click", handleOutputCellClick);
  elements.pricingMatchesBody.addEventListener("dblclick", handleOutputCellOpen);
  elements.pricingMatchesBody.addEventListener("keydown", handleOutputCellKeydown);
  elements.pricingMatchesBody.addEventListener("focusout", handleOutputEditorCommit);
  elements.pricingMatchesBody.addEventListener("change", handleOutputEditorCommit);
  elements.resetImagesButton.addEventListener("click", resetImagesDraft);
  elements.analyseAgainButton.addEventListener("click", () => {
    state.pendingFeedback = "";
    requestStartAnalysis("standard");
  });
  elements.resetQuoteBasisButton.addEventListener("click", resetQuoteBasisToOriginal);
  elements.resetOutputButton.addEventListener("click", resetOutputDraft);
  elements.presetSelect.addEventListener("change", () => {
    state.selectedPresetValue = elements.presetSelect.value || "";
    updatePresetButtons();
    renderPresetStatus();
    syncControlStates();
  });
  elements.savePresetButton.addEventListener("click", saveCurrentPreset);
  elements.loadPresetButton.addEventListener("click", loadSelectedPreset);
  elements.deletePresetButton.addEventListener("click", deleteSelectedPreset);
  elements.clearCustomerButton.addEventListener("click", clearCustomerDetails);
  elements.clearQuoteCompanyButton.addEventListener("click", clearQuoteCompanyDetails);
  [
    elements.clientName,
    elements.clientAttention,
    elements.clientTitle,
    elements.clientAddress,
    elements.projectTitle,
    elements.quoteDate,
    elements.projectNumber,
    elements.headerDetails,
    elements.termsHeading,
    elements.paymentTerms,
    elements.notesHeading,
    elements.standardNotes,
    elements.quoteCompanyName,
    elements.acceptanceText,
    elements.konceptSignatory,
    elements.konceptTitle,
    elements.konceptDateLabel,
    elements.personLabel,
    elements.stampLabel,
    elements.dateLabel,
    elements.taxLabel,
    elements.taxRate,
  ].filter(Boolean).forEach((input) => {
    input.addEventListener("input", syncControlStates);
    input.addEventListener("change", syncControlStates);
  });
  elements.basisChatForm.addEventListener("submit", handleBasisChatSubmit);
  elements.basisChatApplyButton.addEventListener("click", applyBasisChatProposal);
  elements.basisChatKeepButton.addEventListener("click", keepCurrentBasis);
  elements.basisChatCloseButton.addEventListener("click", closeBasisChatOverlay);
  elements.basisChatOverlay.addEventListener("click", (event) => {
    if (event.target.closest("[data-basis-chat-close]")) {
      closeBasisChatOverlay();
    }
  });
  elements.basisChatMessages.addEventListener("click", (event) => {
    const button = event.target.closest("[data-basis-chat-action]");
    if (!button) return;
    if (button.dataset.basisChatAction === "apply") applyBasisChatProposal();
    if (button.dataset.basisChatAction === "discard") keepCurrentBasis();
  });
  elements.basisChatPrompt.addEventListener("keydown", (event) => {
    if (event.key !== "Enter" || event.shiftKey) return;
    event.preventDefault();
    if (!elements.basisChatSendButton.disabled) {
      elements.basisChatForm.requestSubmit();
    }
  });
  elements.basisReviewSurface.addEventListener("click", handleQuoteBasisClick);
  elements.pricingMatchesBody.addEventListener("input", handleOutputRowEdit);
  elements.pricingMatchesBody.addEventListener("change", handleOutputRowEdit);
}

async function setInitialValues() {
  updateGeneratorCopy();
  renderProfileOptions();
  renderPresetOptions();
  renderHeaderLogoPreview();
  renderPresetStatus();
  if (await restoreSessionState()) {
    syncControlStates();
    return;
  }
  applyDefaultQuoteDate();
  applyDefaultQuoteCompanyFields();
  loadDefaultProfilePreset({ silent: true });
  renderFiles();
  renderPricingMatches([]);
  renderMatchSummary({});
  setWorkflowStage("needs_images");
  renderBasisEmptyState();
  syncControlStates();
  setSidePanel("images");
}

async function boot() {
  wireEvents();
  clearLegacyLocalCompanyPresets();
  try {
    await initializeSession();
    await loadProfiles();
    await setInitialValues();
    checkHealth();
  } finally {
    state.isBooting = false;
    syncControlStates();
  }
  await resumeSavedJob();
}

boot();
