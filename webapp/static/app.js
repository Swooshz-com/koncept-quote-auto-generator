const EMPTY_LINE_ITEMS_MESSAGE = "AI analysis will populate line items here.";
const DEFAULT_PROFILE_ID = "koncept";
const DEFAULT_PRICING_REFERENCE_ID = "koncept";
const DEFAULT_SAMPLE_ID = "brazil-pavilion";
const CSRF_HEADER_NAME = "X-Swooshz-CSRF";
const QUOTE_PRESETS_STORAGE_KEY = "swooshz_quote_detail_presets_v1";
const QUOTE_SESSION_STORAGE_KEY = "swooshz_quote_session_v1";
const QUOTE_SESSION_STATE_VERSION = 3;
const LOCAL_PRICING_REFERENCES_STORAGE_KEY = "swooshz_pricing_references_v1";
const FINAL_JOB_STATUSES = new Set(["completed", "degraded", "needs_review", "blocked", "failed"]);
const PROFILE_PRESET_PREFIX = "profile:";
const LOCAL_PRESET_PREFIX = "local:";
const PRICING_REFERENCE_FILE_ACCEPT = ".xlsx,.csv";
const MAX_REFERENCE_IMAGES = 8;
const DEFAULT_DATE_LABEL = "Date:";
const DEFAULT_TERMS_HEADING = "Terms & Conditions:";
const DEFAULT_NOTES_HEADING = "Note:";
const DEFAULT_ACCEPTANCE_TEXT = "We accept the quotation amount and the terms";
const DEFAULT_PERSON_LABEL = "Person in charge";
const DEFAULT_STAMP_LABEL = "Company name & stamp";
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
  assistantSubtitle: "Start with reference images or a sample fixture, confirm one quotation basis, then generate Excel.",
  intakeSubtitle: "Drop reference images for a real quote, or load the demo fixture for a quick test run.",
  dropTitle: "Drop reference images to start",
  dropMeta: `JPG, PNG, or WebP. Up to ${MAX_REFERENCE_IMAGES} references, 12 MB each; files stay local to this runner.`,
  imageNoun: "reference image",
  analyzeLabel: "Start Analysis",
  fallbackAction: "review the reference images",
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
  ["Custom", "Custom", "Not found in pricing reference"],
  ["Confirm", "Confidence", "AI confidence before your decision"],
];

const STAGE_LABELS = {
  needs_images: "Needs images",
  ready_to_analyze: "Ready to analyze",
  analyzing: "Analyzing",
  basis_review: "Confirm basis",
  details_review: "Needs details",
  generating: "Generating",
  pricing_review: "Pricing review",
  completed: "Completed",
};

const state = {
  profileId: "",
  pricingReferenceId: "",
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
  boothDimensions: { ...DEFAULT_BOOTH_DIMENSIONS },
  originalAnalysisSnapshot: null,
  basisConfirmed: false,
  isAnalysisRunning: false,
  isGenerating: false,
  aiFailed: false,
  draftSource: "",
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
    proposal: null,
  },
};

const qs = (selector) => document.querySelector(selector);
let analysisElapsedTimerId = 0;

const elements = {
  healthText: qs("#healthText"),
  statusDot: qs("#statusDot"),
  topbarStatus: qs("#topbarStatus"),
  assistantSubtitle: qs("#assistantSubtitle"),
  imageIntake: qs("#imageIntake"),
  dropTitle: qs("#dropTitle"),
  dropMeta: qs("#dropMeta"),
  dropzone: qs(".dropzone"),
  imageInput: qs("#imageInput"),
  imageUploadStatus: qs("#imageUploadStatus"),
  fileList: qs("#fileList"),
  customerDetailsButton: qs("#customerDetailsButton"),
  quoteCompanyButton: qs("#quoteCompanyButton"),
  quoteBasisButton: qs("#quoteBasisButton"),
  outputButton: qs("#outputButton"),
  customerDetailsPanel: qs("#customerDetailsPanel"),
  quoteCompanyPanel: qs("#quoteCompanyPanel"),
  quoteBasisPanel: qs("#quoteBasisPanel"),
  outputSidePanel: qs("#outputSidePanel"),
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
  workflowStage: qs("#workflowStage"),
  aiFailureBanner: qs("#aiFailureBanner"),
  basisReviewSurface: qs("#basisReviewSurface"),
  busyText: qs("#busyText"),
  assistantOutput: qs("#assistantOutput"),
  resultStatus: qs("#resultStatus"),
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
  discussQuoteButton: qs("#discussQuoteButton"),
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
  pricingReferenceSaveButton: qs("#pricingReferenceSaveButton"),
  pricingReferenceCancelButton: qs("#pricingReferenceCancelButton"),
  pricingReferenceCloseButton: qs("#pricingReferenceCloseButton"),
  analysisConfirmModal: qs("#analysisConfirmModal"),
  analysisConfirmCancelButton: qs("#analysisConfirmCancelButton"),
  analysisConfirmStartButton: qs("#analysisConfirmStartButton"),
  richTextEditors: Array.from(document.querySelectorAll("[data-rich-text-source]")),
  richTextToolbar: Array.from(document.querySelectorAll("[data-rich-command]")),
};

function pricingReferenceSelectValue(reference = {}) {
  const referenceId = String(reference.id || "").trim();
  if (!referenceId) return "";
  const source = reference.source === "local" ? "local" : "bundled";
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

function safeLocalPricingReferences() {
  try {
    const parsed = JSON.parse(window.localStorage.getItem(LOCAL_PRICING_REFERENCES_STORAGE_KEY) || "[]");
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

function saveLocalPricingReferences(references = []) {
  const sanitized = references
    .filter((reference) => reference && reference.source === "local")
    .slice(0, 20);
  window.localStorage.setItem(LOCAL_PRICING_REFERENCES_STORAGE_KEY, JSON.stringify(sanitized));
}

function mergePricingReferences(bundled = []) {
  const localReferences = safeLocalPricingReferences();
  const seen = new Set();
  return [...bundled, ...localReferences].filter((reference) => {
    const key = pricingReferenceSelectValue(reference);
    if (!key || seen.has(key)) return false;
    seen.add(key);
    return true;
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

function normalizeUnit(value = "") {
  const text = String(value || "").trim();
  const lower = text.toLowerCase();
  if (["m2", "m^2", "sq m", "sq.m", "sq.m.", "square metre", "square meter", "square metres", "square meters"].includes(lower)) {
    return "sqm";
  }
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
  if (!pricingReferenceId) return null;
  return state.pricingReferences.find((reference) => reference.id === pricingReferenceId)
    || null;
}

function resolvedProfileIdForPayload() {
  return String(state.profileId || DEFAULT_PROFILE_ID).trim() || DEFAULT_PROFILE_ID;
}

function syncSelectedPricingReference() {
  const selectedReference = currentPricingReference();
  if (selectedReference) {
    state.pricingReferenceId = selectedReference.id || "";
    return;
  }
  const fallbackReference = defaultPricingReference();
  if (fallbackReference) {
    state.pricingReferenceId = fallbackReference.id || "";
    return;
  }
  state.pricingReferenceId = "";
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
  if (elements.assistantSubtitle) elements.assistantSubtitle.textContent = generator.assistantSubtitle;
  if (state.activeSidePanel === "images") elements.sideDrawerSubtitle.textContent = generator.intakeSubtitle;
  elements.dropTitle.textContent = atImageLimit
    ? "Maximum reference images added"
    : state.images.length ? "Add more reference images" : generator.dropTitle;
  elements.dropMeta.textContent = atImageLimit
    ? `Remove one reference to add another. Maximum ${MAX_REFERENCE_IMAGES} images.`
    : generator.dropMeta;
}

function setWorkflowStage(stage) {
  state.workflowStage = stage;
  if (elements.workflowStage) {
    elements.workflowStage.textContent = STAGE_LABELS[stage] || stage;
    elements.workflowStage.dataset.stage = stage;
  }
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

function showAiRunningBanner(message = "Reading the reference images and preparing the quote basis. Please wait.", startedAt = activeJobStartedAt()) {
  setAiStatusBanner("running", "AI analysis running.", message, { elapsed: true });
  startAnalysisElapsedTimer(startedAt);
}

function showAiFailureBanner(message = "AI analysis failed. Try again later.") {
  const detail = message.replace(/^AI analysis failed\.?\s*/i, "").trim() || "Try again later.";
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
    <span>Try again later.</span>
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

async function filesToImageEntries(files, limit = MAX_REFERENCE_IMAGES) {
  return Promise.all(
    Array.from(files)
      .filter((file) => file.type.startsWith("image/"))
      .slice(0, limit)
      .map(async (file) => ({
        name: file.name,
        type: file.type,
        size: file.size,
        data_url: await fileToDataUrl(file),
      }))
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
  const imageFiles = Array.from(files).filter((file) => file.type.startsWith("image/"));
  if (!imageFiles.length) return;
  const capacity = imageCapacity();
  if (!capacity) {
    setImageUploadStatus(`Maximum ${MAX_REFERENCE_IMAGES} reference images reached. Remove one before adding more.`);
    return;
  }
  const overflowCount = Math.max(0, imageFiles.length - capacity);
  const images = await filesToImageEntries(imageFiles, capacity);
  if (!images.length) return;
  const { unique, duplicateCount } = uniqueImageEntries(images);
  if (!unique.length) {
    const duplicateMessage = duplicateCount ? `${duplicateCount} duplicate image${duplicateCount === 1 ? "" : "s"} skipped.` : "";
    const overflowMessage = overflowCount ? `${overflowCount} image${overflowCount === 1 ? "" : "s"} skipped because the maximum is ${MAX_REFERENCE_IMAGES}.` : "";
    setImageUploadStatus([duplicateMessage, overflowMessage].filter(Boolean).join(" "));
    return;
  }
  state.images = [...state.images, ...unique];
  renderFiles();
  setWorkflowStage("ready_to_analyze");
  const generator = currentGenerator();
  const duplicateMessage = duplicateCount ? ` ${duplicateCount} duplicate image${duplicateCount === 1 ? "" : "s"} skipped.` : "";
  const overflowMessage = overflowCount ? ` ${overflowCount} image${overflowCount === 1 ? "" : "s"} skipped because the maximum is ${MAX_REFERENCE_IMAGES}.` : "";
  setImageUploadStatus(`${unique.length} new ${generator.imageNoun}${unique.length === 1 ? "" : "s"} added.${duplicateMessage}${overflowMessage}`);
  noteWorkflowEvent(
    "assistant",
    `${state.images.length} ${generator.imageNoun}${state.images.length === 1 ? "" : "s"} loaded. Click ${generator.analyzeLabel} when you want me to draft the quote basis.${duplicateMessage}${overflowMessage}`
  );
  syncControlStates();
}

function removeImageAt(index) {
  state.images.splice(index, 1);
  renderFiles();
  if (!state.images.length) {
    setWorkflowStage("needs_images");
    setImageUploadStatus("");
    noteWorkflowEvent("assistant", `No ${currentGenerator().imageNoun}s are loaded now. Drop references to start the quote analysis.`);
  } else {
    setImageUploadStatus(`${state.images.length} reference image${state.images.length === 1 ? "" : "s"} loaded.`);
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
    .map((image, index) => `
      <div class="file-item">
        <img class="file-thumb" src="${escapeHtml(image.data_url)}" alt="">
        <div>
          <strong>${escapeHtml(image.name)}</strong>
          <span>${escapeHtml(image.type || "image")} - ${formatBytes(image.size)}</span>
        </div>
        <button class="file-remove" type="button" data-remove-image="${index}" aria-label="Remove ${escapeHtml(image.name)}">x</button>
      </div>
    `)
    .join("");
}

function normalizeLineItem(item = {}) {
  const priceMode = item.price_mode === "Included" || String(item.display_price || "").toLowerCase() === "included"
    ? "Included"
    : "Priced";
  return {
    section: item.section || "",
    quantity: item.quantity ?? "",
    unit: item.unit || "",
    description: item.description || "",
    pricing_keyword: item.pricing_keyword || "",
    display_price: item.display_price || "",
    price_mode: priceMode,
    unit_price_override: item.unit_price_override ?? "",
    catalog_unit_price: item.catalog_unit_price ?? item.unit_price ?? item.sale_unit_price ?? "",
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

function clearSessionState() {
  try {
    window.localStorage.removeItem(QUOTE_SESSION_STORAGE_KEY);
  } catch {
    // Ignore storage failures; the local workflow can continue without refresh recovery.
  }
}

function saveSessionState() {
  if (state.isBooting) return;
  try {
    const snapshot = {
      version: QUOTE_SESSION_STATE_VERSION,
      savedAt: new Date().toISOString(),
      profileId: state.profileId,
      pricingReferenceId: state.pricingReferenceId,
      selectedPresetValue: state.selectedPresetValue,
      images: state.images.slice(0, MAX_REFERENCE_IMAGES),
      quoteDetails: collectQuoteDetails(),
      workflowStage: state.workflowStage,
      quoteBasis: state.quoteBasis,
      quoteBasisSections: state.quoteBasisSections,
      lineItems: state.lineItems,
      outputRows: state.outputRows,
      originalOutputRows: state.originalOutputRows,
      outputErrors: state.outputErrors,
      boothDimensions: state.boothDimensions,
      originalAnalysisSnapshot: state.originalAnalysisSnapshot,
      basisConfirmed: state.basisConfirmed,
      aiFailed: state.aiFailed,
      draftSource: state.draftSource,
      activeSidePanel: state.activeSidePanel,
      downloadFile: state.downloadFile,
      pricingMatches: state.pricingMatches,
      pricingIssues: state.pricingIssues,
      activeJob: state.activeJob,
    };
    window.localStorage.setItem(QUOTE_SESSION_STORAGE_KEY, JSON.stringify(snapshot));
  } catch {
    // Large image payloads can exceed storage quota; refresh recovery is best-effort.
  }
}

function restoreSessionState() {
  const saved = safeSessionJson();
  if (!saved || saved.version !== QUOTE_SESSION_STATE_VERSION) {
    clearSessionState();
    return false;
  }
  state.profileId = saved.profileId || "";
  state.pricingReferenceId = saved.pricingReferenceId || saved.profileId || "";
  state.selectedPresetValue = saved.selectedPresetValue || presetValueFromQuoteDetails(saved.quoteDetails || {});
  syncSelectedPricingReference();
  renderProfileOptions();
  renderPresetOptions();
  applyQuoteDetails(saved.quoteDetails || {}, { includeLogo: true, clearLogo: true });
  state.images = Array.isArray(saved.images) ? saved.images.slice(0, MAX_REFERENCE_IMAGES) : [];
  state.quoteBasis = cloneQuoteBasis(saved.quoteBasis || {});
  state.quoteBasisSections = normalizeQuoteBasisSections(saved.quoteBasisSections || saved.quoteBasis || {});
  state.lineItems = Array.isArray(saved.lineItems) ? saved.lineItems.map(normalizeLineItem) : [];
  state.outputRows = Array.isArray(saved.outputRows) ? saved.outputRows.map(normalizeOutputRow) : [];
  state.originalOutputRows = Array.isArray(saved.originalOutputRows) ? saved.originalOutputRows.map(normalizeOutputRow) : [];
  state.outputErrors = Array.isArray(saved.outputErrors) ? saved.outputErrors : [];
  state.boothDimensions = normalizeBoothDimensions(saved.boothDimensions || saved.quoteDetails?.project || {});
  state.originalAnalysisSnapshot = saved.originalAnalysisSnapshot || null;
  state.basisConfirmed = Boolean(saved.basisConfirmed);
  state.draftSource = saved.draftSource || "";
  state.aiFailed = Boolean(saved.aiFailed || state.draftSource === "local");
  state.downloadFile = saved.downloadFile || null;
  state.pricingMatches = Array.isArray(saved.pricingMatches) ? saved.pricingMatches : [];
  state.pricingIssues = Array.isArray(saved.pricingIssues) ? saved.pricingIssues : [];
  state.activeJob = saved.activeJob && saved.activeJob.id ? saved.activeJob : null;
  renderFiles();
  renderPricingMatches(state.outputRows.length ? state.outputRows : state.pricingMatches, { fromPricingMatches: !state.outputRows.length && state.pricingMatches.length });
  renderMatchSummary({ pricing_matches: state.pricingMatches });
  if (state.pricingIssues.length) {
    renderPricingReviewMessages({ errors: state.pricingIssues, pricing_matches: state.pricingMatches });
  } else {
    clearPricingReviewMessages();
  }
  if (state.aiFailed) {
    clearAiFailedDraftState();
    renderBasisFailureState("AI analysis failed before a usable quote basis was produced. Start analysis again after checking the local runner connection.");
  } else if (state.lineItems.length || state.quoteBasisSections.length || Object.values(state.quoteBasis).some((value) => splitLines(value).length > 0)) {
    updateQuoteBasisCard(saved.draftSource || "restored");
  } else {
    renderBasisEmptyState();
  }
  updateDownloadButton();
  setResultStatus(state.aiFailed ? "No usable AI draft" : state.downloadFile ? "Completed" : "No job yet", state.aiFailed ? "is-bad" : state.downloadFile ? "is-ok" : "");
  setWorkflowStage(saved.workflowStage || (state.images.length ? "ready_to_analyze" : "needs_images"));
  if (state.aiFailed) {
    showAiFailureBanner("Try again later. Regenerate analysis before confirming the quote basis.");
  } else {
    clearAiFailureBanner();
  }
  const restoredPanel = SIDE_PANEL_SEQUENCE.includes(saved.activeSidePanel) ? saved.activeSidePanel : "images";
  setSidePanel(restoredPanel, { force: true });
  return true;
}

function safeStorageJson() {
  try {
    const parsed = JSON.parse(window.localStorage.getItem(QUOTE_PRESETS_STORAGE_KEY) || "[]");
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

function savePresets(presets) {
  window.localStorage.setItem(QUOTE_PRESETS_STORAGE_KEY, JSON.stringify(presets));
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

function localPresetOptionValue(presetId) {
  return `${LOCAL_PRESET_PREFIX}${presetId}`;
}

function presetOptionValue(preset = {}) {
  const presetId = String(preset.id || "").trim();
  if (!presetId) return "";
  return preset.source === "local" ? localPresetOptionValue(presetId) : profilePresetOptionValue(presetId);
}

function defaultPresetOptionValue() {
  const profileDefault = defaultProfilePresetId();
  if (profileDefault) return profilePresetOptionValue(profileDefault);
  const localPreset = safeStorageJson()[0];
  return localPreset ? localPresetOptionValue(localPreset.id) : "";
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
  const localId = value.startsWith(LOCAL_PRESET_PREFIX) ? value.slice(LOCAL_PRESET_PREFIX.length) : value;
  const preset = safeStorageJson().find((item) => item.id === localId);
  return preset ? { ...preset, source: "local" } : null;
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
  const localPreset = safeStorageJson().find(matchesDetails);
  return localPreset ? localPresetOptionValue(localPreset.id) : "";
}

function renderPresetStatus(message = "") {
  if (!elements.presetStatus) return;
  elements.presetStatus.textContent = message || "Presets are stored locally in this browser and can include the uploaded header logo.";
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
  const presets = safeStorageJson();
  const selectedValue = state.selectedPresetValue || elements.presetSelect.value || "";
  const defaultPreset = builtInPresets.find((preset) => preset.id === "default");
  const defaultOption = defaultPreset
    ? `<option value="${escapeHtml(profilePresetOptionValue(defaultPreset.id))}">${escapeHtml(defaultPreset.name)}</option>`
    : "";
  const builtInOptions = builtInPresets
    .filter((preset) => preset.id !== "default")
    .map((preset) => `<option value="${escapeHtml(profilePresetOptionValue(preset.id))}">${escapeHtml(preset.name)}</option>`)
    .join("");
  const localOptions = presets
    .map((preset) => `<option value="${escapeHtml(localPresetOptionValue(preset.id))}">${escapeHtml(preset.name)}</option>`)
    .join("");
  elements.presetSelect.innerHTML = [
    defaultOption,
    builtInOptions ? `<optgroup label="Profile Presets">${builtInOptions}</optgroup>` : "",
    localOptions ? `<optgroup label="Saved Company Presets">${localOptions}</optgroup>` : "",
  ].join("");
  const availableValues = new Set([
    ...builtInPresets.map((preset) => profilePresetOptionValue(preset.id)),
    ...presets.map((preset) => localPresetOptionValue(preset.id)),
  ]);
  state.selectedPresetValue = availableValues.has(selectedValue) ? selectedValue : defaultPresetOptionValue();
  elements.presetSelect.value = state.selectedPresetValue;
  updatePresetButtons();
}

function updatePresetButtons() {
  const preset = selectedPreset();
  elements.loadPresetButton.disabled = !preset;
  elements.deletePresetButton.disabled = !preset || preset.source !== "local";
}

function saveCurrentPreset() {
  const name = elements.presetNameInput.value.trim();
  if (!name) {
    renderPresetStatus("Name the preset first, then save it.");
    elements.presetNameInput.focus();
    return;
  }
  const presets = safeStorageJson();
  const existingIndex = presets.findIndex((preset) => preset.name.toLowerCase() === name.toLowerCase());
  const preset = {
    id: existingIndex >= 0 ? presets[existingIndex].id : `preset-${Date.now().toString(36)}`,
    name,
    profile_id: currentProfile().id || DEFAULT_PROFILE_ID,
    details: collectQuoteDetails(),
    saved_at: new Date().toISOString(),
  };
  if (existingIndex >= 0) {
    presets[existingIndex] = preset;
  } else {
    presets.push(preset);
  }
  savePresets(presets);
  renderPresetOptions();
  state.selectedPresetValue = localPresetOptionValue(preset.id);
  elements.presetSelect.value = state.selectedPresetValue;
  updatePresetButtons();
  syncControlStates();
  renderPresetStatus(`Saved "${name}" as a local company preset.`);
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
  if (!options.silent) {
    noteWorkflowEvent("assistant", `Loaded company preset "${preset.name}". Pricing reference and reference images are unchanged.`);
  }
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
  noteWorkflowEvent("assistant", "Customer details cleared. The selected pricing reference, images, and quote-company defaults were left unchanged.");
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
  noteWorkflowEvent("assistant", "Quote-company defaults reset. Customer details, pricing reference, and images were left unchanged.");
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
  noteWorkflowEvent("assistant", "Image draft reset. Customer, Quote Company, and selected Pricing Reference were kept.");
}

function startNewQuote() {
  if (state.isAnalysisRunning || state.isGenerating) return;
  clearSessionState();
  state.profileId = "";
  state.pricingReferenceId = "";
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
  const value = selectedPresetId();
  const presetId = value.startsWith(LOCAL_PRESET_PREFIX) ? value.slice(LOCAL_PRESET_PREFIX.length) : value;
  const presets = safeStorageJson();
  const preset = presets.find((item) => item.id === presetId);
  if (!preset) {
    renderPresetStatus("Choose a saved preset to delete.");
    return;
  }
  savePresets(presets.filter((item) => item.id !== presetId));
  renderPresetOptions();
  updatePresetButtons();
  syncControlStates();
  renderPresetStatus(`Deleted "${preset.name}".`);
}

function renderProfileOptions() {
  if (!elements.profileSelect) return;
  const references = state.pricingReferences.length ? state.pricingReferences : [];
  const selectedValue = currentPricingReference() ? pricingReferenceSelectValue(currentPricingReference()) : "";
  const referenceOption = (reference) => {
    const referenceId = String(reference.id || "").trim();
    return `<option value="${escapeHtml(pricingReferenceSelectValue(reference))}">${escapeHtml(reference.label || referenceId)}</option>`;
  };
  const profileOptions = references
    .filter((reference) => reference.source !== "local")
    .map(referenceOption)
    .join("");
  const localOptions = references
    .filter((reference) => reference.source === "local")
    .map(referenceOption)
    .join("");
  elements.profileSelect.innerHTML = [
    profileOptions,
    localOptions ? `<optgroup label="Saved Pricing References">${localOptions}</optgroup>` : "",
  ].join("");
  const fallbackReference = currentPricingReference() || defaultPricingReference() || references[0] || null;
  if (fallbackReference) {
    state.pricingReferenceId = fallbackReference.id || "";
  }
  const selectedReference = currentPricingReference();
  elements.profileSelect.value = selectedReference ? pricingReferenceSelectValue(selectedReference) : selectedValue;
  if (elements.deletePricingReferenceButton) {
    elements.deletePricingReferenceButton.disabled = !selectedReference || selectedReference.source !== "local";
  }
}

function pricingReferenceRequiredColumns() {
  return ["id", "section", "description", "unit_hint", "internal_cost", "markup_multiplier"];
}

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
    const text = String(line.text || line.line || line.description || "").trim();
    if (!text) return [];
    const confidence = normalizeConfidence(line.confidence ?? line.confidence_pct);
    const hasCustomPricing = isCustomPricingBasisLine(line);
    return splitBasisDecisionText(text, line.tag).map((parsed) => {
      const next = { tag: normalizeBasisTag(parsed.tag), text: parsed.text };
      if (confidence !== null) next.confidence = confidence;
      if (hasCustomPricing || normalizeBasisTag(parsed.tag) === "Custom") next.custom_pricing = true;
      return next;
    });
  }
  return splitBasisDecisionText(line);
}

function parseBasisLine(line = "") {
  return normalizeBasisLines(line)[0] || { tag: "Confirm", text: "" };
}

function basisLineMeta(line = "") {
  return parseBasisLine(line);
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
        const title = basisDisplayTitle(section?.title || "Section") || "Section";
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
    section: String(row.section || ""),
    description: String(row.description || ""),
    quantity: row.quantity ?? "",
    unit: normalizeUnit(row.unit || ""),
    price_mode: priceMode,
    unit_price_override: row.unit_price_override ?? "",
    catalog_unit_price: row.catalog_unit_price ?? row.unit_price ?? row.sale_unit_price ?? "",
    pricing_keyword: row.pricing_keyword || row.keyword || "",
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
    items,
    rowCount: items.length,
    headers,
    missing,
    skipped,
    layout: "normalized-pricing-reference",
    errors,
    warnings,
    exampleRows: 0,
    canSave: !errors.length && items.length > 0,
  };
}

function renderPricingReferencePreview(result = null) {
  if (!elements.pricingReferencePreview) return;
  if (!result) {
    elements.pricingReferencePreview.innerHTML = "";
    if (elements.pricingReferenceSaveButton) elements.pricingReferenceSaveButton.disabled = true;
    return;
  }
  const required = pricingReferenceRequiredColumns();
  const found = required.filter((column) => !result.missing.includes(column));
  const canSave = result.canSave ?? (!result.errors.length && result.rowCount > 0);
  const tone = result.errors.length ? "error" : result.warnings.length || !canSave ? "warn" : "ok";
  const title = result.errors.length ? "Validation failed" : canSave ? "Validation preview" : "Template preview";
  elements.pricingReferencePreview.className = `pricing-reference-preview ${tone}`;
  elements.pricingReferencePreview.innerHTML = `
    <strong>${title}</strong>
    <ul>
      <li>Layout: ${escapeHtml(result.layout || "normalized pricing reference")}</li>
      <li>Detected row count: ${result.rowCount}</li>
      ${result.exampleRows ? `<li>Example rows ignored: ${result.exampleRows}</li>` : ""}
      <li>Required columns found: ${escapeHtml(found.join(", ") || "None")}</li>
      <li>Required columns missing: ${escapeHtml(result.missing.join(", ") || "None")}</li>
      <li>Skipped rows count: ${result.skipped}</li>
    </ul>
    ${result.errors.concat(result.warnings).map((message) => `<p>${escapeHtml(message)}</p>`).join("")}
  `;
  if (elements.pricingReferenceSaveButton) elements.pricingReferenceSaveButton.disabled = !canSave;
}

async function validatePricingReferenceFile(file) {
  if (!file) return pricingReferenceValidationResult([], [], 0, "");
  if (file.size > 2 * 1024 * 1024) {
    return { ...pricingReferenceValidationResult([], [], 0, file.name), errors: ["Pricing reference file is larger than 2 MB."] };
  }
  const extension = file.name.split(".").pop().toLowerCase();
  if (!["csv", "xlsx"].includes(extension)) {
    return { ...pricingReferenceValidationResult([], [], 0, file.name), errors: ["Upload a .xlsx or .csv pricing reference template."] };
  }
  const { ok, data } = await postJson("/api/pricing-reference/validate", {
    filename: file.name,
    data_url: await fileToDataUrl(file),
  });
  if (ok) return data;
  return {
    ...pricingReferenceValidationResult([], [], 0, file.name),
    errors: data.errors || ["Pricing-reference validation failed."],
  };
}

async function downloadPricingReferenceTemplate(event) {
  event.preventDefault();
  try {
    const response = await fetch("/api/pricing-reference/template.xlsx", { cache: "no-store" });
    if (!response.ok) {
      renderPricingReferencePreview({
        ...pricingReferenceValidationResult([], [], 0, "swooshz-pricing-reference-template.xlsx"),
        errors: [`Template download failed with server status ${response.status}.`],
      });
      return;
    }
    const blob = await response.blob();
    const url = window.URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = "swooshz-pricing-reference-template.xlsx";
    document.body.appendChild(link);
    link.click();
    link.remove();
    window.URL.revokeObjectURL(url);
  } catch (error) {
    renderPricingReferencePreview({
      ...pricingReferenceValidationResult([], [], 0, "swooshz-pricing-reference-template.xlsx"),
      errors: [`Template download failed: ${error.message || String(error)}`],
    });
  }
}

function openPricingReferenceModal() {
  state.pendingPricingReference = null;
  if (elements.pricingReferenceName) elements.pricingReferenceName.value = "";
  if (elements.pricingReferenceFile) elements.pricingReferenceFile.value = "";
  if (elements.pricingReferenceFileName) elements.pricingReferenceFileName.textContent = "No file chosen";
  renderPricingReferencePreview(null);
  elements.pricingReferenceModal.hidden = false;
  elements.pricingReferenceModal.classList.add("is-open");
  window.setTimeout(() => elements.pricingReferenceName?.focus(), 0);
}

function closePricingReferenceModal() {
  elements.pricingReferenceModal.classList.remove("is-open");
  elements.pricingReferenceModal.hidden = true;
  state.pendingPricingReference = null;
}

function savePricingReferenceFromModal(event) {
  event.preventDefault();
  const name = elements.pricingReferenceName.value.trim();
  const result = state.pendingPricingReference;
  const canSave = result?.canSave ?? Boolean(result && !result.errors.length && result.items.length);
  if (!name || !result || !canSave) {
    renderPricingReferencePreview(result || pricingReferenceValidationResult([], [], 0, ""));
    return;
  }
  const reference = {
    id: `local-ref-${Date.now().toString(36)}`,
    label: name,
    source: "local",
    schema_version: 1,
    items: result.items,
    saved_at: new Date().toISOString(),
  };
  const localReferences = safeLocalPricingReferences().filter((item) => item.label.toLowerCase() !== name.toLowerCase());
  localReferences.push(reference);
  saveLocalPricingReferences(localReferences);
  state.pricingReferences = mergePricingReferences(state.pricingReferences.filter((item) => item.source !== "local"));
  state.pricingReferenceId = reference.id;
  syncSelectedPricingReference();
  renderProfileOptions();
  clearGeneratedQuoteState();
  closePricingReferenceModal();
  noteWorkflowEvent("assistant", `Saved and selected pricing reference "${name}". Previous analysis/output state was cleared.`);
  syncControlStates();
}

function deleteSelectedPricingReference() {
  const selected = currentPricingReference();
  if (!selected || selected.source !== "local") return;
  const localReferences = safeLocalPricingReferences().filter((item) => item.id !== selected.id);
  saveLocalPricingReferences(localReferences);
  state.pricingReferences = mergePricingReferences(state.pricingReferences.filter((item) => item.source !== "local"));
  state.pricingReferenceId = "";
  syncSelectedPricingReference();
  renderProfileOptions();
  clearGeneratedQuoteState();
  noteWorkflowEvent("assistant", `Deleted local pricing reference "${selected.label || selected.id}".`);
  syncControlStates();
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
  state.basisConfirmed = false;
  state.aiFailed = false;
  state.draftSource = "";
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
  if (nextSelection.pricingReferenceId === state.pricingReferenceId) {
    return;
  }
  state.pricingReferenceId = nextSelection.pricingReferenceId;
  syncSelectedPricingReference();
  renderProfileOptions();
  clearGeneratedQuoteState();
  setWorkflowStage(state.images.length ? (canStartAnalysis() ? "ready_to_analyze" : "details_review") : "needs_images");
  noteWorkflowEvent("assistant", "Pricing reference changed. I cleared the previous draft so the next analysis uses the selected pricing context. Customer and quote-company details were left unchanged.");
  syncControlStates();
}

async function setSampleDetails() {
  if (state.isBooting || state.isAnalysisRunning || state.isGenerating) return;
  elements.sampleDetailsButton.disabled = true;
  try {
    const { ok, data } = await getJson(`/api/samples/${DEFAULT_SAMPLE_ID}`);
    if (!ok) {
      noteWorkflowEvent("assistant", (data.errors || [data.error || "Sample fixture could not be loaded."]).join("\n"), { tone: "error" });
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
    renderFiles();
    setImageUploadStatus(`${state.images.length} sample reference image${state.images.length === 1 ? "" : "s"} loaded.`);
    setWorkflowStage(state.images.length ? "ready_to_analyze" : "needs_images");
    noteWorkflowEvent("assistant", `${data.label || "Sample"} loaded with ${state.images.length} reference image${state.images.length === 1 ? "" : "s"}.`);
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
    pricing_reference: pricingReference?.source === "local" ? pricingReference : null,
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
    user_feedback: state.pendingFeedback,
    quote_basis: includeDraftContext ? { ...state.quoteBasis, ...quoteBasisFromSections(state.quoteBasisSections) } : {},
    quote_basis_sections: includeDraftContext ? cloneQuoteBasisSections(state.quoteBasisSections) : [],
    line_items: includeDraftContext ? (state.outputRows.length ? outputRowsToLineItems(state.outputRows) : state.lineItems) : [],
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
  const enabled = state.activeSidePanel === "output" && (validation.valid || Boolean(file && file.url)) && !state.isGenerating;
  elements.sideDownloadButton.classList.toggle("is-disabled", !enabled);
  elements.sideDownloadButton.setAttribute("aria-disabled", String(!enabled));
  elements.sideDownloadButton.tabIndex = enabled ? 0 : -1;
  elements.sideDownloadButton.href = enabled && file?.url ? file.url : "#";
  elements.sideDownloadButton.download = file?.url ? file.name || "quotation.xlsx" : "";
  elements.sideDownloadButton.textContent = "Download Excel";
}

function pricingMatchStatus(row = {}) {
  const status = typeof row === "string" ? row : row.status;
  return String(status || "").trim().toLowerCase();
}

function pricingRowNeedsReview(row = {}) {
  const status = pricingMatchStatus(row);
  const amount = String(row.amount || "").trim();
  if (["unmatched", "ambiguous", "matched-from-ambiguous", "custom", "needs-pricing"].includes(status)) return true;
  return status === "manual-display" && (!amount || amount.toLowerCase() === "manual display price");
}

function pricingIssueForRow(row = {}) {
  if (!pricingRowNeedsReview(row)) return "";
  const description = row.description || "Pricing row";
  const status = pricingMatchStatus(row);
  if (status === "manual-display") return `Manual display pricing required: ${description} / enter a display price, choose a catalog keyword, or remove this line.`;
  if (status === "custom" || status === "needs-pricing") return `Manual pricing required: ${description} / enter a unit price, choose a catalog keyword, or remove this line.`;
  if (status === "matched-from-ambiguous") return `Ambiguous pricing: ${description} / confirm selected catalog match.`;
  if (status === "ambiguous") return `Ambiguous pricing: ${description} / choose a catalog match.`;
  return `Unmatched pricing: ${description} / choose a catalog match or remove from quote.`;
}

function pricingReviewIssues(result = {}) {
  const errorIssues = extractPricingIssues(result.errors || []);
  const rowIssues = (result.pricing_matches || []).map(pricingIssueForRow).filter(Boolean);
  return [...new Set([...errorIssues, ...rowIssues])];
}

function pricingReviewBlockReason() {
  const count = state.pricingIssues.length || state.pricingMatches.filter(pricingRowNeedsReview).length;
  return count ? `Resolve ${count} pricing item${count === 1 ? "" : "s"} before opening Output.` : "";
}

function pricingStatusLabel(status = "") {
  const labels = {
    matched: "Catalog match",
    "matched-from-ambiguous": "Ambiguous match selected",
    ambiguous: "Ambiguous match",
    "manual-display": "Manual display price",
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
  const hasUsablePrice = unitPrice !== null && unitPrice > 0;
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

function outputRowCoversBasisLine(row = {}, lineText = "") {
  const rowText = outputComparableText(row.description);
  const basisText = outputComparableText(lineText);
  if (!rowText || !basisText) return false;
  if (rowText.includes(basisText) || basisText.includes(rowText)) return true;
  const basisWords = basisText.split(" ").filter((word) => word.length > 3);
  if (basisWords.length < 4) return false;
  const rowWords = new Set(rowText.split(" "));
  const overlap = basisWords.filter((word) => rowWords.has(word)).length;
  return overlap / Math.min(basisWords.length, 10) >= 0.6;
}

function includedBasisOutputRows(existingRows = []) {
  const rows = [];
  basisSections(state.quoteBasisSections).forEach((section) => {
    (section.lines || []).forEach((line) => {
      const basisTag = normalizeBasisTag(line.tag);
      const customPricing = isCustomPricingBasisLine(line);
      if (basisTag === "Exclude") return;
      if (!["Include", "Custom"].includes(basisTag) && !customPricing) return;
      const text = String(line.text || "").trim();
      if (!text) return;
      const duplicate = existingRows.some((row) => outputRowCoversBasisLine(row, text))
        || rows.some((row) => outputRowCoversBasisLine(row, text));
      if (duplicate) return;
      rows.push(normalizeOutputRow({
        section: basisDisplayTitle(section.title || section.id || "Quote Basis"),
        description: text,
        quantity: "",
        unit: "",
        price_mode: "Priced",
        unit_price_override: "",
        pricing_keyword: "",
        status: customPricing ? "custom" : "needs-pricing",
      }));
    });
  });
  return rows;
}

function outputRowFromLineItem(item = {}) {
  const normalized = normalizeLineItem(item);
  return normalizeOutputRow({
    section: normalized.section,
    description: normalized.description,
    quantity: normalized.quantity,
    unit: normalized.unit,
    price_mode: normalized.price_mode,
    unit_price_override: normalized.unit_price_override || "",
    catalog_unit_price: normalized.catalog_unit_price || "",
    pricing_keyword: normalized.pricing_keyword,
  });
}

function outputRowFromPricingMatch(row = {}) {
  const status = pricingMatchStatus(row);
  const unitPrice = row.unit_price || row.unit_price_override || "";
  const amount = String(row.amount || "").trim();
  return normalizeOutputRow({
    section: row.section,
    description: row.catalog_description || row.description,
    quantity: String(row.quantity || "").replace(/\s+[A-Za-z]+$/, ""),
    unit: row.unit || String(row.quantity || "").match(/\s+([A-Za-z]+)$/)?.[1] || "",
    price_mode: status === "included" ? "Included" : "Priced",
    unit_price_override: status === "manual-price" ? unitPrice : "",
    catalog_unit_price: status === "manual-price" ? "" : unitPrice,
    pricing_keyword: row.keyword || row.pricing_keyword || "",
    status: row.status,
  });
}

function outputCellDisplayValue(row = {}, field = "") {
  if (field === "price_mode") return row.price_mode === "Included" ? "Included" : "Priced";
  if (field === "unit_price_override") {
    if (row.price_mode === "Included") return "Included";
    if (numberOrNull(row.unit_price_override) !== null) return formatAmount(row.unit_price_override);
    if (numberOrNull(row.catalog_unit_price) !== null) return formatAmount(row.catalog_unit_price);
    return "Pending";
  }
  if (field === "amount") {
    if (row.price_mode === "Included") return "0.00";
    return formatAmount(row.amount) || "Pending";
  }
  return String(row[field] ?? "").trim() || "Pending";
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
    const listId = `unitPriceOptions-${index}`;
    return `
      <span class="output-unit-price-editor">
        <input class="output-cell-input is-editing" data-output-editor-field="${field}" data-output-row="${index}" value="${escapeHtml(value)}" inputmode="decimal" list="${listId}">
        <button class="output-included-button" type="button" data-output-included-action="true" data-output-row="${index}">Included</button>
      </span>
      <datalist id="${listId}">
        <option value="Included"></option>
      </datalist>
    `;
  }
  const inputMode = field === "quantity" ? "decimal" : "text";
  return `<input class="output-cell-input is-editing" data-output-editor-field="${field}" data-output-row="${index}" value="${escapeHtml(value)}" inputmode="${inputMode}">`;
}

function renderOutputEditCell(row = {}, index = 0, field = "", extraClass = "") {
  const display = outputCellDisplayValue(row, field);
  const pending = display === "Pending";
  return `
    <td class="output-edit-cell ${extraClass} ${pending ? "is-pending" : ""}" data-output-edit-field="${field}" data-output-row="${index}" tabindex="0" title="Click to edit">
      <span class="output-cell-text">${escapeHtml(display)}</span>
    </td>
  `;
}

function ensureOutputRowsFromLineItems() {
  if (state.outputRows.length) return;
  const generatedRows = state.lineItems.map(outputRowFromLineItem);
  state.outputRows = [...generatedRows, ...includedBasisOutputRows(generatedRows)];
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
    };
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
    if (row.price_mode !== "Included" && (quantity === null || quantity <= 0)) {
      errors.push(`${label}: Quantity must be greater than 0.`);
    }
    if (row.price_mode !== "Included") {
      const unitPrice = numberOrNull(row.unit_price_override);
      const catalogUnitPrice = numberOrNull(row.catalog_unit_price);
      const unitPriceKind = unitPriceEditKind(row.unit_price_override);
      const hasPricingKeyword = Boolean(String(row.pricing_keyword || "").trim());
      if (unitPriceKind === "invalid") {
        errors.push(`${label}: Unit price must be a number or Included.`);
      } else if (unitPrice !== null && unitPrice <= 0) {
        errors.push(`${label}: Unit price must be greater than 0.`);
      } else if (!hasPricingKeyword && unitPrice === null && catalogUnitPrice === null) {
        errors.push(`${label}: Unit price or pricing keyword is required.`);
      }
    }
  });
  return { valid: errors.length === 0 && rows.length > 0, errors };
}

function renderOutputValidationMessages(errors = state.outputErrors) {
  if (!elements.pricingReviewMessages) return;
  state.outputErrors = errors;
  elements.pricingReviewMessages.innerHTML = "";
}

function matchSummaryStats(rows = []) {
  const safeRows = Array.isArray(rows) ? rows : [];
  const hasResolvedStatuses = safeRows.some((row) => pricingMatchStatus(row));
  const total = safeRows.reduce((sum, row) => {
    const amount = Number(String(row.amount || "").replaceAll(",", ""));
    return Number.isFinite(amount) ? sum + amount : sum;
  }, 0);
  const includedRows = safeRows.filter((row) => row.price_mode === "Included" || pricingMatchStatus(row) === "included").length;
  const pricedRows = safeRows.length - includedRows;
  const confident = safeRows.filter((row) => pricingMatchStatus(row) === "matched").length;
  const needsReview = safeRows.filter((row) => {
    const status = pricingMatchStatus(row);
    if (status) return status !== "matched" && status !== "included" && status !== "manual-price";
    if (row.price_mode === "Included") return false;
    const quantity = numberOrNull(row.quantity);
    const unitPrice = effectiveOutputUnitPrice(row);
    return !String(row.pricing_keyword || "").trim() && (quantity === null || quantity <= 0 || unitPrice === null || unitPrice <= 0);
  }).length;
  const pending = safeRows.filter((row) => !pricingMatchStatus(row) && row.price_mode !== "Included").length;
  const totalPending = safeRows.some((row) => row.price_mode !== "Included" && (row.amount === "" || row.amount === null || row.amount === undefined));
  const confidence = hasResolvedStatuses && safeRows.length > 0 ? Math.round((confident / safeRows.length) * 100) : null;
  return { total, confident, needsReview, needsPrice: needsReview, confidence, pending, totalPending, pricedRows, includedRows };
}

function renderMatchSummary(result = {}) {
  const rows = result.pricing_matches || [];
  if (!rows.length) {
    elements.matchSummary.innerHTML = "";
    return;
  }
  const stats = matchSummaryStats(rows);
  const totalValue = stats.totalPending
    ? "Pending"
    : `SGD ${stats.total.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
  elements.matchSummary.innerHTML = `
    <div class="stat-card-row">
      <div class="stat-card">
        <span class="stat-card-icon green" aria-hidden="true">#</span>
        <span class="stat-card-value">${stats.pricedRows}</span>
        <span class="stat-card-label">Priced rows</span>
      </div>
      <div class="stat-card">
        <span class="stat-card-icon blue" aria-hidden="true">In</span>
        <span class="stat-card-value">${stats.includedRows}</span>
        <span class="stat-card-label">Included rows</span>
      </div>
      <div class="stat-card">
        <span class="stat-card-icon red" aria-hidden="true">!</span>
        <span class="stat-card-value">${stats.needsPrice}</span>
        <span class="stat-card-label">Needs price</span>
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
  const outputRows = state.outputRows;
  elements.pricingTableWrap.hidden = !outputRows.length;
  elements.pricingEmptyState.hidden = Boolean(outputRows.length) || Boolean(elements.pricingReviewMessages.innerHTML.trim());
  if (!outputRows.length) {
    elements.pricingMatchesBody.innerHTML = `<tr><td colspan="6">No output rows yet.</td></tr>`;
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
        <td class="amount-cell ${outputCellDisplayValue(row, "amount") === "Pending" ? "is-pending" : ""}">${escapeHtml(outputCellDisplayValue(row, "amount"))}</td>
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

function extractPricingIssues(errors = []) {
  return errors
    .map((error) => String(error || "").trim())
    .filter((error) => /^(Unmatched pricing:|Ambiguous pricing:|Manual display pricing required:)/i.test(error))
    .map((error) => error.replace(/^-?\s*/, ""));
}

function pricingIssueDescription(issue = "") {
  return String(issue)
    .replace(/^Unmatched pricing:\s*/i, "")
    .replace(/^Ambiguous pricing:\s*/i, "")
    .replace(/^Manual display pricing required:\s*/i, "")
    .split(" / keyword ")[0]
    .split(" / enter a display price")[0]
    .split(". Candidate matches:")[0]
    .trim();
}

function findLineItemIndexForPricingIssue(issue = "") {
  const description = pricingIssueDescription(issue).toLowerCase();
  return state.lineItems.findIndex((item) => {
    const itemDescription = String(item.description || "").toLowerCase();
    return itemDescription === description || itemDescription.includes(description) || description.includes(itemDescription);
  });
}

async function handlePricingChoice(action, issue) {
  if (state.isGenerating) return;
  const index = findLineItemIndexForPricingIssue(issue);
  if (index < 0) {
    noteWorkflowEvent("assistant", "I could not map that pricing issue back to a generated line item. Regenerate analysis and try again.", { tone: "warn" });
    return;
  }

  const item = state.lineItems[index];
  if (action === "remove_line") {
    const [removed] = state.lineItems.splice(index, 1);
    noteWorkflowEvent("assistant", `Removed "${removed.description}" from the quotation. Checking pricing again.`);
    await handleGenerate();
    return;
  }

  if (action === "manual_price") {
    const manualPrice = window.prompt("Manual display price", item.display_price || "");
    if (!manualPrice || !manualPrice.trim()) return;
    item.display_price = manualPrice.trim();
    noteWorkflowEvent("assistant", `Set a manual display price for "${item.description}". Checking pricing again.`);
    await handleGenerate();
    return;
  }

  if (action === "nearest_keyword") {
    const keyword = window.prompt("Pricing keyword to try", item.pricing_keyword || item.description);
    if (!keyword || !keyword.trim()) return;
    item.pricing_keyword = keyword.trim();
    item.display_price = "";
    noteWorkflowEvent("assistant", `Updated the pricing keyword for "${item.description}". Checking pricing again.`);
    await handleGenerate();
  }
}

function renderPricingReviewMessages(result = {}) {
  const issues = pricingReviewIssues(result);
  state.pricingIssues = issues;
  if (!issues.length) {
    elements.pricingReviewMessages.innerHTML = (result.errors || ["Pricing needs review."])
      .map((message) => `<div class="message warn">${escapeHtml(message)}</div>`)
      .join("");
    elements.pricingEmptyState.hidden = true;
    return;
  }

  elements.pricingReviewMessages.innerHTML = `
    <div class="message warn">
      <strong>I could not confidently price ${issues.length} item${issues.length === 1 ? "" : "s"}.</strong>
      Review these choices before treating the quotation as final.
    </div>
    <div class="pricing-review-list">
      ${issues.map((issue, index) => `
        <div class="pricing-review-card">
          <strong>${escapeHtml(issue)}</strong>
          <div class="pricing-review-actions">
            <button class="secondary-button" type="button" data-pricing-action="nearest_keyword" data-pricing-issue="${escapeHtml(issue)}">Use nearest match</button>
            <button class="secondary-button" type="button" data-pricing-action="manual_price" data-pricing-issue="${escapeHtml(issue)}">Manual display price</button>
            <button class="secondary-button" type="button" data-pricing-action="remove_line" data-pricing-issue="${escapeHtml(issue)}">Remove from quote</button>
          </div>
        </div>
      `).join("")}
    </div>
  `;
  elements.pricingEmptyState.hidden = true;
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
  const index = Number(editor.dataset.outputRow);
  const field = editor.dataset.outputEditorField;
  if (!Number.isInteger(index) || index < 0 || !state.outputRows[index] || !field) return;
  const currentRow = state.outputRows[index];
  let nextRow = { ...currentRow, [field]: editor.value };
  if (field === "unit_price_override") {
    const value = String(editor.value || "").trim();
    if (value.toLowerCase() === "included") {
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

function noteWorkflowEvent() {}

function renderBasisEmptyState(message = "Load images, complete Customer and Quote Company, then start analysis to review the draft here.") {
  elements.basisReviewSurface.innerHTML = `
    <div class="basis-empty-state">
      <strong>Quotation basis draft</strong>
      <p>${escapeHtml(message)}</p>
    </div>
  `;
}

function renderBasisFailureState(message = "AI analysis failed before a usable quote basis was produced. Check the local runner connection, then start analysis again.") {
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

function renderCurrentActions() {}

function basisSections(sections = state.quoteBasisSections) {
  const normalized = normalizeQuoteBasisSections(sections.length ? sections : state.quoteBasis);
  return normalized.length
    ? normalized
    : [{ id: "draft", title: "Quote Basis", lines: [{ tag: "Confirm", text: "No detail generated yet." }] }];
}

function unresolvedConfirmLines(sections = state.quoteBasisSections) {
  return basisSections(sections)
    .flatMap((section) => (section.lines || [])
      .filter((line) => normalizeBasisTag(line.tag) === "Confirm")
      .map((line) => `${section.title}: ${line.text}`));
}

function basisConfirmBlockReason(sections = state.quoteBasisSections) {
  return unresolvedConfirmLines(sections).length
    ? "Resolve all review lines before confirming quotation basis."
    : "";
}

function basisTagLabel(tag = "") {
  return normalizeBasisTag(tag);
}

function basisLinePillLabel(line = {}) {
  const tag = normalizeBasisTag(line.tag);
  const confidence = normalizeConfidence(line.confidence ?? line.confidence_pct);
  if (tag === "Confirm" && confidence !== null) return `${confidence}%`;
  if (tag === "Confirm") return "Check";
  return basisTagLabel(tag);
}

function renderBasisLine(section, line, index) {
  const tag = normalizeBasisTag(line.tag);
  const customPricing = isCustomPricingBasisLine(line);
  const rowClasses = [
    "basis-line-row",
    `basis-line-${tag.toLowerCase()}`,
    customPricing ? "basis-line-custom-priced" : "",
  ].filter(Boolean).join(" ");
  const primaryAction = customPricing
    ? `<button class="basis-line-tag-button" type="button" data-basis-section="${escapeHtml(section.id)}" data-basis-line-index="${index}" data-basis-tag="Custom" aria-label="Keep this line as custom manual pricing" title="Custom manual pricing">$</button>`
    : `<button class="basis-line-tag-button" type="button" data-basis-section="${escapeHtml(section.id)}" data-basis-line-index="${index}" data-basis-tag="Include" aria-label="Mark this line as included" title="Mark included">&#x2713;</button>`;
  return `
    <li class="${escapeHtml(rowClasses)}">
      <span class="basis-line-icon" aria-hidden="true"></span>
      <span class="basis-line-pill" title="${tag === "Confirm" ? "AI confidence before review" : escapeHtml(basisTagLabel(tag))}">${escapeHtml(basisLinePillLabel(line))}</span>
      <span class="basis-line-text" title="${escapeHtml(line.text)}">${escapeHtml(line.text)}</span>
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
      ${BASIS_TAGS.map(([tag, label, description]) => `
        <span class="basis-tag-legend-item basis-line-${escapeHtml(tag.toLowerCase())}">
          <strong class="basis-line-pill">${escapeHtml(label)}</strong>
          <span>${escapeHtml(description)}</span>
        </span>
      `).join("")}
    </div>
  `;
}

function renderQuoteBasisMessage(basis = state.quoteBasis, source = "") {
  const aiFailed = source === "local" && state.aiFailed;
  if (aiFailed) {
    return `
      <div class="basis-empty-state basis-empty-state-error">
        <strong>AI analysis did not complete</strong>
        <p>Remote AI did not return a usable quote basis. Start analysis again after checking the local runner connection.</p>
      </div>
    `;
  }
  const statusText = aiFailed ? "AI failed" : source === "edited" ? "Edited draft" : "Needs review";
  const summaryText = state.lineItems.length
    ? `${state.lineItems.length} priced line${state.lineItems.length === 1 ? "" : "s"}`
    : "Pricing draft pending";
  const sections = basisSections(state.quoteBasisSections.length ? state.quoteBasisSections : normalizeQuoteBasisSections(basis));
  return `
    <div class="assistant-card quote-basis-card ${aiFailed ? "quote-basis-card-failed" : ""}">
      <div class="quote-basis-header">
        <div>
          <div class="quote-basis-title-row">
            <h3>Quote Basis</h3>
            <span>${escapeHtml(statusText)}</span>
          </div>
          <p>${aiFailed ? "AI analysis failed. Try again later. A local starter draft is shown for reference only." : "Please review the AI takeoff, revise a line, or request changes."}</p>
        </div>
        <div class="quote-basis-source">
          <span>${aiFailed ? "Source: Local fallback only" : "Source: Koncept Pricing Catalog"}</span>
          <strong>${escapeHtml(summaryText)}</strong>
        </div>
      </div>
      ${renderBasisTagLegend()}
      <div class="basis-review-grid">
        ${sections.map((section) => `
          <div class="basis-review-item">
            <div class="basis-section-heading">
              <h4>${escapeHtml(section.title)}</h4>
              <span class="basis-section-actions">
                <button class="basis-section-action-button" type="button" data-basis-section="${escapeHtml(section.id)}" data-basis-section-action="Include" aria-label="Mark all non-custom review lines in ${escapeHtml(section.title)} as included" title="Mark all non-custom lines included">&#x2713;</button>
                <button class="basis-section-action-button" type="button" data-basis-section="${escapeHtml(section.id)}" data-basis-section-action="Exclude" aria-label="Mark all non-custom review lines in ${escapeHtml(section.title)} as excluded" title="Mark all non-custom lines excluded">X</button>
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

function currentQuoteBasisCard() {
  const cards = Array.from(elements.basisReviewSurface.querySelectorAll(".quote-basis-card"));
  return cards.at(-1) || null;
}

function updateQuoteBasisCard(source = "edited") {
  elements.basisReviewSurface.innerHTML = renderQuoteBasisMessage(state.quoteBasis, source);
}

function retagBasisLine(sectionId, lineIndex, nextTag) {
  if (!sectionId || !["Include", "Custom", "Exclude"].includes(nextTag)) return;
  const sections = cloneQuoteBasisSections(state.quoteBasisSections);
  const section = sections.find((item) => item.id === sectionId);
  const index = Number(lineIndex);
  if (!section || !section.lines[index]) return;
  const currentLine = section.lines[index];
  const customPricing = isCustomPricingBasisLine(currentLine) || nextTag === "Custom";
  section.lines[index] = {
    ...currentLine,
    tag: nextTag,
    ...(customPricing ? { custom_pricing: true } : {}),
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
    const currentTag = normalizeBasisTag(line.tag);
    if (currentTag === "Custom" || currentTag === nextTag) return line;
    changed = true;
    return { ...line, tag: nextTag };
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
  if (state.basisChat.scope === "line") {
    return "Tell me what to change. I will draft the full replacement sentence for approval before applying it.";
  }
  return "Ask a question or describe changes to the quotation basis. I will show a proposed update before applying anything.";
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
  const current = JSON.stringify(cloneQuoteBasisSections(state.quoteBasisSections));
  const next = JSON.stringify(cloneQuoteBasisSections(nextSections));
  if (current === next) return [];
  return nextSections.map((section) => section.title);
}

function proposalLinePreview(proposal) {
  if (state.basisChat.scope !== "line" || !state.basisChat.sectionId) return "";
  const nextSections = normalizeQuoteBasisSections(proposal?.quoteBasisSections || proposal?.quoteBasis || {});
  const section = nextSections.find((item) => item.id === state.basisChat.sectionId);
  const nextLine = section?.lines?.[state.basisChat.lineIndex];
  const currentLine = selectedBasisLine();
  if (!nextLine || !currentLine || nextLine.text === currentLine.text && nextLine.tag === currentLine.tag) return "";
  return `${normalizeBasisTag(nextLine.tag)}: ${nextLine.text}`;
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
    currentText: currentLine.text,
    nextTag: normalizeBasisTag(nextLine.tag),
    nextConfidence: normalizeConfidence(nextLine.confidence ?? nextLine.confidence_pct),
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
  const message = proposal?.message || "Review this proposed quote basis change before applying it.";
  return `
    <div class="basis-chat-proposal-header">
      <div>
        <span>Proposed change</span>
        <strong>${escapeHtml(state.basisChat.scope === "line" ? "Basis line update" : "Quote basis update")}</strong>
      </div>
      <span class="basis-chat-proposal-status">Review</span>
    </div>
    <p class="basis-chat-proposal-summary">${escapeHtml(message)}</p>
    ${delta ? `
      <div class="basis-chat-compare">
        <div class="basis-chat-compare-card">
          <span>Current</span>
          <p><strong>${escapeHtml(basisLinePillLabel({ tag: delta.currentTag, confidence: delta.currentConfidence }))}</strong> ${escapeHtml(delta.currentText)}</p>
        </div>
        <div class="basis-chat-compare-card is-proposed">
          <span>Proposed</span>
          <p><strong>${escapeHtml(basisLinePillLabel({ tag: delta.nextTag, confidence: delta.nextConfidence }))}</strong> ${escapeHtml(delta.nextText)}</p>
        </div>
      </div>
    ` : `
      <p class="basis-chat-proposal-summary">This proposal updates the quotation basis and keeps existing line items available for review.</p>
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

function setBasisChatBusy(isBusy) {
  elements.basisChatPrompt.disabled = isBusy;
  elements.basisChatSendButton.disabled = isBusy;
  elements.basisChatApplyButton.disabled = isBusy || !state.basisChat.proposal;
  elements.basisChatKeepButton.disabled = isBusy || !state.basisChat.proposal;
}

function openBasisChatOverlay(scope = "quote", options = {}) {
  state.basisChat = {
    scope,
    field: options.sectionId || options.field || "",
    sectionId: options.sectionId || options.field || "",
    lineIndex: Number.isInteger(options.lineIndex) ? options.lineIndex : -1,
    line: options.line || "",
    proposal: null,
  };
  resetBasisChatProposal();
  elements.basisChatTitle.textContent = scope === "line" ? "Revise basis line" : "Ask for changes";
  elements.basisChatContext.classList.toggle("has-selected-line", scope === "line");
  if (scope === "line") {
    const line = selectedBasisLine() || parseBasisLine(state.basisChat.line);
    const tag = normalizeBasisTag(line.tag);
    const tagClass = `basis-chat-selected-tag-${tag.toLowerCase()}`;
    state.basisChat.line = `${tag}: ${line.text}`;
    elements.basisChatContext.innerHTML = `
      <span class="basis-chat-context-label">${escapeHtml(basisFieldLabel(state.basisChat.sectionId))}</span>
      <span class="basis-chat-selected-line">
        <strong class="basis-chat-selected-tag ${escapeHtml(tagClass)}" title="${tag === "Confirm" ? "AI confidence before review" : escapeHtml(basisTagLabel(tag))}">${escapeHtml(basisLinePillLabel(line))}</strong>
        <span>${escapeHtml(line.text || state.basisChat.line)}</span>
      </span>
    `;
    elements.basisChatPrompt.placeholder = "Describe the change for this line...";
  } else {
    elements.basisChatContext.textContent = "Ask about or revise the whole quote basis.";
    elements.basisChatPrompt.placeholder = "Ask about this basis or describe a change...";
  }
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
      appendBasisChatMessage("assistant", (started.data.errors || ["I could not answer that yet."]).join("\n"));
      return null;
    }
    const polled = await pollJob(started.data.job_id, (job) => {
      setBusyText(job.status === "running" ? "Answering basis question..." : "Queued...");
    });
    const data = polled.data.result || {};
    if (!polled.ok || ["blocked", "failed"].includes(polled.data.status)) {
      removeBasisChatTyping(typingMessage);
      appendBasisChatMessage("assistant", (data.errors || polled.data.errors || ["I could not answer that yet."]).join("\n"));
      return null;
    }
    if (data.proposal) {
      return { proposal: normalizeServerBasisChatProposal(data.proposal) };
    }
    return { answer: data.answer || null };
  } finally {
    removeBasisChatTyping(typingMessage);
    state.isAnalysisRunning = previousRunning;
    setBusyText("");
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

  const aiResult = await buildAiBasisChatResponse(text);
  if (aiResult?.proposal) {
    setBasisChatProposal(aiResult.proposal);
    return;
  }
  if (aiResult?.answer) {
    appendBasisChatMessage("assistant", aiResult.answer);
    return;
  }

  appendBasisChatMessage("assistant", "AI basis chat did not return a usable response. Try rephrasing the change.");
}

function applyBasisChatProposal() {
  const proposal = state.basisChat.proposal;
  if (!proposal) return;
  state.basisConfirmed = false;
  state.quoteBasisSections = normalizeQuoteBasisSections(proposal.quoteBasisSections || proposal.quoteBasis || state.quoteBasisSections);
  state.quoteBasis = { ...cloneQuoteBasis(proposal.quoteBasis || state.quoteBasis), ...quoteBasisFromSections(state.quoteBasisSections) };
  state.lineItems = Array.isArray(proposal.lineItems) ? proposal.lineItems.map(normalizeLineItem) : [];
  state.outputRows = [];
  state.originalOutputRows = [];
  state.outputErrors = [];
  setDownloadFiles([]);
  updateQuoteBasisCard("edited");
  resetBasisChatProposal();
  appendBasisChatMessage("assistant", "Change applied to the quote basis. Review the updated basis before confirming the quotation basis.");
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
  if (!state.images.length) return "Add at least one reference image before starting analysis.";
  return customerDetailsBlockReason("Complete Customer details before starting analysis")
    || quoteCompanyDetailsBlockReason("Complete Quote Company details before starting analysis");
}

function canStartAnalysis() {
  return !startAnalysisBlockReason();
}

function hasSubmittedQuoteBasis() {
  if (state.aiFailed) return false;
  return ["basis_review", "generating", "pricing_review", "completed"].includes(state.workflowStage)
    && (state.lineItems.length > 0 || state.quoteBasisSections.some((section) => (section.lines || []).length > 0) || Object.values(state.quoteBasis).some((value) => splitLines(value).length > 0));
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

function captureOriginalAnalysisSnapshot(data = {}) {
  const sections = normalizeQuoteBasisSections(data.quote_basis_sections || data.quote_basis || state.quoteBasisSections);
  state.originalAnalysisSnapshot = {
    quote_basis_sections: cloneQuoteBasisSections(sections),
    quote_basis: { ...state.quoteBasis, ...quoteBasisFromSections(sections) },
    line_items: state.lineItems.map(normalizeLineItem),
    boothDimensions: { ...state.boothDimensions },
    source: data.source || state.draftSource || "",
    warnings: Array.isArray(data.warnings) ? [...data.warnings] : [],
  };
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
  const warningText = Array.isArray(data.warnings) && data.warnings.length ? data.warnings.join(" ") : "";
  const message = warningText || "Remote AI did not return a usable quote basis. Start analysis again after checking the local runner connection.";
  setWorkflowStage("basis_review");
  showAiFailureBanner(message);
  renderBasisFailureState(message);
  setSidePanel("basis", { force: true });
  noteWorkflowEvent("assistant", "AI analysis failed before a usable quote basis was produced. I cleared the local fallback draft so it is not mistaken for AI output.", { tone: "error" });
  syncControlStates();
}

function resetQuoteBasisToOriginal() {
  const snapshot = state.originalAnalysisSnapshot;
  if (!snapshot) return;
  state.quoteBasisSections = cloneQuoteBasisSections(snapshot.quote_basis_sections || snapshot.quote_basis || []);
  state.quoteBasis = cloneQuoteBasis(snapshot.quote_basis || quoteBasisFromSections(state.quoteBasisSections));
  state.lineItems = Array.isArray(snapshot.line_items) ? snapshot.line_items.map(normalizeLineItem) : [];
  state.boothDimensions = normalizeBoothDimensions(snapshot.boothDimensions || state.boothDimensions);
  state.outputRows = [];
  state.originalOutputRows = [];
  state.outputErrors = [];
  state.basisConfirmed = false;
  setDownloadFiles([]);
  clearPricingReviewMessages();
  setWorkflowStage("basis_review");
  updateQuoteBasisCard(snapshot.source || "restored");
  noteWorkflowEvent("assistant", "Quote Basis reset to the original AI draft. Uploaded images, Customer, Quote Company, and selected Pricing Reference were kept.");
  syncControlStates();
}

function resetOutputDraft() {
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
  noteWorkflowEvent("assistant", "Output draft reset to the rows created when Quote Basis was confirmed. Images, Customer, Quote Company, and Quote Basis were kept.");
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
        errors: ["Local server connection failed. Make sure the local runner is still open, then retry."],
      },
    };
  }

  let data;
  try {
    data = await response.json();
  } catch {
    data = {
      status: "failed",
      errors: ["Local server returned a non-JSON response. Refresh the app and retry."],
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
        errors: ["Local server connection failed. Make sure the local runner is still open, then retry."],
      },
    };
  }

  let data;
  try {
    data = await response.json();
  } catch {
    data = {
      status: "failed",
      errors: ["Local server returned a non-JSON response. Refresh the app and retry."],
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
  return { ok: false, data: { status: "failed", errors: ["Job was not created."] } };
}

function isInterruptedJobPoll(polled) {
  return Boolean(polled?.aborted || polled?.data?.fetch_failed);
}

function handleInterruptedJobPoll(jobType = "draft") {
  setBusyText("");
  if (jobType === "draft") {
    state.isAnalysisRunning = false;
    showAiFailureBanner("Local server connection was interrupted. Refresh this app to resume the active AI analysis job.");
    setWorkflowStage("analyzing");
  } else {
    state.isGenerating = false;
    setResultStatus("Connection interrupted", "is-warn");
    renderMessages(["Local server connection was interrupted. Refresh this app to resume the active Excel job."], "error");
  }
  syncControlStates();
}

function setBusyText(message = "") {
  if (elements.busyText) elements.busyText.textContent = message;
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
    return;
  }
  elements.healthText.textContent = "Local session unavailable";
}

function setAnalysisButtons() {}

function syncControlStates() {
  const busy = state.isAnalysisRunning || state.isGenerating;
  elements.newQuoteButton.disabled = busy;
  elements.newQuoteButton.title = busy
    ? (state.isAnalysisRunning ? "Analysis is running." : "Quotation generation is running.")
    : "";
  setAnalysisButtons(busy);
  updateSidePanelNav();
  renderCurrentActions();
  saveSessionState();
}

async function handleDraftBasis() {
  if (state.isAnalysisRunning) return;
  const hasFeedback = Boolean(state.pendingFeedback.trim());
  const analysisRequestedAt = new Date().toISOString();

  if (!state.images.length) {
    setWorkflowStage("needs_images");
    showBlockedAction(`Please drop at least one ${currentGenerator().imageNoun} before analysis.`, { details: false });
    noteWorkflowEvent("assistant", `Please drop at least one ${currentGenerator().imageNoun} before analysis.`, { tone: "warn" });
    renderCurrentActions();
    syncControlStates();
    return;
  }
  const missing = missingDetailFields();
  if (missing.length) {
    setWorkflowStage("details_review");
    showBlockedAction(`Fill Customer and Quote Company before AI analysis: ${missing.join(", ")}.`);
    renderBasisEmptyState("Complete the missing Customer and Quote Company details, then start analysis to draft the quote basis.");
    noteWorkflowEvent("assistant", `Fill Customer and Quote Company before AI analysis: ${missing.join(", ")}.`, { tone: "warn" });
    setDetailsDrawer(true);
    syncControlStates();
    return;
  }

  state.isAnalysisRunning = true;
  state.aiFailed = false;
  state.draftSource = "";
  clearAiFailureBanner();
  setWorkflowStage("analyzing");
  showAiRunningBanner("Reading the reference images and preparing the quote basis. Please wait.", analysisRequestedAt);
  clearBasisReviewSurface();
  setSidePanel("basis", { force: true });
  setBusyText("Running analysis...");
  setAnalysisButtons(true);
  syncControlStates();

  const started = await startJob("draft", buildPayload({
    includeBoothDimensions: state.boothDimensions.dimension_source !== "default",
    includeDraftContext: hasFeedback,
  }));
  if (!started.ok) {
    state.isAnalysisRunning = false;
    setBusyText("");
    setAnalysisButtons(false);
    const errors = started.data.errors || ["Draft failed."];
    const wasBlocked = started.data.status === "blocked";
    setWorkflowStage(wasBlocked ? "details_review" : "ready_to_analyze");
    if (wasBlocked) {
      showAiBlockedBanner(errors.join(" "));
    } else {
      showAiFailureBanner("Try again later.");
    }
    noteWorkflowEvent("assistant", errors.join("\n"), { tone: wasBlocked ? "warn" : "error" });
    syncControlStates();
    return;
  }
  const startedAt = started.data.created_at || analysisRequestedAt;
  state.activeJob = { id: started.data.job_id, type: "draft", startedAt };
  showAiRunningBanner("Reading the reference images and preparing the quote basis. Please wait.", startedAt);
  saveSessionState();

  const polled = await pollJob(started.data.job_id, (job) => {
    setBusyText(job.status === "running" ? "Running analysis..." : "Queued...");
  });
  if (polled.aborted) return;
  if (isInterruptedJobPoll(polled)) {
    handleInterruptedJobPoll("draft");
    return;
  }
  state.isAnalysisRunning = false;
  state.activeJob = null;
  setBusyText("");
  setAnalysisButtons(false);

  if (!polled.ok || ["blocked", "failed"].includes(polled.data.status)) {
    const errors = polled.data.errors || polled.data.result?.errors || ["Draft failed."];
    const wasBlocked = polled.data.status === "blocked";
    setWorkflowStage(wasBlocked ? "details_review" : "ready_to_analyze");
    if (wasBlocked) {
      showAiBlockedBanner(errors.join(" "));
    } else {
      showAiFailureBanner("Try again later.");
    }
    noteWorkflowEvent("assistant", errors.join("\n"), { tone: wasBlocked ? "warn" : "error" });
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
  const warningText = Array.isArray(data.warnings) && data.warnings.length ? data.warnings.join(" ") : "";
  clearAiFailureBanner();
  if (Array.isArray(data.warnings) && data.warnings.length) {
    noteWorkflowEvent("assistant", data.warnings.join("\n"), { tone: "warn" });
  }
  updateQuoteBasisCard(data.source);
  setSidePanel("basis", { force: true });
  if (!state.lineItems.length) {
    noteWorkflowEvent("assistant", EMPTY_LINE_ITEMS_MESSAGE, { tone: "warn" });
  }
  syncControlStates();
}

async function confirmBasis() {
  if (state.isAnalysisRunning || state.isGenerating) return;
  const confirmBlockReason = basisConfirmBlockReason();
  if (confirmBlockReason) {
    setWorkflowStage("basis_review");
    noteWorkflowEvent("assistant", confirmBlockReason, { tone: "warn" });
    syncControlStates();
    return;
  }
  if (state.aiFailed) {
    showAiFailureBanner("Try again later. Regenerate analysis before confirming the quote basis.");
    noteWorkflowEvent("assistant", "AI analysis failed. Try again later before confirming or generating the quotation.", { tone: "error" });
    renderCurrentActions();
    return;
  }
  if (!state.lineItems.length) {
    setWorkflowStage("basis_review");
    noteWorkflowEvent("assistant", "I do not have any generated quotation line items yet. Regenerate analysis before confirming the basis.", { tone: "warn" });
    renderCurrentActions();
    return;
  }
  const missing = missingDetailFields();
  if (missing.length) {
    setWorkflowStage("details_review");
    noteWorkflowEvent("assistant", `Please fill these details before I generate Excel: ${missing.join(", ")}.`, { tone: "warn" });
    setDetailsDrawer(true);
    syncControlStates();
    return;
  }
  state.basisConfirmed = true;
  ensureOutputRowsFromLineItems();
  state.originalOutputRows = snapshotOutputRows(state.outputRows);
  state.lineItems = outputRowsToLineItems();
  setWorkflowStage("completed");
  renderPricingMatches(state.outputRows);
  renderMatchSummary({ pricing_matches: state.outputRows });
  setResultStatus("Ready for pricing review", "is-warn");
  renderOutputValidationMessages(outputRowsValid().errors);
  setSidePanel("output", { force: true });
  noteWorkflowEvent("assistant", "Basis confirmed. Review the editable Output rows, then Download Excel when pricing is valid.");
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
      noteWorkflowEvent("assistant", "Confirm Quotation Basis before generating the Excel quotation.", { tone: "warn" });
    }
    syncControlStates();
    return;
  }
  if (state.aiFailed) {
    setWorkflowStage("basis_review");
    showAiFailureBanner("Try again later. Regenerate analysis before generating Excel.");
    noteWorkflowEvent("assistant", "Cannot generate because AI analysis failed. Try again later or regenerate analysis first.", { tone: "error" });
    syncControlStates();
    return;
  }
  const missing = missingDetailFields();
  if (missing.length) {
    setWorkflowStage("details_review");
    noteWorkflowEvent("assistant", `Please fill these details before I generate Excel: ${missing.join(", ")}.`, { tone: "warn" });
    setDetailsDrawer(true);
    syncControlStates();
    return;
  }
  if (!state.lineItems.length) {
    setWorkflowStage("basis_review");
    noteWorkflowEvent("assistant", "There are no generated line items yet. Run analysis first so Excel has quotation rows.", { tone: "warn" });
    renderCurrentActions();
    return;
  }

  state.isGenerating = true;
  setWorkflowStage("generating");
  setBusyText("Generating Excel...");
  setResultStatus("Generating Excel", "is-warn");
  renderMessages([]);
  setDownloadFiles([]);
  renderMatchSummary({});
  clearPricingReviewMessages();
  syncControlStates();
  const started = await startJob("generate", buildPayload());
  if (!started.ok) {
    state.isGenerating = false;
    setBusyText("");
    setWorkflowStage(state.activeSidePanel === "output" ? "completed" : "details_review");
    setResultStatus(started.data.status || "Failed", "is-bad");
    renderMessages(started.data.errors || ["Generation failed."], "error");
    noteWorkflowEvent("assistant", (started.data.errors || ["Generation failed."]).join("\n"), { tone: "error" });
    syncControlStates();
    return;
  }
  state.activeJob = { id: started.data.job_id, type: "generate" };
  saveSessionState();

  const polled = await pollJob(started.data.job_id, (job) => {
    setBusyText(job.status === "running" ? "Generating Excel..." : "Queued...");
  });
  if (polled.aborted) return;
  if (isInterruptedJobPoll(polled)) {
    handleInterruptedJobPoll("generate");
    return;
  }
  state.isGenerating = false;
  state.activeJob = null;
  setBusyText("");

  const data = polled.data.result || polled.data || {};
  if (!polled.ok || ["blocked", "failed"].includes(polled.data.status) || data.status === "blocked" || data.status === "failed") {
    setWorkflowStage(state.activeSidePanel === "output" ? "completed" : "details_review");
    setResultStatus(data.status || "Failed", "is-bad");
    renderMessages(data.errors || ["Generation failed."], "error");
    if (data.pricing_matches?.length) renderPricingMatches(data.pricing_matches || [], { fromPricingMatches: true });
    renderMatchSummary(data);
    noteWorkflowEvent("assistant", (data.errors || ["Generation failed."]).join("\n"), { tone: "error" });
    syncControlStates();
    return;
  }

  const needsPricingReview = polled.data.status === "needs_review"
    || data.status === "needs_confirmation"
    || pricingReviewIssues(data).length > 0;
  if (needsPricingReview) {
    setWorkflowStage("completed");
    setResultStatus("Needs pricing review", "is-warn");
    renderPricingReviewMessages(data);
    setSidePanel("output", { force: true });
    setDownloadFiles([]);
    noteWorkflowEvent("assistant", "I found pricing items that need review. Resolve them in Output before downloading Excel.", { tone: "warn" });
  } else {
    setWorkflowStage("completed");
    setResultStatus("Completed", "is-ok");
    renderMessages([]);
    clearPricingReviewMessages();
    setSidePanel("output", { force: true });
    noteWorkflowEvent("assistant", "Excel quotation is ready. Use Download Excel in the Output footer.", { tone: "instruction" });
    setDownloadFiles(data.files || []);
  }
  if (data.pricing_matches?.length) renderPricingMatches(data.pricing_matches || [], { fromPricingMatches: true });
  renderMatchSummary(data);
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
    setBusyText("Checking saved analysis job...");
    syncControlStates();

    const polled = await pollJob(activeJob.id, (job) => {
      if (job.created_at && !state.activeJob?.startedAt) {
        state.activeJob = { ...state.activeJob, startedAt: job.created_at };
        showAiRunningBanner("Resuming the analysis job after refresh.", job.created_at);
        saveSessionState();
      }
      setBusyText(job.status === "running" ? "Running analysis..." : "Queued...");
    });
    if (polled.aborted) return;
    if (isInterruptedJobPoll(polled)) {
      handleInterruptedJobPoll("draft");
      return;
    }
    state.isAnalysisRunning = false;
    state.activeJob = null;
    setBusyText("");
    setAnalysisButtons(false);

    if (!polled.ok || ["blocked", "failed"].includes(polled.data.status)) {
      setWorkflowStage("ready_to_analyze");
      showAiFailureBanner("Try again later.");
      noteWorkflowEvent("assistant", (polled.data.errors || polled.data.result?.errors || ["Draft failed."]).join("\n"), { tone: "error" });
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
    if (Array.isArray(data.warnings) && data.warnings.length) {
      noteWorkflowEvent("assistant", data.warnings.join("\n"), { tone: "warn" });
    }
    updateQuoteBasisCard(data.source);
    setSidePanel("basis", { force: true });
    if (!state.lineItems.length) {
      noteWorkflowEvent("assistant", EMPTY_LINE_ITEMS_MESSAGE, { tone: "warn" });
    }
    syncControlStates();
    return;
  }

  if (activeJob.type === "generate") {
    state.isGenerating = true;
    state.isAnalysisRunning = false;
    setWorkflowStage("generating");
    setBusyText("Checking saved Excel job...");
    setResultStatus("Checking Excel", "is-warn");
    setSidePanel("output", { force: true });
    syncControlStates();

    const polled = await pollJob(activeJob.id, (job) => {
      setBusyText(job.status === "running" ? "Checking Excel..." : "Queued...");
    });
    if (polled.aborted) return;
    if (isInterruptedJobPoll(polled)) {
      handleInterruptedJobPoll("generate");
      return;
    }
    state.isGenerating = false;
    state.activeJob = null;
    setBusyText("");

    const data = polled.data.result || polled.data || {};
    if (!polled.ok || ["blocked", "failed"].includes(polled.data.status) || data.status === "blocked" || data.status === "failed") {
      setWorkflowStage("details_review");
      setResultStatus(data.status || "Failed", "is-bad");
      renderMessages(data.errors || ["Generation failed."], "error");
      if (data.pricing_matches?.length) renderPricingMatches(data.pricing_matches || [], { fromPricingMatches: true });
      renderMatchSummary(data);
      noteWorkflowEvent("assistant", (data.errors || ["Generation failed."]).join("\n"), { tone: "error" });
      syncControlStates();
      return;
    }

    const needsPricingReview = polled.data.status === "needs_review"
      || data.status === "needs_confirmation"
      || pricingReviewIssues(data).length > 0;
    if (needsPricingReview) {
      setWorkflowStage("pricing_review");
      setResultStatus("Needs pricing review", "is-warn");
      renderPricingReviewMessages(data);
      setSidePanel("output", { force: true });
      setDownloadFiles([]);
      noteWorkflowEvent("assistant", "I found pricing items that need review. Resolve them in Output before downloading Excel.", { tone: "warn" });
    } else {
      setWorkflowStage("completed");
      setResultStatus("Completed", "is-ok");
      renderMessages([]);
      clearPricingReviewMessages();
      setSidePanel("output");
      noteWorkflowEvent("assistant", "Excel quotation is ready. Use Download Excel in the Output footer.", { tone: "instruction" });
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

function requestStartAnalysis() {
  const reason = startAnalysisBlockReason();
  if (reason) {
    showBlockedAction(reason);
    noteWorkflowEvent("assistant", reason, { tone: "warn" });
    syncControlStates();
    return;
  }
  elements.analysisConfirmModal.hidden = false;
  elements.analysisConfirmModal.classList.add("is-open");
  window.setTimeout(() => elements.analysisConfirmStartButton?.focus(), 0);
}

function closeAnalysisConfirmModal() {
  elements.analysisConfirmModal.classList.remove("is-open");
  elements.analysisConfirmModal.hidden = true;
}

function confirmStartAnalysis() {
  closeAnalysisConfirmModal();
  handleDraftBasis();
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
  if (!state.images.length) return "Add reference images before opening this step.";
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
    const confirmBlockReason = basisConfirmBlockReason();
    if (confirmBlockReason) return confirmBlockReason;
    if (!hasCompletedQuoteBasis()) return "Confirm Quotation Basis before opening Output.";
  }
  return "";
}

function canOpenSidePanel(panelName) {
  return !sidePanelBlockReason(panelName);
}

function setSidePanel(panelName, options = {}) {
  const panelTitles = {
    images: ["Images", "Reference Inputs", currentGenerator().intakeSubtitle],
    customer: ["Customer", "Customer Details", "Customer, project, booth size, and customer address for this quotation."],
    quote_company: ["Quote Company", "Quotation Defaults", "Reusable quotation-company header, payment terms, notes, and signature defaults."],
    basis: ["Quote Basis", "Confirm Draft", "Review the drafted basis, revise individual lines, or request a full-basis change."],
    output: ["Output", "Editable Pricing", "Review quotation rows, resolve pricing, and download Excel."],
  };
  const nextPanel = panelTitles[panelName] ? panelName : "images";
  const blockReason = sidePanelBlockReason(nextPanel);
  if (blockReason && !options.force) {
    if (options.notify) noteWorkflowEvent("assistant", blockReason, { tone: "warn" });
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
  elements.sideWorkspace.setAttribute("aria-hidden", "false");
  updateSidePanelNav();
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
  elements.discussQuoteButton.hidden = state.activeSidePanel !== "basis";
  elements.resetQuoteBasisButton.hidden = state.activeSidePanel !== "basis";
  elements.resetOutputButton.hidden = state.activeSidePanel !== "output";
  elements.resetImagesButton.disabled = busy || !state.images.length;
  elements.clearCustomerButton.disabled = busy;
  elements.clearQuoteCompanyButton.disabled = busy;
  elements.discussQuoteButton.disabled = busy || !hasSubmittedQuoteBasis();
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
    noteWorkflowEvent("assistant", reason, { tone: "warn" });
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
      line: line ? `${normalizeBasisTag(line.tag)}: ${line.text}` : "",
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
  elements.sideDownloadButton.addEventListener("click", (event) => {
    if (elements.sideDownloadButton.getAttribute("aria-disabled") === "true") {
      event.preventDefault();
      const validation = outputRowsValid();
      if (!validation.valid) renderOutputValidationMessages(validation.errors);
      return;
    }
    if (!state.downloadFile?.url) {
      event.preventDefault();
      handleGenerate();
    }
  });
  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape") {
      if (!elements.basisChatOverlay.hidden) {
        closeBasisChatOverlay();
      } else if (!elements.pricingReferenceModal.hidden) {
        closePricingReferenceModal();
      } else if (!elements.analysisConfirmModal.hidden) {
        closeAnalysisConfirmModal();
      }
    }
  });

  elements.sampleDetailsButton.addEventListener("click", setSampleDetails);
  elements.newQuoteButton.addEventListener("click", startNewQuote);
  elements.profileSelect.addEventListener("change", handleProfileSelectionChange);
  elements.newPricingReferenceButton.addEventListener("click", openPricingReferenceModal);
  elements.deletePricingReferenceButton.addEventListener("click", deleteSelectedPricingReference);
  elements.pricingReferenceForm.addEventListener("submit", savePricingReferenceFromModal);
  elements.pricingReferenceTemplateButton.addEventListener("click", downloadPricingReferenceTemplate);
  elements.pricingReferenceFile.addEventListener("change", async () => {
    const file = elements.pricingReferenceFile.files?.[0];
    if (elements.pricingReferenceFileName) {
      elements.pricingReferenceFileName.textContent = file?.name || "No file chosen";
    }
    const result = await validatePricingReferenceFile(file);
    state.pendingPricingReference = result;
    renderPricingReferencePreview(result);
  });
  elements.pricingReferenceCancelButton.addEventListener("click", closePricingReferenceModal);
  elements.pricingReferenceCloseButton.addEventListener("click", closePricingReferenceModal);
  elements.pricingReferenceModal.addEventListener("click", (event) => {
    if (event.target.closest("[data-pricing-reference-close]")) closePricingReferenceModal();
  });
  elements.analysisConfirmCancelButton.addEventListener("click", closeAnalysisConfirmModal);
  elements.analysisConfirmStartButton.addEventListener("click", confirmStartAnalysis);
  elements.analysisConfirmModal.addEventListener("click", (event) => {
    if (event.target.closest("[data-analysis-confirm-close]")) closeAnalysisConfirmModal();
  });
  elements.pricingMatchesBody.addEventListener("pointerdown", handleOutputIncludedPointerDown);
  elements.pricingMatchesBody.addEventListener("click", handleOutputCellClick);
  elements.pricingMatchesBody.addEventListener("dblclick", handleOutputCellOpen);
  elements.pricingMatchesBody.addEventListener("keydown", handleOutputCellKeydown);
  elements.pricingMatchesBody.addEventListener("focusout", handleOutputEditorCommit);
  elements.pricingMatchesBody.addEventListener("change", handleOutputEditorCommit);
  elements.discussQuoteButton.addEventListener("click", () => openBasisChatOverlay("quote"));
  elements.resetImagesButton.addEventListener("click", resetImagesDraft);
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
  ].forEach((input) => {
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
  elements.pricingReviewMessages.addEventListener("click", (event) => {
    const button = event.target.closest("[data-pricing-action]");
    if (!button || button.disabled) return;
    handlePricingChoice(button.dataset.pricingAction, button.dataset.pricingIssue);
  });
}

function setInitialValues() {
  updateGeneratorCopy();
  renderProfileOptions();
  renderPresetOptions();
  renderHeaderLogoPreview();
  renderPresetStatus();
  if (restoreSessionState()) {
    renderCurrentActions();
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
  renderCurrentActions();
  syncControlStates();
  setSidePanel("images");
}

async function boot() {
  wireEvents();
  try {
    await initializeSession();
    await loadProfiles();
    setInitialValues();
    checkHealth();
  } finally {
    state.isBooting = false;
    syncControlStates();
  }
  await resumeSavedJob();
}

boot();
