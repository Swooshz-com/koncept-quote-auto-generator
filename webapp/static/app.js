const EMPTY_LINE_ITEMS_MESSAGE = "AI analysis will populate line items here.";
const DEFAULT_PROFILE_ID = "koncept";
const DEFAULT_SAMPLE_ID = "brazil-pavilion";
const CSRF_HEADER_NAME = "X-Swooshz-CSRF";
const QUOTE_PRESETS_STORAGE_KEY = "swooshz_quote_detail_presets_v1";
const QUOTE_SESSION_STORAGE_KEY = "swooshz_quote_session_v1";
const FINAL_JOB_STATUSES = new Set(["completed", "degraded", "needs_review", "blocked", "failed"]);
const PROFILE_PRESET_PREFIX = "profile:";
const LOCAL_PRESET_PREFIX = "local:";
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
  dropMeta: "JPG, PNG, or WebP. Add more references anytime; files stay local to this runner.",
  imageNoun: "reference image",
  analyzeLabel: "Start Analysis",
  fallbackAction: "review the reference images",
};

const SIDE_PANEL_SEQUENCE = ["images", "customer", "quote_company", "basis", "pricing", "output"];
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
  ["Include", "Matched", "Quoted in the draft"],
  ["Confirm", "Confirm", "Needs your decision"],
  ["Exclude", "Exclude", "Not included unless requested"],
  ["Assumption", "Assumption", "Must be confirmed or changed"],
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
  profiles: [],
  pricingReferences: [],
  images: [],
  headerLogo: null,
  workflowStage: "needs_images",
  quoteBasis: { ...EMPTY_BASIS },
  lineItems: [],
  boothDimensions: { ...DEFAULT_BOOTH_DIMENSIONS },
  basisConfirmed: false,
  chatMessages: [],
  isAnalysisRunning: false,
  isGenerating: false,
  aiFailed: false,
  draftSource: "",
  isBooting: true,
  csrfHeaderName: CSRF_HEADER_NAME,
  csrfToken: "",
  pendingFeedback: "",
  activeSidePanel: "images",
  downloadFile: null,
  pricingMatches: [],
  pricingIssues: [],
  activeJob: null,
  basisChat: {
    scope: "quote",
    field: "",
    line: "",
    proposal: null,
  },
};

const qs = (selector) => document.querySelector(selector);

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
  fileList: qs("#fileList"),
  customerDetailsButton: qs("#customerDetailsButton"),
  quoteCompanyButton: qs("#quoteCompanyButton"),
  quoteBasisButton: qs("#quoteBasisButton"),
  pricingButton: qs("#pricingButton"),
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
  workflowNotice: qs("#workflowNotice"),
  sampleDetailsButton: qs("#sampleDetailsButton"),
  clientName: qs("#clientName"),
  clientAttention: qs("#clientAttention"),
  clientTitle: qs("#clientTitle"),
  clientAddress: qs("#clientAddress"),
  projectTitle: qs("#projectTitle"),
  quoteDate: qs("#quoteDate"),
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
  chatTranscript: qs("#chatTranscript"),
  chatActions: qs("#chatActions"),
  chatForm: qs("#chatForm"),
  chatPrompt: qs("#chatPrompt"),
  sendChatButton: qs("#sendChatButton"),
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
  clearCustomerButton: qs("#clearCustomerButton"),
  clearQuoteCompanyButton: qs("#clearQuoteCompanyButton"),
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
  richTextEditors: Array.from(document.querySelectorAll("[data-rich-text-source]")),
  richTextToolbar: Array.from(document.querySelectorAll("[data-rich-command]")),
};

function pricingReferenceProfileId(reference = {}) {
  return String(reference.profile_id || DEFAULT_PROFILE_ID).trim() || DEFAULT_PROFILE_ID;
}

function pricingReferenceSelectValue(reference = {}) {
  const referenceId = String(reference.id || "").trim();
  if (!referenceId) return "";
  return `${pricingReferenceProfileId(reference)}::${referenceId}`;
}

function pricingReferenceSelectionFromValue(value = "") {
  const text = String(value || "").trim();
  if (!text) return { profileId: "", pricingReferenceId: "" };
  const delimiterIndex = text.indexOf("::");
  if (delimiterIndex < 0) {
    return { profileId: "", pricingReferenceId: text };
  }
  return {
    profileId: text.slice(0, delimiterIndex) || DEFAULT_PROFILE_ID,
    pricingReferenceId: text.slice(delimiterIndex + 2),
  };
}

function pricingReferenceForProfile(profileId = "") {
  const resolvedProfileId = String(profileId || DEFAULT_PROFILE_ID).trim() || DEFAULT_PROFILE_ID;
  return state.pricingReferences.find((reference) => pricingReferenceProfileId(reference) === resolvedProfileId) || null;
}

function currentProfile() {
  const selectedReference = currentPricingReference();
  const resolvedProfileId = pricingReferenceProfileId(selectedReference || { profile_id: state.profileId || DEFAULT_PROFILE_ID });
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

function currentPricingReference() {
  const pricingReferenceId = String(state.pricingReferenceId || "").trim();
  if (!pricingReferenceId) return null;
  const selectedProfileId = String(state.profileId || "").trim();
  return state.pricingReferences.find((reference) => (
    reference.id === pricingReferenceId
    && (!selectedProfileId || pricingReferenceProfileId(reference) === selectedProfileId)
  ))
    || state.pricingReferences.find((reference) => reference.id === pricingReferenceId)
    || null;
}

function resolvedProfileIdForPayload() {
  const selectedReference = currentPricingReference();
  return pricingReferenceProfileId(selectedReference || { profile_id: state.profileId || DEFAULT_PROFILE_ID });
}

function syncSelectedPricingReference() {
  const selectedReference = currentPricingReference();
  if (selectedReference) {
    state.profileId = pricingReferenceProfileId(selectedReference);
    state.pricingReferenceId = selectedReference.id || "";
    return;
  }
  if (state.profileId) {
    const fallbackReference = pricingReferenceForProfile(state.profileId);
    if (fallbackReference) {
      state.profileId = pricingReferenceProfileId(fallbackReference);
      state.pricingReferenceId = fallbackReference.id || "";
      return;
    }
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
  if (elements.assistantSubtitle) elements.assistantSubtitle.textContent = generator.assistantSubtitle;
  if (state.activeSidePanel === "images") elements.sideDrawerSubtitle.textContent = generator.intakeSubtitle;
  elements.dropTitle.textContent = state.images.length ? "Add more reference images" : generator.dropTitle;
  elements.dropMeta.textContent = generator.dropMeta;
}

function setWorkflowStage(stage) {
  state.workflowStage = stage;
  if (elements.workflowStage) {
    elements.workflowStage.textContent = STAGE_LABELS[stage] || stage;
    elements.workflowStage.dataset.stage = stage;
  }
  document.body.dataset.workflowStage = stage;
}

function setAiStatusBanner(tone, title, message) {
  elements.aiFailureBanner.hidden = false;
  elements.aiFailureBanner.classList.toggle("is-running", tone === "running");
  elements.aiFailureBanner.classList.toggle("is-failure", tone !== "running");
  elements.aiFailureBanner.innerHTML = `
    <strong>${escapeHtml(title)}</strong>
    <span>${escapeHtml(message)}</span>
  `;
}

function showAiRunningBanner(message = "Reading the reference images and preparing the quote basis. Please wait.") {
  setAiStatusBanner("running", "AI analysis running.", message);
}

function showAiFailureBanner(message = "AI analysis failed. Try again later.") {
  const detail = message.replace(/^AI analysis failed\.?\s*/i, "").trim() || "Try again later.";
  setAiStatusBanner("failure", "AI analysis failed.", detail);
}

function showAiBlockedBanner(message = "Complete the required details, then retry analysis.") {
  const detail = String(message || "").replace(/^AI analysis blocked\.?\s*/i, "").trim() || "Complete the required details, then retry analysis.";
  setAiStatusBanner("failure", "AI analysis blocked.", detail);
}

function clearAiFailureBanner() {
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

async function filesToImageEntries(files) {
  return Promise.all(
    Array.from(files)
      .filter((file) => file.type.startsWith("image/"))
      .map(async (file) => ({
        name: file.name,
        type: file.type,
        size: file.size,
        data_url: await fileToDataUrl(file),
      }))
  );
}

async function addImagesFromFiles(files) {
  const images = await filesToImageEntries(files);
  if (!images.length) return;
  state.images = [...state.images, ...images];
  renderFiles();
  setWorkflowStage("ready_to_analyze");
  const generator = currentGenerator();
  appendChatMessage(
    "assistant",
    `${state.images.length} ${generator.imageNoun}${state.images.length === 1 ? "" : "s"} loaded. Click ${generator.analyzeLabel} when you want me to draft the quote basis.`
  );
  syncControlStates();
}

function removeImageAt(index) {
  state.images.splice(index, 1);
  renderFiles();
  if (!state.images.length) {
    setWorkflowStage("needs_images");
    appendChatMessage("assistant", `No ${currentGenerator().imageNoun}s are loaded now. Drop references to start the quote analysis.`);
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

function renderPlainText(value) {
  const lines = normalizeTextNewlines(value)
    .split(/\n+/)
    .map((line) => line.trim())
    .filter(Boolean);
  if (!lines.length) return "<p></p>";
  return lines.map((line) => `<p>${escapeHtml(line)}</p>`).join("");
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

function applyDefaultQuoteDate() {
  if (!elements.quoteDate.value.trim()) {
    setInputValue(elements.quoteDate, todayDateInputValue());
  }
}

function renderFiles() {
  elements.imageIntake.classList.toggle("has-images", Boolean(state.images.length));
  updateGeneratorCopy();
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
  return {
    section: item.section || "",
    quantity: item.quantity ?? "",
    unit: item.unit || "",
    description: item.description || "",
    pricing_keyword: item.pricing_keyword || "",
    display_price: item.display_price || "",
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
  return RICH_TEXT_SOURCE_IDS.reduce((details, id) => {
    const editor = elements.richTextEditors.find((item) => item.dataset.richTextSource === id);
    if (editor) details[id] = sanitizeRichTextHtml(editor.innerHTML);
    return details;
  }, {});
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
      version: 1,
      savedAt: new Date().toISOString(),
      profileId: state.profileId,
      pricingReferenceId: state.pricingReferenceId,
      images: state.images,
      quoteDetails: collectQuoteDetails(),
      workflowStage: state.workflowStage,
      quoteBasis: state.quoteBasis,
      lineItems: state.lineItems,
      boothDimensions: state.boothDimensions,
      basisConfirmed: state.basisConfirmed,
      aiFailed: state.aiFailed,
      draftSource: state.draftSource,
      activeSidePanel: state.activeSidePanel,
      downloadFile: state.downloadFile,
      pricingMatches: state.pricingMatches,
      pricingIssues: state.pricingIssues,
      chatMessages: state.chatMessages.slice(-80),
      activeJob: state.activeJob,
    };
    window.localStorage.setItem(QUOTE_SESSION_STORAGE_KEY, JSON.stringify(snapshot));
  } catch {
    // Large image payloads can exceed storage quota; refresh recovery is best-effort.
  }
}

function restoreSessionState() {
  const saved = safeSessionJson();
  if (!saved || saved.version !== 1) return false;
  state.profileId = saved.profileId || "";
  state.pricingReferenceId = saved.pricingReferenceId || saved.profileId || "";
  syncSelectedPricingReference();
  renderProfileOptions();
  applyQuoteDetails(saved.quoteDetails || {}, { includeLogo: true, clearLogo: true });
  state.images = Array.isArray(saved.images) ? saved.images : [];
  state.quoteBasis = cloneQuoteBasis(saved.quoteBasis || {});
  state.lineItems = Array.isArray(saved.lineItems) ? saved.lineItems.map(normalizeLineItem) : [];
  state.boothDimensions = normalizeBoothDimensions(saved.boothDimensions || saved.quoteDetails?.project || {});
  state.basisConfirmed = Boolean(saved.basisConfirmed);
  state.aiFailed = Boolean(saved.aiFailed);
  state.draftSource = saved.draftSource || "";
  state.downloadFile = saved.downloadFile || null;
  state.pricingMatches = Array.isArray(saved.pricingMatches) ? saved.pricingMatches : [];
  state.pricingIssues = Array.isArray(saved.pricingIssues) ? saved.pricingIssues : [];
  state.chatMessages = Array.isArray(saved.chatMessages) ? saved.chatMessages : [];
  state.activeJob = saved.activeJob && saved.activeJob.id ? saved.activeJob : null;
  renderFiles();
  renderChat();
  renderPricingMatches(state.pricingMatches);
  renderMatchSummary({ pricing_matches: state.pricingMatches });
  if (state.pricingIssues.length) {
    renderPricingReviewMessages({ errors: state.pricingIssues, pricing_matches: state.pricingMatches });
  } else {
    clearPricingReviewMessages();
  }
  if (state.lineItems.length || Object.values(state.quoteBasis).some((value) => splitLines(value).length > 0)) {
    updateQuoteBasisCard(saved.draftSource || "restored");
  } else {
    renderBasisEmptyState();
  }
  updateDownloadButton();
  setResultStatus(state.downloadFile ? "Completed" : "No job yet", state.downloadFile ? "is-ok" : "");
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
  return elements.presetSelect.value || "";
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
  if (configured && presets.some((preset) => preset.id === configured)) return configured;
  if (presets.some((preset) => preset.id === "default")) return "default";
  return presets[0]?.id || "";
}

function profilePresetOptionValue(presetId) {
  return `${PROFILE_PRESET_PREFIX}${presetId}`;
}

function localPresetOptionValue(presetId) {
  return `${LOCAL_PRESET_PREFIX}${presetId}`;
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
  const builtInOptions = builtInPresets
    .map((preset) => `<option value="${escapeHtml(profilePresetOptionValue(preset.id))}">${escapeHtml(preset.name)}</option>`)
    .join("");
  const localOptions = presets
    .map((preset) => `<option value="${escapeHtml(localPresetOptionValue(preset.id))}">${escapeHtml(preset.name)}</option>`)
    .join("");
  elements.presetSelect.innerHTML = [
    `<option value="">No preset selected</option>`,
    builtInOptions ? `<optgroup label="Company presets">${builtInOptions}</optgroup>` : "",
    localOptions ? `<optgroup label="Saved locally">${localOptions}</optgroup>` : "",
    !builtInOptions && !localOptions ? `<option value="">No presets available</option>` : "",
  ].join("");
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
  elements.presetSelect.value = localPresetOptionValue(preset.id);
  updatePresetButtons();
  renderPresetStatus(`Saved "${name}" as a local company preset.`);
}

function loadSelectedPreset(options = {}) {
  const preset = selectedPreset();
  if (!preset) {
    renderPresetStatus("Choose a preset to load.");
    return;
  }
  const details = preset.details || {};
  const clearsLogo = Boolean(details.company && typeof details.company === "object");
  applyQuoteDetails(details, { includeLogo: true, clearLogo: clearsLogo, partial: true });
  clearGeneratedQuoteState();
  setWorkflowStage(state.images.length ? "ready_to_analyze" : "needs_images");
  syncControlStates();
  if (!options.silent) {
    appendChatMessage("assistant", `Loaded company preset "${preset.name}". Pricing reference and reference images are unchanged.`);
  }
  renderPresetStatus(`Loaded "${preset.name}".`);
}

function loadDefaultProfilePreset(options = {}) {
  const defaultPreset = defaultProfilePresetId();
  if (!defaultPreset) return;
  elements.presetSelect.value = profilePresetOptionValue(defaultPreset);
  loadSelectedPreset(options);
}

function clearCustomerDetails() {
  state.profileId = "";
  state.pricingReferenceId = "";
  setInputValue(elements.clientName, "");
  setInputValue(elements.clientAttention, "");
  setInputValue(elements.clientTitle, "");
  setInputValue(elements.clientAddress, "");
  setInputValue(elements.projectTitle, "");
  state.boothDimensions = { ...DEFAULT_BOOTH_DIMENSIONS };
  setInputValue(elements.quoteDate, todayDateInputValue());
  setInputValue(elements.projectNumber, "");
  clearGeneratedQuoteState();
  renderProfileOptions();
  setWorkflowStage(state.images.length ? "ready_to_analyze" : "needs_images");
  syncControlStates();
  appendChatMessage("assistant", "Customer details and pricing reference cleared. Images and quote-company defaults were left unchanged.");
}

function clearQuoteCompanyDetails() {
  elements.presetSelect.value = "";
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
  clearGeneratedQuoteState();
  setWorkflowStage(state.images.length ? "ready_to_analyze" : "needs_images");
  updatePresetButtons();
  renderHeaderLogoPreview();
  syncControlStates();
  renderPresetStatus("Quote-company defaults reset. Header, logo, company, payment terms, notes, and signatory fields still need to be filled or loaded.");
  appendChatMessage("assistant", "Quote-company defaults reset. Customer details, pricing reference, and images were left unchanged.");
}

function startNewQuote() {
  if (state.isAnalysisRunning || state.isGenerating) return;
  clearSessionState();
  state.profileId = "";
  state.pricingReferenceId = "";
  state.images = [];
  state.headerLogo = null;
  state.boothDimensions = { ...DEFAULT_BOOTH_DIMENSIONS };
  state.pendingFeedback = "";
  state.chatMessages = [];
  state.downloadFile = null;
  elements.imageInput.value = "";
  elements.headerLogoInput.value = "";
  elements.presetSelect.value = "";
  elements.presetNameInput.value = "";
  applyQuoteDetails({}, { clearLogo: true });
  applyDefaultQuoteCompanyFields();
  clearGeneratedQuoteState();
  renderFiles();
  renderProfileOptions();
  renderPresetOptions();
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
  renderPresetStatus(`Deleted "${preset.name}".`);
}

function renderProfileOptions() {
  if (!elements.profileSelect) return;
  const references = state.pricingReferences.length ? state.pricingReferences : [];
  const duplicateReferenceIds = references.reduce((counts, reference) => {
    const referenceId = String(reference.id || "").trim();
    if (referenceId) counts.set(referenceId, (counts.get(referenceId) || 0) + 1);
    return counts;
  }, new Map());
  const profileLabels = new Map(state.profiles.map((profile) => [profile.id, profile.label || profile.id]));
  const options = references
    .map((reference) => {
      const referenceId = String(reference.id || "").trim();
      const profileId = pricingReferenceProfileId(reference);
      const duplicateSuffix = duplicateReferenceIds.get(referenceId) > 1 ? ` (${profileLabels.get(profileId) || profileId})` : "";
      return `<option value="${escapeHtml(pricingReferenceSelectValue(reference))}">${escapeHtml(reference.label || referenceId)}${escapeHtml(duplicateSuffix)}</option>`;
    })
    .join("");
  elements.profileSelect.innerHTML = `<option value="">No pricing reference selected</option>${options}`;
  const selectedReference = currentPricingReference();
  elements.profileSelect.value = selectedReference ? pricingReferenceSelectValue(selectedReference) : "";
}

async function loadProfiles() {
  const { ok, data } = await getJson("/api/profiles");
  if (ok && Array.isArray(data.profiles)) {
    state.profiles = data.profiles;
    state.pricingReferences = Array.isArray(data.pricing_references) ? data.pricing_references : [];
    if (state.pricingReferenceId) {
      syncSelectedPricingReference();
    }
  }
  renderProfileOptions();
}

function clearGeneratedQuoteState() {
  state.quoteBasis = { ...EMPTY_BASIS };
  state.lineItems = [];
  state.basisConfirmed = false;
  state.aiFailed = false;
  state.draftSource = "";
  state.activeJob = null;
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
  if (
    nextSelection.profileId === state.profileId
    && nextSelection.pricingReferenceId === state.pricingReferenceId
  ) {
    return;
  }
  state.profileId = nextSelection.profileId;
  state.pricingReferenceId = nextSelection.pricingReferenceId;
  syncSelectedPricingReference();
  renderProfileOptions();
  clearGeneratedQuoteState();
  setWorkflowStage(state.images.length ? (canStartAnalysis() ? "ready_to_analyze" : "details_review") : "needs_images");
  appendChatMessage("assistant", "Pricing reference changed. I cleared the previous draft so the next analysis uses the selected pricing context. Customer and quote-company details were left unchanged.");
  syncControlStates();
}

async function setSampleDetails() {
  if (state.isBooting || state.isAnalysisRunning || state.isGenerating) return;
  elements.sampleDetailsButton.disabled = true;
  try {
    const { ok, data } = await getJson(`/api/samples/${DEFAULT_SAMPLE_ID}`);
    if (!ok) {
      appendChatMessage("assistant", (data.errors || [data.error || "Sample fixture could not be loaded."]).join("\n"), { tone: "error" });
      return;
    }
    if (!state.profiles.length) await loadProfiles();
    state.profileId = data.profile_id || DEFAULT_PROFILE_ID;
    state.pricingReferenceId = data.pricing_reference_id || pricingReferenceForProfile(state.profileId)?.id || "";
    syncSelectedPricingReference();
    renderProfileOptions();
    renderPresetOptions();
    loadDefaultProfilePreset({ silent: true });
    updateGeneratorCopy();
    applyQuoteDetails(data.details || {}, { partial: true });
    state.images = Array.isArray(data.images) ? data.images : [];
    renderFiles();
    setWorkflowStage(state.images.length ? "ready_to_analyze" : "needs_images");
    appendChatMessage("assistant", `${data.label || "Sample"} loaded with ${state.images.length} reference image${state.images.length === 1 ? "" : "s"}.`);
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
    generator_label: generator.label,
    images: state.images,
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
    quote_basis: { ...state.quoteBasis },
    line_items: state.lineItems,
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
  const enabled = Boolean(file && file.url);
  elements.sideDownloadButton.classList.toggle("is-disabled", !enabled);
  elements.sideDownloadButton.setAttribute("aria-disabled", String(!enabled));
  elements.sideDownloadButton.tabIndex = enabled ? 0 : -1;
  elements.sideDownloadButton.href = enabled ? file.url : "#";
  elements.sideDownloadButton.download = enabled ? file.name || "quotation.xlsx" : "";
}

function pricingMatchStatus(row = {}) {
  const status = typeof row === "string" ? row : row.status;
  return String(status || "").trim().toLowerCase();
}

function pricingRowNeedsReview(row = {}) {
  const status = pricingMatchStatus(row);
  const amount = String(row.amount || "").trim();
  if (["unmatched", "ambiguous", "matched-from-ambiguous"].includes(status)) return true;
  return status === "manual-display" && (!amount || amount.toLowerCase() === "manual display price");
}

function pricingIssueForRow(row = {}) {
  if (!pricingRowNeedsReview(row)) return "";
  const description = row.description || "Pricing row";
  const status = pricingMatchStatus(row);
  if (status === "manual-display") return `Manual display pricing required: ${description} / enter a display price, choose a catalog keyword, or remove this line.`;
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
    unmatched: "Unmatched",
  };
  const normalized = pricingMatchStatus(status);
  return labels[normalized] || String(status || "").trim() || "Unknown";
}

function matchSummaryStats(rows = []) {
  const safeRows = Array.isArray(rows) ? rows : [];
  const total = safeRows.reduce((sum, row) => {
    const amount = Number(String(row.amount || "").replaceAll(",", ""));
    return Number.isFinite(amount) ? sum + amount : sum;
  }, 0);
  const confident = safeRows.filter((row) => pricingMatchStatus(row) === "matched").length;
  const needsReview = safeRows.filter((row) => pricingMatchStatus(row) !== "matched").length;
  const confidence = safeRows.length > 0 ? Math.round((confident / safeRows.length) * 100) : 0;
  return { total, confident, needsReview, confidence };
}

function renderMatchSummary(result = {}) {
  const rows = result.pricing_matches || [];
  if (!rows.length) {
    elements.matchSummary.innerHTML = "";
    return;
  }
  const stats = matchSummaryStats(rows);
  elements.matchSummary.innerHTML = `
    <div class="stat-card-row">
      <div class="stat-card">
        <span class="stat-card-icon green" aria-hidden="true">&#x1F4CB;</span>
        <span class="stat-card-value">${rows.length}</span>
        <span class="stat-card-label">Quote lines</span>
      </div>
      <div class="stat-card">
        <span class="stat-card-icon amber" aria-hidden="true">&#x1F3AF;</span>
        <span class="stat-card-value">${stats.confidence}%</span>
        <span class="stat-card-label">Catalog confidence</span>
      </div>
      <div class="stat-card">
        <span class="stat-card-icon red" aria-hidden="true">&#x26A0;</span>
        <span class="stat-card-value">${stats.needsReview}</span>
        <span class="stat-card-label">Needs review</span>
      </div>
      <div class="stat-card">
        <span class="stat-card-icon blue" aria-hidden="true">&#x1F4B0;</span>
        <span class="stat-card-value">SGD ${stats.total.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</span>
        <span class="stat-card-label">Total (excl. GST)</span>
      </div>
    </div>
  `;
}

function renderPricingMatches(rows = []) {
  state.pricingMatches = Array.isArray(rows) ? rows : [];
  elements.pricingTableWrap.hidden = !rows.length;
  elements.pricingEmptyState.hidden = Boolean(rows.length) || Boolean(elements.pricingReviewMessages.innerHTML.trim());
  if (!rows.length) {
    elements.pricingMatchesBody.innerHTML = `<tr><td colspan="7">No pricing matches yet.</td></tr>`;
    return;
  }
  elements.pricingMatchesBody.innerHTML = rows
    .map((row) => `
      <tr>
        <td title="${escapeHtml(row.status)}">${escapeHtml(pricingStatusLabel(row.status))}</td>
        <td>${escapeHtml(row.section)}</td>
        <td>${escapeHtml(row.description)}</td>
        <td>${escapeHtml(row.pricing_id || "")}</td>
        <td>${escapeHtml(row.quantity)}</td>
        <td>${escapeHtml(row.unit_price)}</td>
        <td>${escapeHtml(row.amount)}</td>
      </tr>
    `)
    .join("");
}

function clearPricingReviewMessages() {
  state.pricingIssues = [];
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
    appendChatMessage("assistant", "I could not map that pricing issue back to a generated line item. Regenerate analysis and try again.", { tone: "warn" });
    return;
  }

  const item = state.lineItems[index];
  if (action === "remove_line") {
    const [removed] = state.lineItems.splice(index, 1);
    appendChatMessage("assistant", `Removed "${removed.description}" from the quotation. Checking pricing again.`);
    await handleGenerate();
    return;
  }

  if (action === "manual_price") {
    const manualPrice = window.prompt("Manual display price", item.display_price || "");
    if (!manualPrice || !manualPrice.trim()) return;
    item.display_price = manualPrice.trim();
    appendChatMessage("assistant", `Set a manual display price for "${item.description}". Checking pricing again.`);
    await handleGenerate();
    return;
  }

  if (action === "nearest_keyword") {
    const keyword = window.prompt("Pricing keyword to try", item.pricing_keyword || item.description);
    if (!keyword || !keyword.trim()) return;
    item.pricing_keyword = keyword.trim();
    item.display_price = "";
    appendChatMessage("assistant", `Updated the pricing keyword for "${item.description}". Checking pricing again.`);
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

function appendChatMessage(role, content, options = {}) {
  state.chatMessages.push({
    role,
    content,
    html: Boolean(options.html),
    tone: options.tone || "",
  });
  renderChat();
  if (!elements.chatTranscript && ["assistant", "system"].includes(role)) {
    renderWorkflowNotice(content, options);
  }
}

function renderChat() {
  if (!elements.chatTranscript) return;
  elements.chatTranscript.innerHTML = state.chatMessages
    .map((message) => {
      const label = message.role === "user" ? "You" : "Assistant";
      const body = message.html ? message.content : renderPlainText(message.content);
      return `
        <div class="chat-message ${escapeHtml(message.role)} ${escapeHtml(message.tone)}">
          <div class="chat-meta">${label}</div>
          <div class="chat-body">${body}</div>
        </div>
      `;
    })
    .join("");
  const lastMessage = state.chatMessages.at(-1);
  if (lastMessage?.html && String(lastMessage.content || "").includes("quote-basis-card")) {
    const latestBasis = elements.chatTranscript.querySelector(".chat-message:last-child");
    elements.chatTranscript.scrollTop = Math.max(0, (latestBasis?.offsetTop || 0) - 118);
    return;
  }
  elements.chatTranscript.scrollTop = elements.chatTranscript.scrollHeight;
}

function noticeTextFromMessage(content, options = {}) {
  const text = options.html
    ? String(content || "").replace(/<[^>]*>/g, " ")
    : String(content || "");
  return normalizeTextNewlines(decodeHtmlEntities(text)).replace(/\s+/g, " ").trim();
}

function renderWorkflowNotice(content, options = {}) {
  if (!elements.workflowNotice) return;
  const text = noticeTextFromMessage(content, options);
  if (!text) return;
  const tone = options.tone || "instruction";
  const title = tone === "error" ? "Action needed" : tone === "warn" ? "Check before continuing" : "Workflow update";
  elements.workflowNotice.hidden = false;
  elements.workflowNotice.classList.remove("error", "warn", "instruction");
  elements.workflowNotice.classList.add(tone);
  elements.workflowNotice.innerHTML = `<strong>${escapeHtml(title)}</strong><span>${escapeHtml(text)}</span>`;
}

function renderBasisEmptyState(message = "Load images, complete Customer and Quote Company, then start analysis to review the draft here.") {
  elements.basisReviewSurface.innerHTML = `
    <div class="basis-empty-state">
      <strong>Quotation basis draft</strong>
      <p>${escapeHtml(message)}</p>
    </div>
  `;
}

function basisStatusParts(message, tone = "") {
  const text = normalizeTextNewlines(message).replace(/\s+/g, " ").trim();
  if (!text) {
    return {
      title: tone === "error" ? "Action needed" : "Quote basis status",
      detail: "",
    };
  }
  const separator = text.indexOf(":");
  if (separator > 0 && separator <= 80) {
    return {
      title: text.slice(0, separator).replace(/[.]+$/g, ""),
      detail: text.slice(separator + 1).trim(),
    };
  }
  if (tone === "error") return { title: "Action needed", detail: text };
  if (tone === "warn") return { title: "Quote basis status", detail: text };
  return { title: "Quote basis status", detail: text };
}

function setBasisReviewStatus(message, tone = "") {
  const status = basisStatusParts(message, tone);
  elements.basisReviewSurface.innerHTML = `
    <div class="basis-status-card ${escapeHtml(tone)}">
      <strong>${escapeHtml(status.title)}</strong>
      ${status.detail ? `<p>${escapeHtml(status.detail)}</p>` : ""}
    </div>
  `;
}

function actionButton(action) {
  const className = action.primary ? "primary-button" : "secondary-button";
  return `
    <button class="${className}" type="button" data-chat-action="${escapeHtml(action.action)}" ${action.disabled ? "disabled" : ""}>
      ${escapeHtml(action.label)}
    </button>
  `;
}

function renderChatActions(actions = []) {
  if (!elements.chatActions) return;
  elements.chatActions.innerHTML = actions.map(actionButton).join("");
}

function renderCurrentActions() {
  const busy = state.isAnalysisRunning || state.isGenerating;
  const readyForAnalysis = canStartAnalysis();
  if (state.workflowStage === "basis_review") {
    renderChatActions([]);
    return;
  }
  if (state.workflowStage === "details_review") {
    renderChatActions([
      { label: "Regenerate Analysis", action: "regenerate", disabled: busy || !readyForAnalysis },
    ]);
    return;
  }
  if (state.workflowStage === "pricing_review") {
    renderChatActions([
      { label: "Regenerate Analysis", action: "regenerate", primary: true, disabled: busy || !readyForAnalysis },
    ]);
    return;
  }
  if (state.workflowStage === "completed") {
    renderChatActions([
      { label: "Revise Quote", action: "open_basis_chat", primary: true, disabled: busy || !readyForAnalysis },
      { label: "Regenerate Analysis", action: "regenerate", disabled: busy || !readyForAnalysis },
    ]);
    return;
  }
  renderChatActions([
    { label: currentGenerator().analyzeLabel, action: "analyze", primary: true, disabled: busy || !readyForAnalysis },
  ]);
}

function basisLines(value) {
  const lines = splitLines(value);
  return lines.length ? lines : ["Confirm: No detail generated yet."];
}

function basisLineMeta(line) {
  const match = String(line || "").match(/^(Include|Confirm|Exclude|Assumption|Note):\s*(.*)$/i);
  if (!match) {
    return { tag: "Detail", text: String(line || "") };
  }
  const rawTag = match[1][0].toUpperCase() + match[1].slice(1).toLowerCase();
  const tag = rawTag === "Note" ? "Assumption" : rawTag;
  return { tag, text: match[2] || "" };
}

function unresolvedConfirmLines(basis = state.quoteBasis) {
  return BASIS_FIELDS.flatMap(([key, label]) => basisLines(basis[key])
    .filter((line) => ["Confirm", "Assumption"].includes(basisLineMeta(line).tag))
    .map((line) => `${label}: ${basisLineMeta(line).text || line}`));
}

function basisConfirmBlockReason(basis = state.quoteBasis) {
  return unresolvedConfirmLines(basis).length
    ? "Resolve all Confirm or Assumption lines before confirming quotation basis."
    : "";
}

function basisTagLabel(tag = "") {
  return tag === "Include" ? "Matched" : tag;
}

function renderBasisLine(key, line) {
  const meta = basisLineMeta(line);
  const rawLine = escapeHtml(line);
  return `
    <li class="basis-line-row basis-line-${escapeHtml(meta.tag.toLowerCase())}">
      <span class="basis-line-icon" aria-hidden="true"></span>
      <span class="basis-line-pill">${escapeHtml(basisTagLabel(meta.tag))}</span>
      <span class="basis-line-text" title="${escapeHtml(meta.text)}">${escapeHtml(meta.text)}</span>
      <span class="basis-line-actions">
        <button class="basis-line-tag-button" type="button" data-basis-tag-field="${escapeHtml(key)}" data-basis-tag-line="${rawLine}" data-basis-tag="Include" aria-label="Mark this line as matched" title="Mark matched">&#x2713;</button>
        <button class="basis-line-tag-button" type="button" data-basis-tag-field="${escapeHtml(key)}" data-basis-tag-line="${rawLine}" data-basis-tag="Exclude" aria-label="Mark this line as excluded" title="Mark excluded">X</button>
        <button class="basis-line-tool" type="button" data-revise-field="${escapeHtml(key)}" data-revise-line="${rawLine}" aria-label="Revise this line" title="Revise this line">Re</button>
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
  const statusText = aiFailed ? "AI failed" : source === "edited" ? "Edited draft" : "Needs your confirmation";
  const summaryText = state.lineItems.length
    ? `${state.lineItems.length} priced line${state.lineItems.length === 1 ? "" : "s"}`
    : "Pricing draft pending";
  return `
    <div class="assistant-card quote-basis-card ${aiFailed ? "quote-basis-card-failed" : ""}">
      <div class="quote-basis-header">
        <div>
          <div class="quote-basis-title-row">
            <h3>Quote basis to confirm</h3>
            <span>${escapeHtml(statusText)}</span>
          </div>
          <p>${aiFailed ? "AI analysis failed. Try again later. A local starter draft is shown for reference only." : "Please review the AI takeoff and confirm, revise a line, or request changes."}</p>
        </div>
        <div class="quote-basis-source">
          <span>${aiFailed ? "Source: Local fallback only" : "Source: Koncept Pricing Catalog"}</span>
          <strong>${escapeHtml(summaryText)}</strong>
        </div>
      </div>
      ${renderBasisTagLegend()}
      <div class="basis-review-grid">
        ${BASIS_FIELDS.map(([key, label]) => `
          <div class="basis-review-item">
            <h4>${escapeHtml(label)}</h4>
            <ul class="basis-line-list">
              ${basisLines(basis[key]).map((line) => renderBasisLine(key, line)).join("")}
            </ul>
          </div>
        `).join("")}
      </div>
    </div>
  `;
}

function basisFieldLabel(field) {
  return BASIS_FIELDS.find(([key]) => key === field)?.[1] || "Quote basis";
}

function currentQuoteBasisCard() {
  const cards = Array.from(elements.basisReviewSurface.querySelectorAll(".quote-basis-card"));
  return cards.at(-1) || null;
}

function updateQuoteBasisCard(source = "edited") {
  elements.basisReviewSurface.innerHTML = renderQuoteBasisMessage(state.quoteBasis, source);
}

function retagBasisLine(field, line, nextTag) {
  if (!field || !["Include", "Exclude"].includes(nextTag)) return;
  const currentLines = splitLines(state.quoteBasis[field]);
  const currentIndex = currentLines.findIndex((item) => item === line);

  const meta = basisLineMeta(line);
  const text = meta.tag === "Detail" ? line : meta.text;
  if (currentIndex >= 0) {
    currentLines[currentIndex] = `${nextTag}: ${text}`;
  } else if (!currentLines.length && meta.tag === "Confirm") {
    currentLines.push(`${nextTag}: ${text}`);
  } else {
    return;
  }
  state.quoteBasis[field] = currentLines.join("\n");
  state.basisConfirmed = false;
  updateQuoteBasisCard("edited");
  syncControlStates();
}

function basisChatIntroMessage() {
  if (state.basisChat.scope === "line") {
    return "Type a replacement like 200mm, or describe the change you want. I will show a proposed update before applying anything.";
  }
  return "Ask a question or describe changes to the quotation basis. I will show a proposed update before applying anything.";
}

function appendBasisChatMessage(role, text) {
  const message = document.createElement("div");
  message.className = `basis-chat-message ${role}`;
  message.innerHTML = `
    <span>${role === "user" ? "You" : "Assistant"}</span>
    ${renderPlainText(text)}
  `;
  elements.basisChatMessages.appendChild(message);
  elements.basisChatMessages.scrollTop = elements.basisChatMessages.scrollHeight;
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
  const nextBasis = proposal?.quoteBasis || {};
  return BASIS_FIELDS
    .filter(([key]) => normalizeTextNewlines(nextBasis[key]) !== normalizeTextNewlines(state.quoteBasis[key]))
    .map(([, label]) => label);
}

function proposalLinePreview(proposal) {
  if (state.basisChat.scope !== "line" || !state.basisChat.field) return "";
  const currentLines = splitLines(state.quoteBasis[state.basisChat.field]);
  const proposedLines = splitLines(proposal.quoteBasis[state.basisChat.field]);
  const index = currentLines.findIndex((line) => line === state.basisChat.line);
  if (index < 0 || proposedLines[index] === currentLines[index]) return "";
  return proposedLines[index] || "";
}

function setBasisChatProposal(proposal) {
  state.basisChat.proposal = proposal;
  const changedFields = proposalChangedFields(proposal);
  const linePreview = proposalLinePreview(proposal);
  elements.basisChatProposal.hidden = false;
  elements.basisChatProposal.innerHTML = `
    <strong>Proposed update</strong>
    <p>${escapeHtml(proposal.message || "Review this proposed quote basis change before applying it.")}</p>
    ${linePreview ? `<div class="basis-chat-proposed-line"><span>Current quote text</span><p>${escapeHtml(state.basisChat.line)}</p><span>Proposed quote text</span><p>${escapeHtml(linePreview)}</p></div>` : ""}
    <p>${escapeHtml(changedFields.length ? `Changes: ${changedFields.join(", ")}` : "This updates quotation line items without changing visible basis text.")}</p>
  `;
  elements.basisChatProposalActions.hidden = false;
  elements.basisChatApplyButton.disabled = false;
  elements.basisChatKeepButton.disabled = false;
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
    field: options.field || "",
    line: options.line || "",
    proposal: null,
  };
  resetBasisChatProposal();
  elements.basisChatTitle.textContent = scope === "line" ? "Revise basis line" : "Ask for changes";
  elements.basisChatContext.classList.toggle("has-selected-line", scope === "line");
  if (scope === "line") {
    const meta = basisLineMeta(state.basisChat.line);
    elements.basisChatContext.innerHTML = `
      <span class="basis-chat-context-label">${escapeHtml(basisFieldLabel(state.basisChat.field))}</span>
      <span class="basis-chat-selected-line">
        <strong>${escapeHtml(basisTagLabel(meta.tag))}</strong>
        <span>${escapeHtml(meta.text || state.basisChat.line)}</span>
      </span>
    `;
    elements.basisChatPrompt.placeholder = "Type 200mm, exclude this item, or describe a change...";
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

function basisChatRevisionText(text) {
  if (state.basisChat.line) {
    return `> ${state.basisChat.line}\n\n${text}`;
  }
  return text;
}

function wantsBasisRevision(normalizedText) {
  return /\b(add|change|delete|exclude|include|make|omit|remove|replace|revise|switch|turn|update|use)\b/.test(normalizedText);
}

function wantsBasisExplanation(normalizedText) {
  return /\b(what|why|explain|meaning|mean|clarify|question)\b/.test(normalizedText);
}

function basisExplanationText() {
  if (state.basisChat.scope === "line") {
    return `AI was not used for this reply. This line sits under ${basisFieldLabel(state.basisChat.field)}. It is part of the quotation basis that must be confirmed before Excel generation. If the wording is wrong, describe the change and I will draft a replacement for you to apply.`;
  }
  return "AI was not used for this reply. The quotation basis is the customer-facing checklist of what the draft quote includes, excludes, assumes, or still needs confirmed. Changes here should be reviewed before generating the Excel quotation.";
}

function extractLineReplacementText(text) {
  const match = String(text || "").match(/\b(?:change|replace|make|update|revise|switch)(?:\s+(?:this|it|line|item))?\s+(?:to|into|as)\s+([\s\S]+)$/i);
  const fallback = String(text || "").match(/^\s*(?:to|as)\s+([\s\S]+)$/i);
  const value = (match?.[1] || fallback?.[1] || "").trim().replace(/^["']|["']$/g, "");
  return value;
}

function directLineRevisionTarget(text, normalizedText) {
  const compact = String(text || "").trim().replace(/^["']|["']$/g, "");
  if (!compact || wantsBasisExplanation(normalizedText) || /\?$/.test(compact)) return "";
  const extracted = extractLineReplacementText(compact);
  if (extracted) return extracted;
  if (wantsBasisRevision(normalizedText)) return "";
  const wordCount = compact.split(/\s+/).filter(Boolean).length;
  if (wordCount > 6 || compact.length > 80) return "";
  return compact;
}

function sentenceCaseLine(value) {
  const text = String(value || "").trim();
  if (!text) return "";
  const punctuated = /[.!?]$/.test(text) ? text : `${text}.`;
  return `${punctuated[0].toUpperCase()}${punctuated.slice(1)}`;
}

function replaceMeasurementToken(currentLine, target) {
  const compact = String(target || "").trim();
  const targetMatch = compact.match(/^\d+(?:\.\d+)?\s*(mm|cm|m|sqm)$/i);
  if (!targetMatch) return "";
  const meta = basisLineMeta(currentLine);
  const tag = meta.tag === "Detail" ? "" : meta.tag;
  const body = meta.tag === "Detail" ? String(currentLine || "") : meta.text;
  const unit = targetMatch[1].toLowerCase();
  const sameUnitPattern = new RegExp(`\\b\\d+(?:\\.\\d+)?\\s*${unit}\\b`, "i");
  if (!sameUnitPattern.test(body)) return "";
  const nextBody = body.replace(sameUnitPattern, compact);
  return tag ? `${tag}: ${nextBody}` : nextBody;
}

function lineReplacementText(target, field, tag, currentLine = "") {
  if (/^(Include|Confirm|Exclude|Assumption|Note):/i.test(target)) return target.replace(/^Note:/i, "Assumption:");
  const compact = target.trim();
  const measurementReplacement = replaceMeasurementToken(currentLine, compact);
  if (measurementReplacement) return measurementReplacement;
  const isShortToken = /^[A-Za-z0-9+./-]{2,12}$/.test(compact);
  if (tag === "Confirm" && isShortToken) {
    const fieldText = basisFieldLabel(field).toLowerCase().replace(/\s*\/\s*/g, " ");
    return `Confirm: Please confirm ${compact} ${fieldText} requirement.`;
  }
  return `${tag}: ${sentenceCaseLine(compact)}`;
}

function draftLineTextRevision(text, normalizedText, basis = state.quoteBasis, lineItems = state.lineItems) {
  if (state.basisChat.scope !== "line" || !state.basisChat.field || !state.basisChat.line) return null;
  const target = directLineRevisionTarget(text, normalizedText);
  if (!target) return null;

  const currentLines = splitLines(basis[state.basisChat.field]);
  const currentIndex = currentLines.findIndex((line) => line === state.basisChat.line);
  const meta = basisLineMeta(state.basisChat.line);
  const tag = meta.tag === "Detail" ? "Confirm" : meta.tag;
  const nextLine = lineReplacementText(target, state.basisChat.field, tag, state.basisChat.line);
  const nextLines = currentLines.length ? [...currentLines] : [state.basisChat.line];
  if (currentIndex >= 0) {
    nextLines[currentIndex] = nextLine;
  } else {
    nextLines.push(nextLine);
  }
  const nextBasis = cloneQuoteBasis(basis);
  nextBasis[state.basisChat.field] = nextLines.join("\n");
  return {
    message: "Drafted a visible replacement for this basis line.",
    quoteBasis: nextBasis,
    lineItems: lineItems.map(normalizeLineItem),
  };
}

function buildRevisionProposal(text, normalizedText) {
  const contextualText = basisChatRevisionText(text);
  const contextualNormalized = contextualText.toLowerCase();
  return draftColorRevision(contextualText, contextualNormalized)
    || draftFlooringRevision(normalizedText)
    || draftLineTextRevision(text, normalizedText);
}

async function buildAiRevisionProposal(text) {
  if (!canStartAnalysis() || state.aiFailed) return null;
  const previousFeedback = state.pendingFeedback;
  const previousRunning = state.isAnalysisRunning;
  state.pendingFeedback = basisChatRevisionText(text);
  state.isAnalysisRunning = true;
  setBasisChatBusy(true);
  syncControlStates();
  appendBasisChatMessage("assistant", "Drafting a proposed update from the current quote basis.");
  try {
    const started = await startJob("draft", buildPayload());
    if (!started.ok) {
      appendBasisChatMessage("assistant", (started.data.errors || ["I could not draft that proposed change yet."]).join("\n"));
      return null;
    }
    const polled = await pollJob(started.data.job_id, (job) => {
      setBusyText(job.status === "running" ? "Drafting proposal..." : "Queued...");
    });
    const data = polled.data.result || {};
    if (!polled.ok || ["blocked", "failed"].includes(polled.data.status) || data.ai_failed) {
      appendBasisChatMessage("assistant", (data.errors || polled.data.errors || ["I could not draft that proposed change yet."]).join("\n"));
      return null;
    }
    return {
      message: "AI drafted an updated quote basis from your request.",
      quoteBasis: cloneQuoteBasis(data.quote_basis || state.quoteBasis),
      lineItems: Array.isArray(data.line_items) ? data.line_items.map(normalizeLineItem) : state.lineItems.map(normalizeLineItem),
    };
  } finally {
    state.pendingFeedback = previousFeedback;
    state.isAnalysisRunning = previousRunning;
    setBusyText("");
    setBasisChatBusy(false);
    syncControlStates();
  }
}

function basisChatPayload(text) {
  return {
    ...buildPayload(),
    basis_chat: {
      question: text,
      scope: state.basisChat.scope,
      field: state.basisChat.field,
      line: state.basisChat.line,
    },
  };
}

async function buildAiBasisChatAnswer(text) {
  if (!canStartAnalysis()) return null;
  const previousRunning = state.isAnalysisRunning;
  state.isAnalysisRunning = true;
  setBasisChatBusy(true);
  syncControlStates();
  appendBasisChatMessage("assistant", "Checking the selected basis.");
  try {
    const started = await startJob("basis_chat", basisChatPayload(text));
    if (!started.ok) {
      appendBasisChatMessage("assistant", (started.data.errors || ["I could not answer that yet."]).join("\n"));
      return null;
    }
    const polled = await pollJob(started.data.job_id, (job) => {
      setBusyText(job.status === "running" ? "Answering basis question..." : "Queued...");
    });
    const data = polled.data.result || {};
    if (!polled.ok || ["blocked", "failed"].includes(polled.data.status)) {
      appendBasisChatMessage("assistant", (data.errors || polled.data.errors || ["I could not answer that yet."]).join("\n"));
      return null;
    }
    return data.answer || null;
  } finally {
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

  const localProposal = buildRevisionProposal(text, normalized);
  if (localProposal) {
    appendBasisChatMessage("assistant", "AI was not used for this draft. I made a local proposed change from the selected line. Review it first, then apply it if it looks right.");
    setBasisChatProposal(localProposal);
    return;
  }

  if (wantsBasisRevision(normalized)) {
    const aiProposal = await buildAiRevisionProposal(text);
    if (aiProposal) {
      appendBasisChatMessage("assistant", "I drafted a proposed change. Review it first, then apply it if it looks right.");
      setBasisChatProposal(aiProposal);
      return;
    }
  }

  if (wantsBasisExplanation(normalized) || !wantsBasisRevision(normalized)) {
    const aiAnswer = await buildAiBasisChatAnswer(text);
    appendBasisChatMessage("assistant", aiAnswer || basisExplanationText());
    return;
  }

  appendBasisChatMessage("assistant", "I can answer questions about the basis or draft a proposed change. For edits, try a specific request like \"change floor finish to laminate\" or \"exclude this item\".");
}

function applyBasisChatProposal() {
  const proposal = state.basisChat.proposal;
  if (!proposal) return;
  state.basisConfirmed = false;
  state.quoteBasis = cloneQuoteBasis(proposal.quoteBasis);
  state.lineItems = Array.isArray(proposal.lineItems) ? proposal.lineItems.map(normalizeLineItem) : [];
  updateQuoteBasisCard("edited");
  resetBasisChatProposal();
  appendBasisChatMessage("assistant", "Change applied to the quote basis. Review the updated basis before confirming the quotation basis.");
  closeBasisChatOverlay();
  syncControlStates();
}

function keepCurrentBasis() {
  resetBasisChatProposal();
  appendBasisChatMessage("assistant", "Kept the current quote basis unchanged.");
  closeBasisChatOverlay();
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
  return ["basis_review", "generating", "pricing_review", "completed"].includes(state.workflowStage)
    && (state.lineItems.length > 0 || Object.values(state.quoteBasis).some((value) => splitLines(value).length > 0));
}

function hasCompletedQuoteBasis() {
  return state.basisConfirmed || ["pricing_review", "completed"].includes(state.workflowStage);
}

function applyDraftBasis(basis = {}) {
  state.basisConfirmed = false;
  state.quoteBasis = cloneQuoteBasis(basis);
}

function applyDraftLineItems(lineItems = []) {
  state.basisConfirmed = false;
  state.lineItems = lineItems.map(normalizeLineItem);
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
    logClientEvent("client_error", { url, message: error.message || String(error) });
    return {
      ok: false,
      data: {
        status: "failed",
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

async function getJson(url) {
  let response;
  try {
    response = await fetch(url);
  } catch (error) {
    logClientEvent("client_error", { url, message: error.message || String(error) });
    return {
      ok: false,
      data: {
        status: "failed",
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
  while (jobId) {
    const { ok, data } = await getJson(`/api/jobs/${encodeURIComponent(jobId)}`);
    if (!ok) return { ok, data };
    if (typeof onStatus === "function") onStatus(data);
    if (FINAL_JOB_STATUSES.has(data.status)) return { ok: true, data };
    await delay(900);
  }
  return { ok: false, data: { status: "failed", errors: ["Job was not created."] } };
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

function setAnalysisButtons(disabled = false) {
  const analysisBlockReason = startAnalysisBlockReason();
  const readyForAnalysis = !analysisBlockReason;
  const busy = state.isAnalysisRunning || state.isGenerating || disabled;
  const inputDisabled = busy || !readyForAnalysis;
  if (!elements.chatPrompt || !elements.sendChatButton) return;
  elements.chatPrompt.disabled = inputDisabled;
  elements.sendChatButton.disabled = inputDisabled;
  if (!readyForAnalysis) {
    elements.chatPrompt.placeholder = state.images.length
      ? "Fill Customer and Quote Company details to enable assistant replies."
      : "Drop reference images and fill Customer and Quote Company details to enable assistant replies.";
    return;
  }
  if (state.workflowStage === "completed" || state.workflowStage === "pricing_review") {
    elements.chatPrompt.placeholder = "Type a revision, e.g. change carpet to laminate, or ask me to regenerate analysis.";
    return;
  }
  if (state.workflowStage === "basis_review") {
    elements.chatPrompt.placeholder = "Reply confirm, regenerate, or describe a change to the draft basis.";
    return;
  }
  elements.chatPrompt.placeholder = "Reply with confirm, regenerate, sample details, or generate.";
}

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

  if (!state.images.length) {
    setWorkflowStage("needs_images");
    appendChatMessage("assistant", `Please drop at least one ${currentGenerator().imageNoun} before analysis.`, { tone: "warn" });
    renderCurrentActions();
    syncControlStates();
    return;
  }
  const missing = missingDetailFields();
  if (missing.length) {
    setWorkflowStage("details_review");
    appendChatMessage("assistant", `Fill Customer and Quote Company before AI analysis: ${missing.join(", ")}.`, { tone: "warn" });
    setDetailsDrawer(true);
    syncControlStates();
    return;
  }

  state.isAnalysisRunning = true;
  state.aiFailed = false;
  state.draftSource = "";
  clearAiFailureBanner();
  setWorkflowStage("analyzing");
  showAiRunningBanner();
  setBasisReviewStatus(
    hasFeedback
      ? "Using your feedback to revise the AI takeoff now. I will return an updated basis for confirmation."
      : "Analyzing reference images now. I will list the basis for confirmation before generating anything.",
    "warn"
  );
  setSidePanel("basis", { force: true });
  setBusyText("Running analysis...");
  setAnalysisButtons(true);
  renderChatActions([{ label: "Analyzing...", action: "noop", disabled: true }]);

  const started = await startJob("draft", buildPayload({
    includeBoothDimensions: state.boothDimensions.dimension_source !== "default",
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
      setBasisReviewStatus(errors.join("\n"), "warn");
    } else {
      showAiFailureBanner("Try again later.");
      setBasisReviewStatus(errors.join("\n"), "error");
    }
    appendChatMessage("assistant", errors.join("\n"), { tone: wasBlocked ? "warn" : "error" });
    syncControlStates();
    return;
  }
  state.activeJob = { id: started.data.job_id, type: "draft" };
  saveSessionState();

  const polled = await pollJob(started.data.job_id, (job) => {
    setBusyText(job.status === "running" ? "Running analysis..." : "Queued...");
  });
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
      setBasisReviewStatus(errors.join("\n"), "warn");
    } else {
      showAiFailureBanner("Try again later.");
      setBasisReviewStatus(errors.join("\n"), "error");
    }
    appendChatMessage("assistant", errors.join("\n"), { tone: wasBlocked ? "warn" : "error" });
    syncControlStates();
    return;
  }

  const data = polled.data.result || {};
  applyDraftBasis(data.quote_basis || {});
  applyDraftLineItems(data.line_items || []);
  state.boothDimensions = normalizeBoothDimensions(data.project || state.boothDimensions);
  state.draftSource = data.source || "";
  state.aiFailed = Boolean(data.ai_failed || polled.data.status === "degraded" || (data.source === "local" && Array.isArray(data.warnings) && data.warnings.length));
  setWorkflowStage("basis_review");
  const warningText = Array.isArray(data.warnings) && data.warnings.length ? data.warnings.join(" ") : "";
  if (state.aiFailed) {
    showAiFailureBanner(warningText || "Try again later. A local fallback draft is shown for reference only.");
    appendChatMessage("assistant", "AI analysis failed. Try again later. I showed a local fallback draft, but it is blocked from quotation generation until remote AI analysis succeeds.", { tone: "error" });
  } else {
    clearAiFailureBanner();
  }
  if (Array.isArray(data.warnings) && data.warnings.length) {
    appendChatMessage("assistant", data.warnings.join("\n"), { tone: "warn" });
  }
  updateQuoteBasisCard(data.source);
  setSidePanel("basis", { force: true });
  if (!state.lineItems.length) {
    appendChatMessage("assistant", EMPTY_LINE_ITEMS_MESSAGE, { tone: "warn" });
  }
  syncControlStates();
}

async function confirmBasis() {
  if (state.isAnalysisRunning || state.isGenerating) return;
  const confirmBlockReason = basisConfirmBlockReason();
  if (confirmBlockReason) {
    setWorkflowStage("basis_review");
    appendChatMessage("assistant", confirmBlockReason, { tone: "warn" });
    syncControlStates();
    return;
  }
  if (state.aiFailed) {
    showAiFailureBanner("Try again later. Regenerate analysis before confirming the quote basis.");
    appendChatMessage("assistant", "AI analysis failed. Try again later before confirming or generating the quotation.", { tone: "error" });
    renderCurrentActions();
    return;
  }
  if (!state.lineItems.length) {
    setWorkflowStage("basis_review");
    appendChatMessage("assistant", "I do not have any generated quotation line items yet. Regenerate analysis before confirming the basis.", { tone: "warn" });
    renderCurrentActions();
    return;
  }
  const missing = missingDetailFields();
  if (missing.length) {
    setWorkflowStage("details_review");
    appendChatMessage("assistant", `Please fill these details before I generate Excel: ${missing.join(", ")}.`, { tone: "warn" });
    setDetailsDrawer(true);
    syncControlStates();
    return;
  }
  state.basisConfirmed = true;
  appendChatMessage("assistant", "Basis confirmed. Checking pricing before creating the final output.");
  await handleGenerate();
}

async function handleGenerate() {
  if (state.isGenerating) return;
  if (!state.basisConfirmed) {
    if (hasSubmittedQuoteBasis()) {
      setWorkflowStage("basis_review");
      setSidePanel("basis", { force: true });
      appendChatMessage("assistant", "Confirm Quotation Basis before generating the Excel quotation.", { tone: "warn" });
    }
    syncControlStates();
    return;
  }
  if (state.aiFailed) {
    setWorkflowStage("basis_review");
    showAiFailureBanner("Try again later. Regenerate analysis before generating Excel.");
    appendChatMessage("assistant", "Cannot generate because AI analysis failed. Try again later or regenerate analysis first.", { tone: "error" });
    syncControlStates();
    return;
  }
  const missing = missingDetailFields();
  if (missing.length) {
    setWorkflowStage("details_review");
    appendChatMessage("assistant", `Please fill these details before I generate Excel: ${missing.join(", ")}.`, { tone: "warn" });
    setDetailsDrawer(true);
    syncControlStates();
    return;
  }
  if (!state.lineItems.length) {
    setWorkflowStage("basis_review");
    appendChatMessage("assistant", "There are no generated line items yet. Run analysis first so Excel has quotation rows.", { tone: "warn" });
    renderCurrentActions();
    return;
  }

  state.isGenerating = true;
  setWorkflowStage("generating");
  setBusyText("Checking pricing...");
  setResultStatus("Checking pricing", "is-warn");
  if (elements.chatPrompt) elements.chatPrompt.value = "";
  renderMessages([]);
  setDownloadFiles([]);
  renderMatchSummary({});
  renderPricingMatches([]);
  clearPricingReviewMessages();
  syncControlStates();
  const started = await startJob("generate", buildPayload());
  if (!started.ok) {
    state.isGenerating = false;
    setBusyText("");
    setWorkflowStage("details_review");
    setResultStatus(started.data.status || "Failed", "is-bad");
    renderMessages(started.data.errors || ["Generation failed."], "error");
    appendChatMessage("assistant", (started.data.errors || ["Generation failed."]).join("\n"), { tone: "error" });
    syncControlStates();
    return;
  }
  state.activeJob = { id: started.data.job_id, type: "generate" };
  saveSessionState();

  const polled = await pollJob(started.data.job_id, (job) => {
    setBusyText(job.status === "running" ? "Checking pricing..." : "Queued...");
  });
  state.isGenerating = false;
  state.activeJob = null;
  setBusyText("");

  const data = polled.data.result || polled.data || {};
  if (!polled.ok || ["blocked", "failed"].includes(polled.data.status) || data.status === "blocked" || data.status === "failed") {
    setWorkflowStage("details_review");
    setResultStatus(data.status || "Failed", "is-bad");
    renderMessages(data.errors || ["Generation failed."], "error");
    renderPricingMatches(data.pricing_matches || []);
    renderMatchSummary(data);
    appendChatMessage("assistant", (data.errors || ["Generation failed."]).join("\n"), { tone: "error" });
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
    setSidePanel("pricing");
    setDownloadFiles([]);
    appendChatMessage("assistant", "I found pricing items that need review. I have shown them in Pricing instead of creating the final output.", { tone: "warn" });
  } else {
    setWorkflowStage("completed");
    setResultStatus("Completed", "is-ok");
    renderMessages([]);
    clearPricingReviewMessages();
    setSidePanel("output");
    appendChatMessage("assistant", "Pricing is clear. The Excel quotation is ready in Output.", { tone: "instruction" });
    setDownloadFiles(data.files || []);
  }
  renderPricingMatches(data.pricing_matches || []);
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
    showAiRunningBanner("Resuming the analysis job after refresh.");
    setBasisReviewStatus("Resuming the saved analysis job.", "warn");
    setSidePanel("basis", { force: true });
    setBusyText("Checking saved analysis job...");
    syncControlStates();

    const polled = await pollJob(activeJob.id, (job) => {
      setBusyText(job.status === "running" ? "Running analysis..." : "Queued...");
    });
    state.isAnalysisRunning = false;
    state.activeJob = null;
    setBusyText("");
    setAnalysisButtons(false);

    if (!polled.ok || ["blocked", "failed"].includes(polled.data.status)) {
      setWorkflowStage("ready_to_analyze");
      showAiFailureBanner("Try again later.");
      setBasisReviewStatus((polled.data.errors || polled.data.result?.errors || ["Draft failed."]).join("\n"), "error");
      appendChatMessage("assistant", (polled.data.errors || polled.data.result?.errors || ["Draft failed."]).join("\n"), { tone: "error" });
      syncControlStates();
      return;
    }

    const data = polled.data.result || {};
    applyDraftBasis(data.quote_basis || {});
    applyDraftLineItems(data.line_items || []);
    state.draftSource = data.source || "";
    state.aiFailed = Boolean(data.ai_failed || polled.data.status === "degraded" || (data.source === "local" && Array.isArray(data.warnings) && data.warnings.length));
    setWorkflowStage("basis_review");
    if (state.aiFailed) {
      const warningText = Array.isArray(data.warnings) && data.warnings.length ? data.warnings.join(" ") : "";
      showAiFailureBanner(warningText || "Try again later. A local fallback draft is shown for reference only.");
      appendChatMessage("assistant", "AI analysis failed. Try again later. I showed a local fallback draft, but it is blocked from quotation generation until remote AI analysis succeeds.", { tone: "error" });
    } else {
      clearAiFailureBanner();
    }
    if (Array.isArray(data.warnings) && data.warnings.length) {
      appendChatMessage("assistant", data.warnings.join("\n"), { tone: "warn" });
    }
    updateQuoteBasisCard(data.source);
    setSidePanel("basis", { force: true });
    if (!state.lineItems.length) {
      appendChatMessage("assistant", EMPTY_LINE_ITEMS_MESSAGE, { tone: "warn" });
    }
    syncControlStates();
    return;
  }

  if (activeJob.type === "generate") {
    state.isGenerating = true;
    state.isAnalysisRunning = false;
    setWorkflowStage("generating");
    setBusyText("Checking saved pricing job...");
    setResultStatus("Checking pricing", "is-warn");
    setSidePanel("pricing", { force: true });
    syncControlStates();

    const polled = await pollJob(activeJob.id, (job) => {
      setBusyText(job.status === "running" ? "Checking pricing..." : "Queued...");
    });
    state.isGenerating = false;
    state.activeJob = null;
    setBusyText("");

    const data = polled.data.result || polled.data || {};
    if (!polled.ok || ["blocked", "failed"].includes(polled.data.status) || data.status === "blocked" || data.status === "failed") {
      setWorkflowStage("details_review");
      setResultStatus(data.status || "Failed", "is-bad");
      renderMessages(data.errors || ["Generation failed."], "error");
      renderPricingMatches(data.pricing_matches || []);
      renderMatchSummary(data);
      appendChatMessage("assistant", (data.errors || ["Generation failed."]).join("\n"), { tone: "error" });
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
      setSidePanel("pricing");
      setDownloadFiles([]);
      appendChatMessage("assistant", "I found pricing items that need review. I have shown them in Pricing instead of creating the final output.", { tone: "warn" });
    } else {
      setWorkflowStage("completed");
      setResultStatus("Completed", "is-ok");
      renderMessages([]);
      clearPricingReviewMessages();
      setSidePanel("output");
      appendChatMessage("assistant", "Pricing is clear. The Excel quotation is ready in Output.", { tone: "instruction" });
      setDownloadFiles(data.files || []);
    }
    renderPricingMatches(data.pricing_matches || []);
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

function handleChatAction(action) {
  if (action === "analyze" || action === "regenerate") {
    handleDraftBasis();
    return;
  }
  if (action === "confirm_basis") {
    confirmBasis();
    return;
  }
  if (action === "sample_details") {
    setSampleDetails();
    return;
  }
  if (action === "generate") {
    handleGenerate();
    return;
  }
  if (action === "open_basis_chat") {
    openBasisChatOverlay("quote");
  }
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

function estimatedAreaQuantity() {
  const dimensions = normalizeBoothDimensions(state.boothDimensions);
  const width = Number(dimensions.booth_width);
  const depth = Number(dimensions.booth_depth);
  if (!Number.isFinite(width) || !Number.isFinite(depth) || width <= 0 || depth <= 0) return "";
  return Math.round(width * depth * 100) / 100;
}

function findFloorFinishLineIndex(lineItems = state.lineItems) {
  const finishPattern = /(carpet|flooring|laminate|pvc)/i;
  const floorPattern = /(floor|platform)/i;
  const finishIndex = lineItems.findIndex((item) => finishPattern.test(`${item.section} ${item.description} ${item.pricing_keyword}`));
  if (finishIndex >= 0) return finishIndex;
  return lineItems.findIndex((item) => floorPattern.test(`${item.section} ${item.description} ${item.pricing_keyword}`));
}

function revisedBasisText(value, line) {
  const lines = splitLines(value);
  const targetIndex = lines.findIndex((item) => /(floor|platform|carpet|laminate|pvc)/i.test(item));
  if (targetIndex >= 0) {
    lines[targetIndex] = line;
  } else {
    lines.unshift(line);
  }
  return lines.join("\n");
}

const COLOR_WORDS = [
  "black",
  "blue",
  "brown",
  "cream",
  "gold",
  "gray",
  "green",
  "grey",
  "orange",
  "pink",
  "purple",
  "red",
  "silver",
  "white",
  "yellow",
];
const COLOR_WORD_PATTERN = "(black|blue|brown|cream|gold|gr[ae]y|green|orange|pink|purple|red|silver|white|yellow)";

function colorRegex(color) {
  return new RegExp(`\\b${color}\\b`, "gi");
}

function replaceColorWord(value, fromColor, toColor) {
  return String(value || "").replace(colorRegex(fromColor), (match) => {
    if (match === match.toUpperCase()) return toColor.toUpperCase();
    if (match[0] === match[0].toUpperCase()) return `${toColor[0].toUpperCase()}${toColor.slice(1)}`;
    return toColor;
  });
}

function normalizeColorWord(color) {
  return color === "grey" ? "gray" : color;
}

function quotedRevisionLines(text) {
  return normalizeTextNewlines(text)
    .split("\n")
    .map((line) => line.trim())
    .filter((line) => line.startsWith(">"))
    .map((line) => line.replace(/^>\s?/, "").trim())
    .filter(Boolean);
}

function colorWordsInText(text) {
  return Array.from(String(text || "").toLowerCase().matchAll(new RegExp(`\\b${COLOR_WORD_PATTERN}\\b`, "g")))
    .map((match) => normalizeColorWord(match[1]));
}

function findColorRevision(text, normalizedText) {
  const explicitMatch = normalizedText.match(new RegExp(`\\b(?:change|replace|switch|revise|update|make|turn)\\s+(?:the\\s+)?${COLOR_WORD_PATTERN}\\b[\\s\\S]{0,80}?\\b(?:to|into)\\s+${COLOR_WORD_PATTERN}\\b`));
  if (explicitMatch) {
    return {
      sourceColor: normalizeColorWord(explicitMatch[1]),
      targetColor: normalizeColorWord(explicitMatch[2]),
      quotedLines: quotedRevisionLines(text),
    };
  }

  const targetMatch = normalizedText.match(new RegExp(`\\b(?:change|replace|switch|revise|update|make|turn)\\b[\\s\\S]{0,80}?\\b(?:to|into)\\s+${COLOR_WORD_PATTERN}\\b`));
  if (!targetMatch) return null;

  const targetColor = normalizeColorWord(targetMatch[1]);
  const quotedLines = quotedRevisionLines(text);
  const quotedColors = colorWordsInText(quotedLines.join("\n")).filter((color) => color !== targetColor);
  if (/\bthis\b/.test(normalizedText) && quotedColors.length) {
    return { sourceColor: quotedColors[0], targetColor, quotedLines };
  }

  const beforeTarget = normalizedText.slice(0, targetMatch.index);
  const previousColors = colorWordsInText(beforeTarget)
    .filter((color) => color !== targetColor);
  const sourceColor = previousColors.at(-1);
  if (!sourceColor) return null;
  return { sourceColor, targetColor, quotedLines };
}

function draftColorRevision(text, normalizedText, basis = state.quoteBasis, lineItems = state.lineItems) {
  const revision = findColorRevision(text, normalizedText);
  if (!revision) return null;

  let changedCount = 0;
  const nextBasis = cloneQuoteBasis(basis);
  const quotedLines = revision.quotedLines || [];
  const quoteMatched = quotedLines.length
    ? Object.keys(nextBasis).some((key) => {
        let changedField = false;
        const nextLines = normalizeTextNewlines(nextBasis[key]).split("\n").map((line) => {
          const shouldEdit = quotedLines.some((quotedLine) => line.trim() === quotedLine || line.includes(quotedLine) || quotedLine.includes(line.trim()));
          if (!shouldEdit) return line;
          const nextLine = replaceColorWord(line, revision.sourceColor, revision.targetColor);
          if (nextLine !== line) {
            changedCount += 1;
            changedField = true;
          }
          return nextLine;
        });
        if (changedField) nextBasis[key] = nextLines.join("\n");
        return changedField;
      })
    : false;

  if (!quoteMatched) {
    Object.keys(nextBasis).forEach((key) => {
      const nextValue = replaceColorWord(nextBasis[key], revision.sourceColor, revision.targetColor);
      if (nextValue !== nextBasis[key]) changedCount += 1;
      nextBasis[key] = nextValue;
    });
  }

  const nextLineItems = lineItems.map((item) => {
    const nextDescription = replaceColorWord(item.description, revision.sourceColor, revision.targetColor);
    if (nextDescription !== item.description) changedCount += 1;
    return { ...item, description: nextDescription };
  });

  if (!changedCount) return null;
  return {
    message: `Updated ${changedCount} draft field${changedCount === 1 ? "" : "s"} from ${revision.sourceColor} to ${revision.targetColor}.`,
    quoteBasis: nextBasis,
    lineItems: nextLineItems,
  };
}

function applyColorRevision(text, normalizedText) {
  const proposal = draftColorRevision(text, normalizedText);
  if (!proposal) return null;
  state.basisConfirmed = false;
  state.quoteBasis = proposal.quoteBasis;
  state.lineItems = proposal.lineItems;
  return proposal.message;
}

function draftFlooringRevision(normalizedText, basis = state.quoteBasis, lineItems = state.lineItems) {
  if (!/laminat|carpet|floor|flooring|platform/.test(normalizedText)) return null;
  if (!/(change|replace|switch|revise|update|make|use|laminat)/.test(normalizedText)) return null;

  const woodGrain = /wood|timber|grain/.test(normalizedText);
  const description = woodGrain
    ? "Wood grain laminated flooring on raised platform"
    : "White laminated flooring on raised platform";
  const pricingKeyword = woodGrain
    ? "floor-design.wood-grain-laminated-flooring-on-raised-platform"
    : "floor-design.white-laminated-flooring-on-raised-platform";
  const nextItem = {
    section: "Floor Design",
    quantity: estimatedAreaQuantity(),
    unit: "sqm",
    description,
    pricing_keyword: pricingKeyword,
    display_price: "",
  };

  const nextBasis = cloneQuoteBasis(basis);
  const nextLineItems = lineItems.map((item) => ({ ...item }));
  const index = findFloorFinishLineIndex(nextLineItems);
  if (index >= 0) {
    nextLineItems[index] = { ...nextLineItems[index], ...nextItem };
  } else {
    nextLineItems.unshift(nextItem);
  }
  nextBasis.platform = revisedBasisText(
    nextBasis.platform,
    `Confirm: Platform flooring revised to ${description.toLowerCase()}.`
  );
  return {
    message: `Updated the floor finish to ${description.toLowerCase()} using catalog item ${pricingKeyword}.`,
    quoteBasis: nextBasis,
    lineItems: nextLineItems,
  };
}

function applyFlooringRevision(normalizedText) {
  const proposal = draftFlooringRevision(normalizedText);
  if (!proposal) return null;
  state.basisConfirmed = false;
  state.quoteBasis = proposal.quoteBasis;
  state.lineItems = proposal.lineItems;
  return proposal.message;
}

async function applyRevisionRequest(text, normalizedText) {
  const applied = applyColorRevision(text, normalizedText) || applyFlooringRevision(normalizedText);
  if (!applied) return false;

  if (state.workflowStage === "basis_review") {
    appendChatMessage("assistant", `${applied} Review the updated basis below, then confirm when it looks right.`);
    updateQuoteBasisCard("edited");
    syncControlStates();
    return true;
  }

  appendChatMessage("assistant", `${applied} I am regenerating the Excel quotation now.`);
  setResultStatus("Revision running", "is-warn");
  renderMessages(["Revision applied. Regenerating quotation."]);
  setDownloadFiles([]);
  renderMatchSummary({});
  renderPricingMatches([]);
  await handleGenerate();
  return true;
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
  if (panelName === "pricing") {
    const missing = missingDetailFields();
    if (missing.length) return `Complete Customer and Quote Company details before opening Pricing: ${missing.join(", ")}.`;
    if (!hasSubmittedQuoteBasis()) return "Click Start Analysis from Quote Company before opening Pricing.";
    const confirmBlockReason = basisConfirmBlockReason();
    if (confirmBlockReason) return confirmBlockReason;
    if (!hasCompletedQuoteBasis()) return "Confirm Quotation Basis before opening Pricing.";
  }
  if (panelName === "output") {
    const pricingBlockReason = sidePanelBlockReason("pricing");
    if (pricingBlockReason) return pricingBlockReason.replace("opening Pricing", "opening Output");
    const reviewBlockReason = pricingReviewBlockReason();
    if (reviewBlockReason) return reviewBlockReason;
    if (!state.downloadFile) return "Generate a priced quotation before opening Output.";
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
    pricing: ["Pricing", "Price Review", "Catalog matches, manual display prices, and unresolved pricing must be cleared before output."],
    output: ["Output", "Generated Quotation", "Final quotation status and Excel download appear here after pricing review clears."],
  };
  const nextPanel = panelTitles[panelName] ? panelName : "images";
  const blockReason = sidePanelBlockReason(nextPanel);
  if (blockReason && !options.force) {
    if (options.notify) appendChatMessage("assistant", blockReason, { tone: "warn" });
    updateSidePanelNav();
    return false;
  }
  const [title, eyebrow, subtitle] = panelTitles[nextPanel] || panelTitles.images;
  state.activeSidePanel = nextPanel;
  document.body.dataset.sidePanel = state.activeSidePanel;
  elements.sideDrawerTitle.textContent = title;
  elements.sideDrawerEyebrow.textContent = eyebrow;
  elements.sideDrawerSubtitle.textContent = subtitle || "";
  document.querySelectorAll("[data-side-panel]").forEach((button) => {
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
  const isPricingStep = state.activeSidePanel === "pricing";
  const isOutputStep = state.activeSidePanel === "output";
  const busy = state.isAnalysisRunning || state.isGenerating;
  elements.sampleDetailsButton.hidden = state.activeSidePanel !== "images";
  elements.sampleDetailsButton.disabled = state.isBooting || busy;
  elements.clearCustomerButton.hidden = state.activeSidePanel !== "customer";
  elements.clearQuoteCompanyButton.hidden = state.activeSidePanel !== "quote_company";
  elements.clearCustomerButton.disabled = busy;
  elements.clearQuoteCompanyButton.disabled = busy;
  document.querySelectorAll("[data-side-panel]").forEach((button) => {
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
    pricing: "Next: Output",
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
  } else if (isPricingStep) {
    nextBlockReason = pricingReviewBlockReason() || sidePanelBlockReason(nextPanel);
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
    appendChatMessage("assistant", reason, { tone: "warn" });
    return;
  }
  if (state.activeSidePanel === "quote_company") {
    handleDraftBasis();
    return;
  }
  if (state.activeSidePanel === "basis") {
    confirmBasis();
    return;
  }
  if (state.activeSidePanel === "pricing") {
    setSidePanel("output", { notify: true });
    return;
  }
  if (state.activeSidePanel === "output") return;
  setSidePanel(SIDE_PANEL_SEQUENCE[index + 1], { notify: true });
}

async function handleChatSubmit(event) {
  event.preventDefault();
  if (!elements.chatPrompt) return;
  const text = elements.chatPrompt.value.trim();
  if (!text) return;

  appendChatMessage("user", text);
  elements.chatPrompt.value = "";
  const normalized = text.toLowerCase();

  if (isSensitiveChatRequest(normalized)) {
    appendChatMessage("assistant", `I cannot access or reveal secrets, API keys, .env values, tokens, authorization headers, system prompts, or hidden instructions. I can still help ${currentGenerator().fallbackAction} and generate the Excel quote.`);
    renderCurrentActions();
    return;
  }

  if (normalized.includes("sample")) {
    setSampleDetails();
    return;
  }
  if (normalized.includes("regenerate") || normalized.includes("rerun") || normalized.includes("re-run")) {
    handleDraftBasis();
    return;
  }
  if (normalized.includes("analyze") || normalized.includes("analyse") || normalized.includes("analysis")) {
    handleDraftBasis();
    return;
  }
  if (normalized.includes("generate")) {
    handleGenerate();
    return;
  }
  if (["yes", "ok", "okay", "confirm", "confirmed", "proceed", "go ahead"].includes(normalized)) {
    if (state.workflowStage === "basis_review") {
      confirmBasis();
      return;
    }
    if (state.workflowStage === "details_review") {
      handleGenerate();
      return;
    }
  }

  if (await applyRevisionRequest(text, normalized)) {
    return;
  }

  if (canStartAnalysis() && ["basis_review", "completed", "pricing_review"].includes(state.workflowStage)) {
    state.pendingFeedback = text;
    await handleDraftBasis();
    state.pendingFeedback = "";
    return;
  }

  appendChatMessage(
    "assistant",
    `I could not apply that safely as a direct edit yet. Try a specific revision like "change floor finish to laminate", or regenerate analysis for a fresh AI takeoff. Customer and Quote Company still control customer, terms, notes, and signature text.`
  );
  renderCurrentActions();
}

function handleQuoteBasisClick(event) {
  const tagButton = event.target.closest("[data-basis-tag]");
  if (tagButton) {
    retagBasisLine(
      tagButton.dataset.basisTagField || "",
      tagButton.dataset.basisTagLine || "",
      tagButton.dataset.basisTag || ""
    );
    return;
  }
  const reviseButton = event.target.closest("[data-revise-line]");
  if (reviseButton) {
    openBasisChatOverlay("line", {
      field: reviseButton.dataset.reviseField || "",
      line: reviseButton.dataset.reviseLine || "",
    });
    return;
  }
  const button = event.target.closest("[data-chat-action]");
  if (!button || button.disabled) return;
  handleChatAction(button.dataset.chatAction);
}

function wireEvents() {
  wireRichTextEditors();
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

  document.querySelectorAll("[data-side-panel]").forEach((button) => {
    button.addEventListener("click", () => setSidePanel(button.dataset.sidePanel || "images", { notify: true }));
  });
  elements.sideBackButton.addEventListener("click", goToPreviousSidePanel);
  elements.sideNextButton.addEventListener("click", goToNextSidePanel);
  elements.sideDownloadButton.addEventListener("click", (event) => {
    if (elements.sideDownloadButton.getAttribute("aria-disabled") === "true") {
      event.preventDefault();
    }
  });
  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape") {
      if (!elements.basisChatOverlay.hidden) {
        closeBasisChatOverlay();
      }
    }
  });

  elements.sampleDetailsButton.addEventListener("click", setSampleDetails);
  elements.newQuoteButton.addEventListener("click", startNewQuote);
  elements.profileSelect.addEventListener("change", handleProfileSelectionChange);
  elements.presetSelect.addEventListener("change", () => {
    updatePresetButtons();
    renderPresetStatus();
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
  if (elements.chatForm) elements.chatForm.addEventListener("submit", handleChatSubmit);
  elements.basisChatForm.addEventListener("submit", handleBasisChatSubmit);
  elements.basisChatApplyButton.addEventListener("click", applyBasisChatProposal);
  elements.basisChatKeepButton.addEventListener("click", keepCurrentBasis);
  elements.basisChatCloseButton.addEventListener("click", closeBasisChatOverlay);
  elements.basisChatOverlay.addEventListener("click", (event) => {
    if (event.target.closest("[data-basis-chat-close]")) {
      closeBasisChatOverlay();
    }
  });
  if (elements.chatPrompt && elements.chatForm && elements.sendChatButton) {
    elements.chatPrompt.addEventListener("keydown", (event) => {
      if (event.key !== "Enter" || event.shiftKey) return;
      event.preventDefault();
      if (!elements.sendChatButton.disabled) {
        elements.chatForm.requestSubmit();
      }
    });
  }
  elements.basisChatPrompt.addEventListener("keydown", (event) => {
    if (event.key !== "Enter" || event.shiftKey) return;
    event.preventDefault();
    if (!elements.basisChatSendButton.disabled) {
      elements.basisChatForm.requestSubmit();
    }
  });
  if (elements.chatActions) {
    elements.chatActions.addEventListener("click", (event) => {
      const button = event.target.closest("[data-chat-action]");
      if (!button || button.disabled) return;
      handleChatAction(button.dataset.chatAction);
    });
  }
  elements.basisReviewSurface.addEventListener("click", handleQuoteBasisClick);
  if (elements.chatTranscript) elements.chatTranscript.addEventListener("click", handleQuoteBasisClick);
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
