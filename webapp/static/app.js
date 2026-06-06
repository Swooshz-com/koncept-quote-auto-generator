const EMPTY_LINE_ITEMS_MESSAGE = "AI analysis will populate line items here.";
const DEFAULT_PROFILE_ID = "koncept";
const DEFAULT_SAMPLE_ID = "brazil-pavilion";
const FINAL_JOB_STATUSES = new Set(["completed", "degraded", "needs_review", "blocked", "failed"]);

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

const STAGE_LABELS = {
  needs_images: "Needs images",
  ready_to_analyze: "Ready to analyze",
  analyzing: "Analyzing",
  basis_review: "Confirm basis",
  details_review: "Confirm details",
  generating: "Generating",
  pricing_review: "Pricing review",
  completed: "Completed",
};

const GENERATOR_TYPES = {
  booth: {
    label: "Exhibition Booth",
    assistantSubtitle: "Drop booth images, confirm the AI takeoff, then generate Excel.",
    intakeTitle: "Booth render images",
    intakeSubtitle: "Images are the starting point for this booth quote workflow.",
    dropTitle: "Drop booth render images to start",
    dropMeta: "JPG, PNG, or WebP. Add more booth renders anytime; files stay local to this runner.",
    imageNoun: "booth image",
    analyzeLabel: "Analyze Booth",
    basisSource: "I analysed the booth images",
    fallbackAction: "analyze booth images",
    widthLabel: "Booth width (m)",
    depthLabel: "Booth depth (m)",
    sizeLabel: "Booth",
  },
  event_setup: {
    label: "Event Setup",
    assistantSubtitle: "Drop event references, confirm the AI takeoff, then generate Excel.",
    intakeTitle: "Event reference images",
    intakeSubtitle: "Photos, layouts, and visual references are the starting point for this quote workflow.",
    dropTitle: "Drop event reference images to start",
    dropMeta: "JPG, PNG, or WebP. Add more event references anytime; files stay local to this runner.",
    imageNoun: "event reference",
    analyzeLabel: "Analyze Event",
    basisSource: "I analysed the event references",
    fallbackAction: "analyze event references",
    widthLabel: "Area width (m)",
    depthLabel: "Area depth (m)",
    sizeLabel: "Area",
  },
  renovation: {
    label: "Renovation",
    assistantSubtitle: "Drop site photos or layout references, confirm the AI takeoff, then generate Excel.",
    intakeTitle: "Site reference images",
    intakeSubtitle: "Site photos, layout references, and visible measurements start this quote workflow.",
    dropTitle: "Drop site reference images to start",
    dropMeta: "JPG, PNG, or WebP. Add more site references anytime; files stay local to this runner.",
    imageNoun: "site reference",
    analyzeLabel: "Analyze Site",
    basisSource: "I analysed the site references",
    fallbackAction: "analyze site references",
    widthLabel: "Area width (m)",
    depthLabel: "Area depth (m)",
    sizeLabel: "Area",
  },
  custom: {
    label: "Custom Quote",
    assistantSubtitle: "Drop reference images, confirm the AI takeoff, then generate Excel.",
    intakeTitle: "Reference images",
    intakeSubtitle: "Images and visual references are the starting point for this quote workflow.",
    dropTitle: "Drop reference images to start",
    dropMeta: "JPG, PNG, or WebP. Add more references anytime; files stay local to this runner.",
    imageNoun: "reference image",
    analyzeLabel: "Analyze References",
    basisSource: "I analysed the reference images",
    fallbackAction: "analyze reference images",
    widthLabel: "Width (m)",
    depthLabel: "Depth (m)",
    sizeLabel: "Size",
  },
};

const state = {
  profileId: DEFAULT_PROFILE_ID,
  images: [],
  headerLogo: null,
  workflowStage: "needs_images",
  quoteBasis: { ...EMPTY_BASIS },
  lineItems: [],
  chatMessages: [],
  isAnalysisRunning: false,
  isGenerating: false,
};

const qs = (selector) => document.querySelector(selector);

const elements = {
  healthText: qs("#healthText"),
  statusDot: qs(".status-dot"),
  assistantSubtitle: qs("#assistantSubtitle"),
  generatorType: qs("#generatorType"),
  imageIntake: qs("#imageIntake"),
  intakeTitle: qs("#intakeTitle"),
  intakeSubtitle: qs("#intakeSubtitle"),
  dropTitle: qs("#dropTitle"),
  dropMeta: qs("#dropMeta"),
  imageCount: qs("#imageCount"),
  dropzone: qs(".dropzone"),
  imageInput: qs("#imageInput"),
  fileList: qs("#fileList"),
  quoteDetailsButton: qs("#quoteDetailsButton"),
  closeDetailsDrawerButton: qs("#closeDetailsDrawerButton"),
  detailsDrawer: qs("#detailsDrawer"),
  detailsBackdrop: qs("#detailsBackdrop"),
  sampleDetailsButton: qs("#sampleDetailsButton"),
  quoteCompanyName: qs("#quoteCompanyName"),
  headerDetails: qs("#headerDetails"),
  headerLogoInput: qs("#headerLogoInput"),
  clientName: qs("#clientName"),
  clientAttention: qs("#clientAttention"),
  clientTitle: qs("#clientTitle"),
  clientAddress: qs("#clientAddress"),
  projectTitle: qs("#projectTitle"),
  boothWidth: qs("#boothWidth"),
  boothDepth: qs("#boothDepth"),
  widthLabel: qs("#widthLabel"),
  depthLabel: qs("#depthLabel"),
  quoteDate: qs("#quoteDate"),
  projectNumber: qs("#projectNumber"),
  termsHeading: qs("#termsHeading"),
  paymentTerms: qs("#paymentTerms"),
  chequePayee: qs("#chequePayee"),
  notesHeading: qs("#notesHeading"),
  standardNotes: qs("#standardNotes"),
  acceptanceText: qs("#acceptanceText"),
  konceptSignatory: qs("#konceptSignatory"),
  konceptTitle: qs("#konceptTitle"),
  personLabel: qs("#personLabel"),
  stampLabel: qs("#stampLabel"),
  dateLabel: qs("#dateLabel"),
  workflowStage: qs("#workflowStage"),
  chatTranscript: qs("#chatTranscript"),
  chatActions: qs("#chatActions"),
  chatForm: qs("#chatForm"),
  chatPrompt: qs("#chatPrompt"),
  sendChatButton: qs("#sendChatButton"),
  generateButton: qs("#generateButton"),
  busyText: qs("#busyText"),
  assistantOutput: qs("#assistantOutput"),
  resultStatus: qs("#resultStatus"),
  messageList: qs("#messageList"),
  downloads: qs("#downloads"),
  excelPreview: qs("#excelPreview"),
  pricingMatchesBody: qs("#pricingMatchesBody"),
};

function currentGenerator() {
  return GENERATOR_TYPES[elements.generatorType.value] || GENERATOR_TYPES.booth;
}

function updateGeneratorCopy() {
  const generator = currentGenerator();
  elements.assistantSubtitle.textContent = generator.assistantSubtitle;
  elements.intakeTitle.textContent = generator.intakeTitle;
  elements.intakeSubtitle.textContent = generator.intakeSubtitle;
  elements.dropTitle.textContent = state.images.length ? "Add more reference images" : generator.dropTitle;
  elements.dropMeta.textContent = generator.dropMeta;
  elements.widthLabel.textContent = generator.widthLabel;
  elements.depthLabel.textContent = generator.depthLabel;
}

function setWorkflowStage(stage) {
  state.workflowStage = stage;
  elements.workflowStage.textContent = STAGE_LABELS[stage] || stage;
  elements.workflowStage.dataset.stage = stage;
}

function setDetailsDrawer(open) {
  elements.detailsDrawer.classList.toggle("is-open", open);
  elements.detailsDrawer.setAttribute("aria-hidden", open ? "false" : "true");
  elements.detailsBackdrop.hidden = !open;
  document.body.classList.toggle("drawer-open", open);
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
  renderCurrentActions();
}

function removeImageAt(index) {
  state.images.splice(index, 1);
  renderFiles();
  if (!state.images.length) {
    setWorkflowStage("needs_images");
    appendChatMessage("assistant", `No ${currentGenerator().imageNoun}s are loaded now. Drop references to start the quote analysis.`);
  }
  renderCurrentActions();
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function splitLines(value) {
  return String(value || "")
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);
}

function renderFiles() {
  elements.imageIntake.classList.toggle("has-images", Boolean(state.images.length));
  elements.imageCount.textContent = `${state.images.length} loaded`;
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

function applyQuoteDetails(details = {}) {
  const client = details.client || {};
  const project = details.project || {};
  const company = details.company || {};
  const quoteText = details.quote_text || {};
  const signature = details.signature || {};

  elements.clientName.value = client.name || "";
  elements.clientAttention.value = client.attention || "";
  elements.clientTitle.value = client.title || "";
  elements.clientAddress.value = client.address || "";
  elements.projectTitle.value = project.title || "";
  elements.boothWidth.value = project.booth_width || "";
  elements.boothDepth.value = project.booth_depth || "";
  elements.quoteDate.value = details.quote_date || "";
  elements.projectNumber.value = details.project_number || "";
  elements.quoteCompanyName.value = company.name || "";
  elements.headerDetails.value = company.header_details || "";
  elements.termsHeading.value = quoteText.terms_heading || "";
  elements.paymentTerms.value = linesValue(quoteText.payment_terms);
  elements.chequePayee.value = quoteText.cheque_payee || "";
  elements.notesHeading.value = quoteText.notes_heading || "";
  elements.standardNotes.value = linesValue(quoteText.standard_notes);
  elements.acceptanceText.value = quoteText.acceptance_text || "";
  elements.konceptSignatory.value = signature.koncept_signatory || "";
  elements.konceptTitle.value = signature.koncept_title || "";
  elements.personLabel.value = quoteText.person_label || "";
  elements.stampLabel.value = quoteText.stamp_label || "";
  elements.dateLabel.value = quoteText.date_label || "";
}

async function setSampleDetails() {
  if (state.isAnalysisRunning || state.isGenerating) return;
  elements.sampleDetailsButton.disabled = true;
  try {
    const { ok, data } = await getJson(`/api/samples/${DEFAULT_SAMPLE_ID}`);
    if (!ok) {
      appendChatMessage("assistant", (data.errors || [data.error || "Sample fixture could not be loaded."]).join("\n"), { tone: "error" });
      return;
    }
    state.profileId = data.profile_id || DEFAULT_PROFILE_ID;
    elements.generatorType.value = data.generator_type || "booth";
    updateGeneratorCopy();
    applyQuoteDetails(data.details || {});
    state.images = Array.isArray(data.images) ? data.images : [];
    state.headerLogo = null;
    renderFiles();
    setWorkflowStage(state.images.length ? "ready_to_analyze" : "needs_images");
    appendChatMessage("assistant", `${data.label || "Sample"} loaded with ${state.images.length} reference image${state.images.length === 1 ? "" : "s"}.`);
    if (state.workflowStage === "details_review") {
      showDetailReview();
    } else {
      renderCurrentActions();
    }
  } finally {
    elements.sampleDetailsButton.disabled = false;
  }
}

function buildPayload() {
  const generator = currentGenerator();
  return {
    profile_id: state.profileId || DEFAULT_PROFILE_ID,
    generator_type: elements.generatorType.value,
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
    project: {
      title: elements.projectTitle.value.trim(),
      booth_width: elements.boothWidth.value,
      booth_depth: elements.boothDepth.value,
    },
    company: {
      name: elements.quoteCompanyName.value.trim(),
      header_details: elements.headerDetails.value,
      logo_data_url: state.headerLogo ? state.headerLogo.data_url : "",
    },
    quote_basis: { ...state.quoteBasis },
    line_items: state.lineItems,
    quote_text: {
      terms_heading: elements.termsHeading.value.trim(),
      payment_terms: splitLines(elements.paymentTerms.value),
      cheque_payee: elements.chequePayee.value.trim(),
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
    },
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

function renderDownloads(files = []) {
  if (!files.length) {
    elements.downloads.innerHTML = "";
    return;
  }
  elements.downloads.innerHTML = files
    .map((file) => `
      <a class="download-link" href="${escapeHtml(file.url)}" download>
        ${escapeHtml(file.name)} - ${formatBytes(Number(file.bytes))}
      </a>
    `)
    .join("");
}

function renderExcelPreview(result = {}) {
  const rows = result.pricing_matches || [];
  if (!rows.length) {
    elements.excelPreview.innerHTML = "";
    return;
  }
  const total = rows.reduce((sum, row) => {
    const amount = Number(String(row.amount || "").replaceAll(",", ""));
    return Number.isFinite(amount) ? sum + amount : sum;
  }, 0);
  elements.excelPreview.innerHTML = `
    <div class="preview-heading">
      <strong>Excel preview</strong>
      <span>${rows.length} priced line${rows.length === 1 ? "" : "s"} - SGD ${total.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</span>
    </div>
    <div class="table-wrap preview-table-wrap">
      <table class="preview-table">
        <thead>
          <tr>
            <th>Section</th>
            <th>Quantity</th>
            <th>Description</th>
            <th>Amount</th>
          </tr>
        </thead>
        <tbody>
          ${rows.map((row) => `
            <tr>
              <td>${escapeHtml(row.section)}</td>
              <td>${escapeHtml(row.quantity)}</td>
              <td>${escapeHtml(row.description)}</td>
              <td>${escapeHtml(row.amount)}</td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    </div>
  `;
}

function renderPricingMatches(rows = []) {
  if (!rows.length) {
    elements.pricingMatchesBody.innerHTML = `<tr><td colspan="7">No pricing matches yet.</td></tr>`;
    return;
  }
  elements.pricingMatchesBody.innerHTML = rows
    .map((row) => `
      <tr>
        <td>${escapeHtml(row.status)}</td>
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

function extractPricingIssues(errors = []) {
  return errors
    .map((error) => String(error || "").trim())
    .filter((error) => /pricing:/i.test(error))
    .map((error) => error.replace(/^-?\s*/, ""));
}

function pricingIssueDescription(issue = "") {
  return String(issue)
    .replace(/^Unmatched pricing:\s*/i, "")
    .replace(/^Ambiguous pricing:\s*/i, "")
    .split(" / keyword ")[0]
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
  if (action === "mark_included") {
    item.display_price = "Included";
    appendChatMessage("assistant", `Marked "${item.description}" as included and regenerated the quotation.`);
    await handleGenerate();
    return;
  }

  if (action === "remove_line") {
    const [removed] = state.lineItems.splice(index, 1);
    appendChatMessage("assistant", `Removed "${removed.description}" from the quotation and regenerated.`);
    await handleGenerate();
    return;
  }

  if (action === "manual_price") {
    const manualPrice = window.prompt("Manual display price", item.display_price || "");
    if (!manualPrice || !manualPrice.trim()) return;
    item.display_price = manualPrice.trim();
    appendChatMessage("assistant", `Set a manual display price for "${item.description}" and regenerated.`);
    await handleGenerate();
    return;
  }

  if (action === "nearest_keyword") {
    const keyword = window.prompt("Pricing keyword to try", item.pricing_keyword || item.description);
    if (!keyword || !keyword.trim()) return;
    item.pricing_keyword = keyword.trim();
    item.display_price = "";
    appendChatMessage("assistant", `Updated the pricing keyword for "${item.description}" and regenerated.`);
    await handleGenerate();
  }
}

function renderPricingReviewMessages(result = {}) {
  const issues = extractPricingIssues(result.errors || []);
  if (!issues.length) {
    renderMessages(result.errors || ["Pricing needs review."], "warn");
    return;
  }

  elements.messageList.innerHTML = `
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
            <button class="secondary-button" type="button" data-pricing-action="mark_included" data-pricing-issue="${escapeHtml(issue)}">Mark included</button>
            <button class="secondary-button" type="button" data-pricing-action="manual_price" data-pricing-issue="${escapeHtml(issue)}">Manual display price</button>
            <button class="secondary-button" type="button" data-pricing-action="remove_line" data-pricing-issue="${escapeHtml(issue)}">Remove from quote</button>
          </div>
        </div>
      `).join("")}
    </div>
  `;
}

function appendChatMessage(role, content, options = {}) {
  state.chatMessages.push({
    role,
    content,
    html: Boolean(options.html),
    tone: options.tone || "",
  });
  logClientEvent("chat_message", {
    role,
    tone: options.tone || "",
    stage: state.workflowStage,
    content: loggableContent(content, Boolean(options.html)),
  });
  renderChat();
}

function renderChat() {
  elements.chatTranscript.innerHTML = state.chatMessages
    .map((message) => {
      const label = message.role === "user" ? "You" : "Assistant";
      const body = message.html ? message.content : `<p>${escapeHtml(message.content)}</p>`;
      return `
        <div class="chat-message ${escapeHtml(message.role)} ${escapeHtml(message.tone)}">
          <div class="chat-meta">${label}</div>
          <div class="chat-body">${body}</div>
        </div>
      `;
    })
    .join("");
  elements.chatTranscript.scrollTop = elements.chatTranscript.scrollHeight;
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
  elements.chatActions.innerHTML = actions.map(actionButton).join("");
}

function renderCurrentActions() {
  const busy = state.isAnalysisRunning || state.isGenerating;
  if (state.workflowStage === "basis_review") {
    renderChatActions([
      { label: "Confirm Basis", action: "confirm_basis", primary: true, disabled: busy },
      { label: "Regenerate Analysis", action: "regenerate", disabled: busy },
    ]);
    return;
  }
  if (state.workflowStage === "details_review") {
    renderChatActions([
      { label: "Confirm Details", action: "confirm_details", primary: true, disabled: busy },
      { label: "Regenerate Analysis", action: "regenerate", disabled: busy },
    ]);
    return;
  }
  if (state.workflowStage === "pricing_review") {
    renderChatActions([
      { label: "Regenerate Analysis", action: "regenerate", primary: true, disabled: busy },
      { label: "Generate Again", action: "generate", disabled: busy || !state.lineItems.length },
    ]);
    return;
  }
  if (state.workflowStage === "completed") {
    renderChatActions([
      { label: "Regenerate Analysis", action: "regenerate", disabled: busy },
      { label: "Generate Again", action: "generate", disabled: busy || !state.lineItems.length },
    ]);
    return;
  }
  renderChatActions([
    { label: currentGenerator().analyzeLabel, action: "analyze", primary: true, disabled: busy || !state.images.length },
  ]);
}

function basisLines(value) {
  const lines = splitLines(value);
  return lines.length ? lines : ["Confirm: No detail generated yet."];
}

function renderQuoteBasisMessage(basis = state.quoteBasis, source = "") {
  const sourceText = source === "openai"
    ? `${currentGenerator().basisSource} with OpenAI and drafted the quotation basis below.`
    : source === "gemini"
      ? `OpenAI was unavailable, so I used Gemini fallback to draft the quotation basis below from the ${currentGenerator().label.toLowerCase()} references.`
      : "I used a local starter draft for now. Review it carefully, or regenerate analysis later when a remote AI provider is available.";
  return `
    <div class="assistant-card">
      <h3>Quote basis to confirm</h3>
      <p>${escapeHtml(sourceText)} I also prepared ${state.lineItems.length} internal quotation line${state.lineItems.length === 1 ? "" : "s"} for Excel generation.</p>
      <div class="basis-review-grid">
        ${BASIS_FIELDS.map(([key, label]) => `
          <div class="basis-review-item">
            <strong>${escapeHtml(label)}</strong>
            <ul>
              ${basisLines(basis[key]).map((line) => `<li>${escapeHtml(line)}</li>`).join("")}
            </ul>
          </div>
        `).join("")}
      </div>
    </div>
  `;
}

function detailValue(value, fallback = "Not filled") {
  const clean = String(value || "").trim();
  return clean || fallback;
}

function renderDetailSummary() {
  return `
    <div class="assistant-card">
      <h3>Confirm quote details</h3>
      <p>Edit exact wording in Quote Details, then confirm here and I will generate the Excel quotation.</p>
      <dl class="detail-summary">
        <div><dt>Client</dt><dd>${escapeHtml(detailValue(elements.clientName.value))}</dd></div>
        <div><dt>Attention</dt><dd>${escapeHtml(detailValue(elements.clientAttention.value))}</dd></div>
        <div><dt>Project</dt><dd>${escapeHtml(detailValue(elements.projectTitle.value))}</dd></div>
        <div><dt>${escapeHtml(currentGenerator().sizeLabel)}</dt><dd>${escapeHtml(detailValue(elements.boothWidth.value))}m x ${escapeHtml(detailValue(elements.boothDepth.value))}m</dd></div>
        <div><dt>Quote date</dt><dd>${escapeHtml(detailValue(elements.quoteDate.value))}</dd></div>
        <div><dt>Company</dt><dd>${escapeHtml(detailValue(elements.quoteCompanyName.value))}</dd></div>
      </dl>
    </div>
  `;
}

function showDetailReview() {
  setWorkflowStage("details_review");
  appendChatMessage("assistant", renderDetailSummary(), { html: true });
  renderCurrentActions();
  syncControlStates();
}

function missingDetailFields() {
  const missing = [];
  if (!elements.clientName.value.trim()) missing.push("Client name");
  if (!elements.clientAttention.value.trim()) missing.push("Attention person");
  if (!elements.projectTitle.value.trim()) missing.push("Project / event");
  if (!elements.boothWidth.value.trim()) missing.push(currentGenerator().widthLabel);
  if (!elements.boothDepth.value.trim()) missing.push(currentGenerator().depthLabel);
  if (!elements.quoteDate.value.trim()) missing.push("Quote date");
  if (!elements.quoteCompanyName.value.trim()) missing.push("Quotation company name");
  return missing;
}

function applyDraftBasis(basis = {}) {
  state.quoteBasis = cloneQuoteBasis(basis);
}

function applyDraftLineItems(lineItems = []) {
  state.lineItems = lineItems.map(normalizeLineItem);
}

async function postJson(url, payload) {
  let response;
  try {
    response = await fetch(url, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
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

function loggableContent(content, html = false) {
  if (!html) return String(content || "");
  const container = document.createElement("div");
  container.innerHTML = String(content || "");
  return container.textContent.replace(/\s+/g, " ").trim();
}

function logClientEvent(event, details = {}) {
  fetch("/api/log", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ event, details }),
  }).catch(() => {});
}

function setAnalysisButtons(disabled) {
  elements.sendChatButton.disabled = disabled;
}

function syncControlStates() {
  const busy = state.isAnalysisRunning || state.isGenerating;
  setAnalysisButtons(busy);
  elements.generateButton.disabled = busy || !state.lineItems.length;
  renderCurrentActions();
}

async function handleDraftBasis() {
  if (state.isAnalysisRunning) return;

  if (!state.images.length) {
    setWorkflowStage("needs_images");
    appendChatMessage("assistant", `Please drop at least one ${currentGenerator().imageNoun} before analysis.`, { tone: "warn" });
    renderCurrentActions();
    syncControlStates();
    return;
  }

  state.isAnalysisRunning = true;
  setWorkflowStage("analyzing");
  elements.busyText.textContent = "Running analysis...";
  setAnalysisButtons(true);
  renderChatActions([{ label: "Analyzing...", action: "noop", disabled: true }]);
  appendChatMessage("assistant", `Analyzing the ${currentGenerator().label.toLowerCase()} references now. I will list the basis for confirmation before generating anything.`);

  const started = await startJob("draft", buildPayload());
  if (!started.ok) {
    state.isAnalysisRunning = false;
    elements.busyText.textContent = "";
    setAnalysisButtons(false);
    setWorkflowStage("ready_to_analyze");
    appendChatMessage("assistant", (started.data.errors || ["Draft failed."]).join("\n"), { tone: "error" });
    syncControlStates();
    return;
  }

  const polled = await pollJob(started.data.job_id, (job) => {
    elements.busyText.textContent = job.status === "running" ? "Running analysis..." : "Queued...";
  });
  state.isAnalysisRunning = false;
  elements.busyText.textContent = "";
  setAnalysisButtons(false);

  if (!polled.ok || ["blocked", "failed"].includes(polled.data.status)) {
    setWorkflowStage("ready_to_analyze");
    appendChatMessage("assistant", (polled.data.errors || polled.data.result?.errors || ["Draft failed."]).join("\n"), { tone: "error" });
    syncControlStates();
    return;
  }

  const data = polled.data.result || {};
  applyDraftBasis(data.quote_basis || {});
  applyDraftLineItems(data.line_items || []);
  setWorkflowStage("basis_review");
  appendChatMessage("assistant", renderQuoteBasisMessage(state.quoteBasis, data.source), { html: true, tone: data.source === "openai" ? "" : "warn" });
  if (Array.isArray(data.warnings) && data.warnings.length) {
    appendChatMessage("assistant", data.warnings.join("\n"), { tone: "warn" });
  }
  if (!state.lineItems.length) {
    appendChatMessage("assistant", EMPTY_LINE_ITEMS_MESSAGE, { tone: "warn" });
  }
  syncControlStates();
}

function confirmBasis() {
  if (state.isAnalysisRunning || state.isGenerating) return;
  if (!state.lineItems.length) {
    setWorkflowStage("basis_review");
    appendChatMessage("assistant", "I do not have any generated quotation line items yet. Regenerate analysis before confirming the basis.", { tone: "warn" });
    renderCurrentActions();
    return;
  }
  appendChatMessage("assistant", "Basis confirmed. Next I need the customer, project, company, and fixed quote text checked.");
  showDetailReview();
}

async function confirmDetails() {
  if (state.isAnalysisRunning || state.isGenerating) return;
  const missing = missingDetailFields();
  if (missing.length) {
    setWorkflowStage("details_review");
    appendChatMessage("assistant", `Please fill these details before I generate Excel: ${missing.join(", ")}.`, { tone: "warn" });
    renderCurrentActions();
    return;
  }
  appendChatMessage("assistant", "Details confirmed. Generating the Excel quotation now.");
  await handleGenerate();
}

async function handleGenerate() {
  if (state.isGenerating) return;
  if (!state.lineItems.length) {
    setWorkflowStage("basis_review");
    appendChatMessage("assistant", "There are no generated line items yet. Run analysis first so Excel has quotation rows.", { tone: "warn" });
    renderCurrentActions();
    return;
  }

  state.isGenerating = true;
  setWorkflowStage("generating");
  elements.busyText.textContent = "Generating quotation...";
  setResultStatus("Running", "is-warn");
  renderMessages([]);
  renderDownloads([]);
  renderExcelPreview({});
  renderPricingMatches([]);
  syncControlStates();
  const started = await startJob("generate", buildPayload());
  if (!started.ok) {
    state.isGenerating = false;
    elements.busyText.textContent = "";
    setWorkflowStage("details_review");
    setResultStatus(started.data.status || "Failed", "is-bad");
    renderMessages(started.data.errors || ["Generation failed."], "error");
    appendChatMessage("assistant", (started.data.errors || ["Generation failed."]).join("\n"), { tone: "error" });
    syncControlStates();
    return;
  }

  const polled = await pollJob(started.data.job_id, (job) => {
    elements.busyText.textContent = job.status === "running" ? "Generating quotation..." : "Queued...";
  });
  state.isGenerating = false;
  elements.busyText.textContent = "";

  const data = polled.data.result || polled.data || {};
  if (!polled.ok || ["blocked", "failed"].includes(polled.data.status) || data.status === "blocked" || data.status === "failed") {
    setWorkflowStage("details_review");
    setResultStatus(data.status || "Failed", "is-bad");
    renderMessages(data.errors || ["Generation failed."], "error");
    renderPricingMatches(data.pricing_matches || []);
    renderExcelPreview(data);
    appendChatMessage("assistant", (data.errors || ["Generation failed."]).join("\n"), { tone: "error" });
    syncControlStates();
    return;
  }

  if (polled.data.status === "needs_review" || data.status === "needs_confirmation") {
    setWorkflowStage("pricing_review");
    setResultStatus("Needs review", "is-warn");
    renderPricingReviewMessages(data);
    appendChatMessage("assistant", "I found pricing items that need review. I have shown the pricing review below instead of dumping the raw generator log.", { tone: "warn" });
  } else {
    setWorkflowStage("completed");
    setResultStatus("Completed", "is-ok");
    renderMessages(["Quotation package generated."]);
    appendChatMessage("assistant", "Quotation package generated. The Excel download is ready below.");
  }
  renderDownloads(data.files || []);
  renderPricingMatches(data.pricing_matches || []);
  renderExcelPreview(data);
  syncControlStates();
}

async function checkHealth() {
  try {
    const response = await fetch("/api/health");
    const data = await response.json();
    elements.statusDot.classList.toggle("is-ok", response.ok);
    elements.healthText.textContent = response.ok ? "Local server ready" : data.error || "Server unavailable";
  } catch {
    elements.healthText.textContent = "Server unavailable";
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
  if (action === "confirm_details") {
    confirmDetails();
    return;
  }
  if (action === "sample_details") {
    setSampleDetails();
    return;
  }
  if (action === "generate") {
    handleGenerate();
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

function handleChatSubmit(event) {
  event.preventDefault();
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
      confirmDetails();
      return;
    }
  }

  appendChatMessage(
    "assistant",
    `I am keeping this guarded for now: I can ${currentGenerator().fallbackAction}, regenerate the basis, load sample details, confirm basis, confirm details, or generate Excel. Edit exact customer and company text in Quote Details.`
  );
  renderCurrentActions();
}

function wireEvents() {
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

  elements.quoteDetailsButton.addEventListener("click", () => setDetailsDrawer(true));
  elements.closeDetailsDrawerButton.addEventListener("click", () => setDetailsDrawer(false));
  elements.detailsBackdrop.addEventListener("click", () => setDetailsDrawer(false));
  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape") setDetailsDrawer(false);
  });

  elements.sampleDetailsButton.addEventListener("click", setSampleDetails);
  elements.generatorType.addEventListener("change", () => {
    updateGeneratorCopy();
    renderCurrentActions();
  });
  elements.generateButton.addEventListener("click", handleGenerate);
  elements.chatForm.addEventListener("submit", handleChatSubmit);
  elements.chatActions.addEventListener("click", (event) => {
    const button = event.target.closest("[data-chat-action]");
    if (!button || button.disabled) return;
    handleChatAction(button.dataset.chatAction);
  });
  elements.messageList.addEventListener("click", (event) => {
    const button = event.target.closest("[data-pricing-action]");
    if (!button || button.disabled) return;
    handlePricingChoice(button.dataset.pricingAction, button.dataset.pricingIssue);
  });
}

function setInitialValues() {
  updateGeneratorCopy();
  renderFiles();
  renderPricingMatches([]);
  renderExcelPreview({});
  setWorkflowStage("needs_images");
  appendChatMessage("assistant", "Drop booth render images to start. Use Quote Details for customer, company, header, and terms text, or Load Sample for a quick test.");
  renderCurrentActions();
  syncControlStates();
}

wireEvents();
setInitialValues();
checkHealth();
