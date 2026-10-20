(() => {
  "use strict";

  const state = {
    batchId: null,
    epoch: 0, // bumped by "start new report"; in-flight uploads/generations from before it are dropped
    files: [], // {file_id, filename, valid, error, warnings, unit_name, patient_count}
  };

  const el = {
    dropzone: document.getElementById("dropzone"),
    fileInput: document.getElementById("file-input"),
    btnChoose: document.getElementById("btn-choose"),
    btnAddMore: document.getElementById("btn-add-more"),

    panelUpload: document.getElementById("panel-upload"),
    panelFiles: document.getElementById("panel-files"),
    panelGenerate: document.getElementById("panel-generate"),
    panelResult: document.getElementById("panel-result"),

    tableBody: document.getElementById("file-table-body"),

    generateSummary: document.getElementById("generate-summary"),
    btnGenerate: document.getElementById("btn-generate"),
    outputFilename: document.getElementById("output-filename"),
    progress: document.getElementById("progress"),
    progressSteps: document.getElementById("progress-steps"),

    resultTitle: document.getElementById("result-title"),
    resultSub: document.getElementById("result-sub"),
    resultDetails: document.getElementById("result-details"),
    btnDownload: document.getElementById("btn-download"),

    errorBanner: document.getElementById("error-banner"),
    errorMessage: document.getElementById("error-message"),
  };

  function showError(message) {
    el.errorMessage.textContent = message;
    el.errorBanner.hidden = false;
    el.errorBanner.scrollIntoView({ behavior: "smooth", block: "center" });
  }

  function clearError() {
    el.errorBanner.hidden = true;
  }

  // Parse a fetch Response as JSON without throwing; a proxy/timeout/size error page
  // (HTML, empty body) becomes a readable message instead of a silent failure.
  async function readJson(resp) {
    try {
      return await resp.json();
    } catch (e) {
      return { error: `The server returned an unexpected response (HTTP ${resp.status}). ` +
                      (resp.status === 413 ? "The files are too large to upload. " : "") +
                      (resp.status === 502 || resp.status === 503 || resp.status === 504
                        ? "The server may be starting up or timed out - please wait a moment and try again." : "") };
    }
  }

  // ---------- Upload ----------

  async function uploadFiles(fileList) {
    if (!fileList || fileList.length === 0) return;
    clearError();

    const epoch = state.epoch;
    const formData = new FormData();
    if (state.batchId) formData.append("batch_id", state.batchId);
    for (const f of fileList) formData.append("files", f);

    let resp;
    try {
      resp = await fetch("/api/upload", { method: "POST", body: formData });
    } catch (e) {
      if (epoch === state.epoch) showError("Could not reach the server. Please check your connection and try again.");
      return;
    }

    const data = await readJson(resp);
    if (epoch !== state.epoch) return; // the user started a new report while this was uploading
    if (!resp.ok) {
      showError(data.error || "Upload failed.");
      return;
    }

    state.batchId = data.batch_id;
    // Merge: server returns the running order of file_ids; refetch full batch to stay in sync.
    await refreshBatch();

    // Anything the server refused outright (e.g. not an .xlsx) never enters the list, so
    // say so instead of letting it silently vanish.
    const rejected = (data.results || []).filter(r => !r.file_id);
    if (rejected.length > 0) {
      showError(rejected.map(r => `${r.filename}: ${r.error || "not accepted"}`).join(" "));
    }
  }

  async function refreshBatch() {
    if (!state.batchId) return;
    let resp;
    try {
      resp = await fetch(`/api/batch/${state.batchId}`);
    } catch (e) {
      showError("Could not reach the server. Please check your connection and try again.");
      return;
    }
    const data = await readJson(resp);
    if (!resp.ok || !data.files) {
      showError(data.error || "Could not load the uploaded files.");
      return;
    }
    state.files = data.files;
    render();
  }

  // ---------- File table rendering ----------

  // Short status label only; the explanation lives in a tooltip (hover / keyboard focus on
  // a desktop, tap on a touch screen) so the list stays compact with 15-17 files. Blocking
  // errors keep their red "Not recognized" / "Not compatible" label and still carry the full
  // explanation - they are just no longer printed under every row.
  function statusBadge(f) {
    let cls, label, tip;
    if (!f.valid) {
      cls = "error";
      label = f.error_kind === "incompatible" ? "Not compatible" : "Not recognized";
      tip = f.error || "";
    } else if (f.warnings && f.warnings.length > 0) {
      cls = "warn";
      label = "Review";
      tip = f.warnings.join("\n\n");
    } else {
      cls = "ok";
      label = "Ready";
      tip = "";
    }
    if (!tip) return `<span class="status-badge ${cls}">${label}</span>`;
    return `<button type="button" class="status-badge has-tip ${cls}" data-tip="${escapeHtml(tip)}"
              aria-label="${label}. Show details">${label}<span class="tip-icon" aria-hidden="true">i</span></button>`;
  }

  // ---------- Status tooltip (hover, keyboard focus, tap) ----------

  const tipEl = document.createElement("div");
  tipEl.className = "status-tooltip";
  tipEl.setAttribute("role", "tooltip");
  tipEl.id = "status-tooltip";
  tipEl.hidden = true;
  document.body.appendChild(tipEl);
  let tipOwner = null;
  let tipPinned = false; // opened by tap/click: stays until tapped elsewhere

  function showTip(target) {
    tipEl.textContent = target.dataset.tip;
    tipEl.hidden = false;
    tipOwner = target;
    target.setAttribute("aria-describedby", "status-tooltip");
    const r = target.getBoundingClientRect();
    const tw = tipEl.offsetWidth, th = tipEl.offsetHeight;
    let left = r.left + r.width / 2 - tw / 2;
    left = Math.max(8, Math.min(left, window.innerWidth - tw - 8));
    let top = r.bottom + 8;
    if (top + th > window.innerHeight - 8 && r.top - th - 8 > 8) top = r.top - th - 8;
    tipEl.style.left = `${left + window.scrollX}px`;
    tipEl.style.top = `${top + window.scrollY}px`;
  }

  function hideTip() {
    tipEl.hidden = true;
    if (tipOwner) tipOwner.removeAttribute("aria-describedby");
    tipOwner = null;
    tipPinned = false;
  }

  const hoverCapable = window.matchMedia && window.matchMedia("(hover: hover)").matches;
  if (hoverCapable) {
    el.tableBody.addEventListener("mouseover", (e) => {
      const t = e.target.closest(".has-tip");
      if (t && !tipPinned) showTip(t);
    });
    el.tableBody.addEventListener("mouseout", (e) => {
      const t = e.target.closest(".has-tip");
      if (t && !tipPinned && !t.contains(e.relatedTarget)) hideTip();
    });
  }
  el.tableBody.addEventListener("focusin", (e) => {
    const t = e.target.closest(".has-tip");
    if (t && !tipPinned) showTip(t);
  });
  el.tableBody.addEventListener("focusout", () => { if (!tipPinned) hideTip(); });
  document.addEventListener("click", (e) => {
    const t = e.target.closest && e.target.closest(".has-tip");
    if (t) {
      // tap/click toggles a pinned tooltip (this is how touch devices open it)
      if (tipOwner === t && tipPinned) { hideTip(); return; }
      showTip(t);
      tipPinned = true;
      return;
    }
    if (!tipEl.hidden) hideTip();
  });
  document.addEventListener("keydown", (e) => { if (e.key === "Escape") hideTip(); });
  window.addEventListener("scroll", () => { if (!tipEl.hidden && tipOwner) showTip(tipOwner); }, { passive: true });
  window.addEventListener("resize", () => { if (!tipEl.hidden && tipOwner) showTip(tipOwner); });

  function render() {
    const hasFiles = state.files.length > 0;
    el.panelUpload.hidden = hasFiles;
    el.panelFiles.hidden = !hasFiles;
    el.panelGenerate.hidden = !hasFiles;

    hideTip();
    el.tableBody.innerHTML = state.files.map((f, i) => `
      <tr data-file-id="${f.file_id}">
        <td class="col-order"><span class="order-badge">${i + 1}</span></td>
        <td class="col-name"><span class="file-name">${escapeHtml(f.filename)}</span></td>
        <td class="col-unit">
          ${f.unit_name
            ? `<span class="file-unit">${escapeHtml(f.unit_name)}</span>`
            : `<span class="file-unit unknown">not detected</span>`}
        </td>
        <td class="col-count">
          <span class="patient-count">${f.patient_count ? f.patient_count : "&mdash;"}</span>
        </td>
        <td class="col-status">${statusBadge(f)}</td>
        <td class="col-actions">
          <div class="row-actions">
            <button class="btn-icon" data-action="up" ${i === 0 ? "disabled" : ""} title="Move up" aria-label="Move up">&uarr;</button>
            <button class="btn-icon" data-action="down" ${i === state.files.length - 1 ? "disabled" : ""} title="Move down" aria-label="Move down">&darr;</button>
            <button class="btn-icon" data-action="remove" title="Remove" aria-label="Remove">&times;</button>
          </div>
        </td>
      </tr>
    `).join("");

    updateGenerateState();
  }

  function updateGenerateState() {
    const total = state.files.length;
    const invalid = state.files.filter(f => !f.valid);
    const totalPatients = state.files.reduce((sum, f) => sum + (f.patient_count || 0), 0);

    if (total === 0) {
      el.generateSummary.innerHTML = "";
      el.btnGenerate.disabled = true;
      return;
    }

    if (invalid.length > 0) {
      el.generateSummary.innerHTML =
        `<strong>${invalid.length}</strong> of ${total} file${total === 1 ? "" : "s"} ` +
        `${invalid.length === 1 ? "wasn't" : "weren't"} recognized. Remove or replace ${invalid.length === 1 ? "it" : "them"} before generating.`;
      el.btnGenerate.disabled = true;
    } else {
      el.generateSummary.innerHTML =
        `<strong>${total}</strong> unit${total === 1 ? "" : "s"} ready &middot; ` +
        `<strong>${totalPatients}</strong> patient column${totalPatients === 1 ? "" : "s"} total.`;
      el.btnGenerate.disabled = false;
    }
  }

  function escapeHtml(s) {
    if (s === null || s === undefined) return "";
    return String(s)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  // ---------- Row actions: reorder / remove ----------

  el.tableBody.addEventListener("click", async (e) => {
    const btn = e.target.closest("button[data-action]");
    if (!btn) return;
    const row = btn.closest("tr");
    const fileId = row.dataset.fileId;
    const idx = state.files.findIndex(f => f.file_id === fileId);
    if (idx === -1) return;

    if (btn.dataset.action === "remove") {
      const resp = await fetch(`/api/batch/${state.batchId}/remove/${fileId}`, { method: "POST" });
      const data = await resp.json();
      state.files = data.files;
      render();
      return;
    }

    if (btn.dataset.action === "up" && idx > 0) {
      swap(idx, idx - 1);
    } else if (btn.dataset.action === "down" && idx < state.files.length - 1) {
      swap(idx, idx + 1);
    }
  });

  async function swap(i, j) {
    const arr = state.files;
    [arr[i], arr[j]] = [arr[j], arr[i]];
    render();
    await fetch(`/api/batch/${state.batchId}/reorder`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ order: arr.map(f => f.file_id) }),
    });
    // Compatibility with the master (first file) and duplicate-patient handling depend on
    // the order, so re-read the statuses after a move.
    await refreshBatch();
  }

  // ---------- Upload interactions ----------

  el.btnChoose.addEventListener("click", () => el.fileInput.click());
  el.btnAddMore.addEventListener("click", () => el.fileInput.click());
  el.dropzone.addEventListener("click", () => el.fileInput.click());
  el.dropzone.addEventListener("keydown", (e) => {
    if (e.key === "Enter" || e.key === " ") { e.preventDefault(); el.fileInput.click(); }
  });

  el.fileInput.addEventListener("change", (e) => {
    uploadFiles(e.target.files);
    el.fileInput.value = "";
  });

  ["dragenter", "dragover"].forEach(evt =>
    el.dropzone.addEventListener(evt, (e) => {
      e.preventDefault();
      el.dropzone.classList.add("drag-over");
    })
  );
  ["dragleave", "drop"].forEach(evt =>
    el.dropzone.addEventListener(evt, (e) => {
      e.preventDefault();
      el.dropzone.classList.remove("drag-over");
    })
  );
  el.dropzone.addEventListener("drop", (e) => {
    const dt = e.dataTransfer;
    if (dt && dt.files) uploadFiles(dt.files);
  });

  // ---------- Generate ----------

  const STEP_IDS = ["reading", "validating", "extracting", "consolidating", "scoring", "summary", "finalizing"];

  // Light live sanitization as the user types - purely cosmetic feedback; the server
  // re-sanitizes authoritatively (see app.py _sanitize_output_filename) and this never
  // touches anything about the consolidation itself, only the typed display text.
  const INVALID_FILENAME_CHARS = /[\\/:*?"<>|]/g;
  el.outputFilename.addEventListener("input", () => {
    const cleaned = el.outputFilename.value.replace(INVALID_FILENAME_CHARS, "");
    if (cleaned !== el.outputFilename.value) {
      const pos = el.outputFilename.selectionStart;
      el.outputFilename.value = cleaned;
      el.outputFilename.setSelectionRange(pos - 1, pos - 1);
    }
  });

  function setStep(activeIndex, doneUpTo) {
    STEP_IDS.forEach((id, i) => {
      const li = el.progressSteps.querySelector(`[data-step="${id}"]`);
      li.classList.toggle("done", i < doneUpTo);
      li.classList.toggle("active", i === activeIndex);
    });
  }

  function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

  el.btnGenerate.addEventListener("click", async () => {
    clearError();
    el.btnGenerate.disabled = true;
    el.progress.hidden = false;
    el.panelResult.hidden = true;

    // Visual progress: step through the first stages while the request runs, then hold on
    // "finalizing" until the server actually responds - real work happens server-side.
    let cancelled = false;
    const stepThrough = (async () => {
      for (let i = 0; i < STEP_IDS.length - 1; i++) {
        if (cancelled) return;
        setStep(i, i);
        await sleep(280);
      }
      if (!cancelled) setStep(STEP_IDS.length - 1, STEP_IDS.length - 1);
    })();

    const epoch = state.epoch;
    let resp, data;
    try {
      resp = await fetch("/api/generate", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ batch_id: state.batchId, filename: el.outputFilename.value }),
      });
      data = await readJson(resp);
    } catch (e) {
      cancelled = true;
      if (epoch !== state.epoch) return;
      el.progress.hidden = true;
      el.btnGenerate.disabled = false;
      showError("Could not reach the server. Please check your connection and try again.");
      return;
    }

    await stepThrough;
    if (epoch !== state.epoch) return; // a new report was started while this one was generating

    if (!resp.ok) {
      el.progress.hidden = true;
      el.btnGenerate.disabled = false;
      let msg = data.error || "Something went wrong while generating the report.";
      if (data.invalid_files) msg += " (" + data.invalid_files.join(", ") + ")";
      showError(msg);
      return;
    }

    setStep(STEP_IDS.length, STEP_IDS.length);
    showResult(data);
  });

  function showResult(report) {
    el.panelResult.hidden = false;
    el.btnDownload.href = report.download_url;
    el.btnDownload.download = report.download_filename || "";

    const fileRows = report.files.map(f => `
      <tr>
        <td>${escapeHtml(f.filename)}</td>
        <td>${escapeHtml(f.unit || "\u2014")}</td>
        <td>${f.patient_count}</td>
        <td>${escapeHtml(f.columns)}</td>
      </tr>
    `).join("");

    el.resultDetails.innerHTML = `
      <dl>
        <dt>File name</dt><dd>${escapeHtml(report.download_filename || "")}</dd>
        <dt>Total patients</dt><dd>${report.total_patients}</dd>
        <dt>Score column</dt><dd>${escapeHtml(report.score_column)}</dd>
        <dt>Units included</dt><dd>${report.files.filter(f => f.patient_count > 0).length}</dd>
      </dl>
      <table class="file-table" style="margin-top:14px;">
        <thead><tr><th>File</th><th>Unit</th><th>Patients</th><th>Columns</th></tr></thead>
        <tbody>${fileRows}</tbody>
      </table>
    `;

    // Per-file diagnostic notes (unusual merges, text mismatches, multiple Unit/Ward
    // values, excluded empty columns, etc.) are intentionally not shown here - a
    // successful run should end with a clean, simple confirmation. The same checks
    // still run every time; their output goes to the terminal running `python app.py`
    // for anyone who wants to review it (see app.py's logging around consolidate()).

    el.panelResult.scrollIntoView({ behavior: "smooth", block: "start" });
  }

  // ---------- Home: title link -> back to the start (upload) screen ----------
  //
  // The title is a real link to "/", so it works from every screen (upload, review,
  // generating, result, error). If nothing has been added yet it just returns to the start
  // screen. If there ARE files or a generated report on screen, nothing is cleared until the
  // user explicitly confirms "Start new report" - an accidental click never loses their work.

  const homeDialog = document.getElementById("home-dialog");
  const homeLinks = [document.getElementById("home-link"), document.getElementById("logo-link")];

  function hasWorkOnScreen() {
    return state.files.length > 0 || !el.panelResult.hidden || !el.progress.hidden;
  }

  function resetToHome() {
    state.epoch += 1;
    state.batchId = null;
    state.files = [];
    el.outputFilename.value = "";
    el.progress.hidden = true;
    el.panelResult.hidden = true;
    el.resultDetails.innerHTML = "";
    el.btnDownload.removeAttribute("href");
    el.progressSteps.querySelectorAll("li").forEach(li => li.classList.remove("done", "active"));
    clearError();
    el.fileInput.value = "";
    render(); // shows the upload panel, hides review/generate, disables Generate
    window.scrollTo({ top: 0, behavior: "smooth" });
  }

  homeLinks.forEach(link => link.addEventListener("click", (e) => {
    // let "open in new tab" gestures behave like a normal link
    if (e.metaKey || e.ctrlKey || e.shiftKey || e.altKey || e.button !== 0) return;
    e.preventDefault();
    if (!hasWorkOnScreen()) { resetToHome(); return; }
    if (homeDialog && typeof homeDialog.showModal === "function") {
      homeDialog.showModal();
    } else if (window.confirm("Start a new report? The files you added and any generated report will be cleared.")) {
      resetToHome();
    }
  }));
  if (homeDialog) {
    document.getElementById("home-cancel").addEventListener("click", () => homeDialog.close());
    document.getElementById("home-confirm").addEventListener("click", () => { homeDialog.close(); resetToHome(); });
  }

  render();
})();
