/* =========================================================================
   FORM ENGINE — shared visibility/validation logic and renderResponsesHtml()
   used by intake/index.html (public form), preview/index.html's staff
   "Intake Forms" tab, and portal/index.html's minister view. Reads question
   structure from window.FPHM_INTAKE_CONFIG (formConfig.js), never hardcodes
   questions itself, so editing formConfig.js is enough to change the form.
========================================================================= */

const FPHM = (function () {
  const { SECTIONS } = window.FPHM_INTAKE_CONFIG;

  // ---- conditional visibility ------------------------------------------
  function evalCond(cond, answers) {
    if (!cond) return true;
    const val = answers[cond.id];
    if ("equals" in cond) return val === cond.equals;
    if ("notEquals" in cond) return val !== cond.notEquals;
    if ("in" in cond) return cond.in.includes(val);
    if ("includes" in cond) return Array.isArray(val) && val.includes(cond.includes);
    return true;
  }

  function isQuestionVisible(q, answers) {
    return evalCond(q.visibleIf, answers);
  }

  function isSectionVisible(section, answers) {
    return evalCond(section.visibleIf, answers);
  }

  function getVisibleSections(answers) {
    return SECTIONS.filter((s) => isSectionVisible(s, answers));
  }

  function getVisibleQuestions(section, answers) {
    return section.questions.filter((q) => isQuestionVisible(q, answers));
  }

  // ---- validation ---------------------------------------------------
  // Returns array of question ids in this section that are required,
  // visible, and currently unanswered/blank.
  function validateSection(section, answers) {
    const missing = [];
    for (const q of getVisibleQuestions(section, answers)) {
      if (!q.required) continue;
      const val = answers[q.id];
      const blank =
        val === undefined ||
        val === null ||
        val === "" ||
        (Array.isArray(val) && val.length === 0);
      if (blank) missing.push(q.id);
    }
    return missing;
  }

  function validateAll(answers) {
    const missing = [];
    for (const section of getVisibleSections(answers)) {
      missing.push(...validateSection(section, answers));
    }
    return missing;
  }

  // ---- formatting for display/print -------------------------------------
  function formatAnswer(q, val) {
    if (val === undefined || val === null || val === "") return "";
    if (Array.isArray(val)) {
      return val
        .map((v) => (v && v.startsWith && v.startsWith("__other__:") ? v.slice("__other__:".length) : v))
        .join(", ");
    }
    return String(val);
  }

  // Renders every answered question (regardless of current visibility —
  // if it has an answer, it was relevant when it was answered) as HTML,
  // grouped by section. Used for the print/PDF view and the staff view.
  function renderResponsesHtml(responses) {
    let html = "";
    for (const section of SECTIONS) {
      const rows = section.questions
        .filter((q) => responses[q.id] !== undefined && responses[q.id] !== null && responses[q.id] !== "" &&
          !(Array.isArray(responses[q.id]) && responses[q.id].length === 0))
        .map(
          (q) =>
            `<div class="ans-row"><div class="ans-q">${escapeHtml(q.label)}</div><div class="ans-a">${escapeHtml(
              formatAnswer(q, responses[q.id])
            ).replace(/\n/g, "<br>")}</div></div>`
        )
        .join("");
      if (!rows) continue;
      html += `<section class="ans-section"><h3>${escapeHtml(section.title)}</h3>${rows}</section>`;
    }
    return html;
  }

  function escapeHtml(str) {
    return String(str)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;");
  }

  // ---- signature pad ------------------------------------------------
  function attachSignaturePad(canvas) {
    const ctx = canvas.getContext("2d");
    let drawing = false;
    let hasSignature = false;
    let last = null;

    function resize() {
      const ratio = window.devicePixelRatio || 1;
      const rect = canvas.getBoundingClientRect();
      const img = hasSignature ? ctx.getImageData(0, 0, canvas.width, canvas.height) : null;
      canvas.width = rect.width * ratio;
      canvas.height = rect.height * ratio;
      ctx.scale(ratio, ratio);
      ctx.lineWidth = 2;
      ctx.lineCap = "round";
      ctx.strokeStyle = "#1a1a1a";
      if (img) ctx.putImageData(img, 0, 0);
    }
    resize();
    window.addEventListener("resize", resize);

    function pos(evt) {
      const rect = canvas.getBoundingClientRect();
      const point = evt.touches ? evt.touches[0] : evt;
      return { x: point.clientX - rect.left, y: point.clientY - rect.top };
    }
    function start(evt) {
      evt.preventDefault();
      drawing = true;
      hasSignature = true;
      last = pos(evt);
    }
    function move(evt) {
      if (!drawing) return;
      evt.preventDefault();
      const p = pos(evt);
      ctx.beginPath();
      ctx.moveTo(last.x, last.y);
      ctx.lineTo(p.x, p.y);
      ctx.stroke();
      last = p;
    }
    function end() {
      drawing = false;
    }
    canvas.addEventListener("mousedown", start);
    canvas.addEventListener("mousemove", move);
    window.addEventListener("mouseup", end);
    canvas.addEventListener("touchstart", start, { passive: false });
    canvas.addEventListener("touchmove", move, { passive: false });
    canvas.addEventListener("touchend", end);

    return {
      clear() {
        ctx.clearRect(0, 0, canvas.width, canvas.height);
        hasSignature = false;
      },
      isEmpty() {
        return !hasSignature;
      },
      toDataUrl() {
        return canvas.toDataURL("image/png");
      }
    };
  }

  return {
    SECTIONS,
    evalCond,
    isQuestionVisible,
    isSectionVisible,
    getVisibleSections,
    getVisibleQuestions,
    validateSection,
    validateAll,
    formatAnswer,
    renderResponsesHtml,
    escapeHtml,
    attachSignaturePad
  };
})();
