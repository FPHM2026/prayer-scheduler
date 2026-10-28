/* =========================================================================
   FORM ENGINE — shared visibility/validation logic and renderResponsesHtml()
   used by intake/index.html (public form), preview/index.html's staff
   "Intake Forms" tab, and portal/index.html's minister view.

   Question structure comes from window.FPHM_INTAKE_CONFIG, which starts as
   the hardcoded default from formConfig.js but is meant to be OVERWRITTEN
   at runtime by the live, staff-editable schema stored in the
   IntakeFormSchema SharePoint list (see the "Intake Forms" section of
   CLAUDE.md and the Form Editor tab in preview/index.html). Every function
   below reads `window.FPHM_INTAKE_CONFIG.SECTIONS` fresh on each call
   rather than capturing a snapshot at load time — deliberately, so
   `applyLiveSchema()` swapping in a fetched schema after this script has
   already run takes effect immediately, without needing to reload the
   page or re-run this file.
========================================================================= */

const FPHM = (function () {
  function currentSections() {
    return window.FPHM_INTAKE_CONFIG.SECTIONS;
  }

  // Overwrites the active schema (e.g. after fetching the live version
  // from SharePoint/the Worker). Accepts either a parsed object or a JSON
  // string. Throws on anything that doesn't look like a real schema, so a
  // caller can catch it and fall back to the hardcoded default instead of
  // silently rendering a broken form.
  function applyLiveSchema(schema) {
    const parsed = typeof schema === "string" ? JSON.parse(schema) : schema;
    if (!parsed || !Array.isArray(parsed.SECTIONS)) throw new Error("Invalid schema: missing SECTIONS array");
    window.FPHM_INTAKE_CONFIG = parsed;
  }

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
    return currentSections().filter((s) => isSectionVisible(s, answers));
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

  // Renders every answered question (regardless of current visibility -
  // if it has an answer, it was relevant when it was answered) as HTML,
  // grouped by section. Used for the print/PDF view and the staff view.
  //
  // Redactions (added 2026-10-21): staff can mark words in an answer as
  // personal info. They're stored alongside the answers as
  // responses._redactions = { <questionId>: ["phrase", ...] } (a key no
  // question can have, so the public form and validation ignore it).
  // opts.redact = 'hide' (minister portal) replaces each phrase with a
  // block; 'mark' (staff view) highlights it instead; omitted = plain text.
  // Each .ans-row carries data-qid so the staff view can tell which answer
  // a text selection came from.
  //
  // opts.autoNames (added 2026-10-27): with a redact mode, also hide the
  // recipient's own name and the names typed into the spouse/partner
  // questions, everywhere they appear, without staff having to pick them.
  function renderResponsesHtml(responses, opts) {
    const mode = (opts && opts.redact) || "";
    const redactions = (responses && responses._redactions) || {};
    const autoRe = mode && opts.autoNames ? autoNameRegex(responses) : null;
    let html = "";
    for (const section of currentSections()) {
      const rows = section.questions
        .filter((q) => responses[q.id] !== undefined && responses[q.id] !== null && responses[q.id] !== "" &&
          !(Array.isArray(responses[q.id]) && responses[q.id].length === 0))
        .map(
          (q) =>
            `<div class="ans-row" data-qid="${escapeHtml(q.id)}"><div class="ans-q">${escapeHtml(q.label)}</div><div class="ans-a">${answerTextHtml(
              formatAnswer(q, responses[q.id]), mode ? redactions[q.id] : null, mode, autoRe
            )}</div></div>`
        )
        .join("");
      if (!rows) continue;
      html += `<section class="ans-section"><h3>${escapeHtml(section.title)}</h3>${rows}</section>`;
    }
    return html;
  }

  // Read-only view of a saved Session Notes record (added 2026-10-28):
  // { schema:{SECTIONS}, answers } - the note carries the exact form it was
  // written with, so it's rendered from that, never the current form. Uses
  // the same .ans-* classes as renderResponsesHtml(). Handles every notes
  // question type: text/number/textarea/select, radio (+ its "__text" box),
  // checkboxes (+ "Other"), add-rows tables and choice/checkbox grids (+ the
  // grid's "__notes" box).
  function renderNotesHtml(record) {
    const sections = (record && record.schema && record.schema.SECTIONS) || [];
    const a = (record && record.answers) || {};
    const has = (v) => v !== undefined && v !== null && String(v).trim() !== "";
    const line = (label, val) => `<div>${label ? `<strong>${escapeHtml(label)}:</strong> ` : ""}${escapeHtml(val)}</div>`;
    let html = "";
    for (const section of sections) {
      const rows = (section.questions || []).map((q) => {
        const v = a[q.id];
        let body = "";
        if (q.type === "rowtable") {
          const cols = q.tableColumns || [];
          body = (Array.isArray(v) ? v : []).map((row, ri) => {
            const parts = cols.filter((c) => has(row[c.id]) || has(row[c.id + "__text"])).map((c) => {
              let val = has(row[c.id]) ? String(row[c.id]) : "";
              if (has(row[c.id + "__text"])) val += (val ? " - " : "") + row[c.id + "__text"];
              return line(c.label, val);
            }).join("");
            return parts ? `<div class="ans-subrow">${parts}</div>` : "";
          }).join("");
        } else if (q.type === "choicegrid" || q.type === "checkboxgrid") {
          if (v && typeof v === "object" && !Array.isArray(v)) {
            body = Object.keys(v).filter((r) => (Array.isArray(v[r]) ? v[r].length : has(v[r])))
              .map((r) => line(r, Array.isArray(v[r]) ? v[r].join(", ") : v[r])).join("");
          }
          if (has(a[q.id + "__notes"])) body += line(q.notesLabel || "Notes", a[q.id + "__notes"]);
        } else {
          const txt = Array.isArray(v) ? formatAnswer(q, v) : has(v) ? String(v) : "";
          if (txt) body = escapeHtml(txt);
          if (has(a[q.id + "__text"])) body += (body ? "\n" : "") + escapeHtml(a[q.id + "__text"]);
        }
        if (!body) return "";
        return `<div class="ans-row"><div class="ans-q">${escapeHtml(q.label || "")}</div><div class="ans-a">${body}</div></div>`;
      }).join("");
      if (!rows) continue;
      html += `<section class="ans-section"><h3>${escapeHtml(section.title || "")}</h3>${rows}</section>`;
    }
    return html || `<p class="ans-empty">No answers were filled in.</p>`;
  }

  // ---- redaction ----------------------------------------------------
  // A phrase matches case-insensitively and with any run of whitespace
  // standing in for a space, so a phrase picked from the rendered page
  // (where line breaks became <br>) still matches the stored answer.
  function redactionRegex(phrases) {
    const parts = (phrases || [])
      .map((p) => String(p).trim())
      .filter(Boolean)
      .sort((a, b) => b.length - a.length)
      .map((p) => p.replace(/[.*+?^${}()|[\]\\]/g, "\\$&").replace(/\s+/g, "\\s+"));
    return parts.length ? new RegExp(parts.join("|"), "gi") : null;
  }
  const REDACTED_BLOCK = "\u2588\u2588\u2588\u2588\u2588\u2588";
  function answerTextHtml(text, phrases, mode, autoRe) {
    const re = mode ? redactionRegex(phrases) : null;
    const plain = (t) => escapeHtml(t).replace(/\n/g, "<br>");
    if (!re && !autoRe) return plain(text);
    // Collect the staff-picked and automatic matches, then merge overlaps.
    const ranges = [];
    [re, autoRe].forEach((r) => {
      if (!r) return;
      r.lastIndex = 0;
      let m;
      while ((m = r.exec(text))) {
        if (!m[0]) { r.lastIndex++; continue; }
        ranges.push([m.index, m.index + m[0].length]);
      }
    });
    ranges.sort((a, b) => a[0] - b[0]);
    const merged = [];
    ranges.forEach((r) => {
      const prev = merged[merged.length - 1];
      if (prev && r[0] <= prev[1]) prev[1] = Math.max(prev[1], r[1]);
      else merged.push(r.slice());
    });
    let out = "", last = 0;
    merged.forEach(([a, b]) => {
      out += plain(text.slice(last, a));
      out += mode === "hide"
        ? `<span class="ans-redacted" aria-label="Redacted">${REDACTED_BLOCK}</span>`
        : `<mark class="ans-redacted-mark" title="Hidden from prayer ministers">${plain(text.slice(a, b))}</mark>`;
      last = b;
    });
    return out + plain(text.slice(last));
  }
  // Names hidden automatically: the recipient's own name and the spouse/
  // partner names, as a whole and word by word (so "Gary" alone is hidden
  // too). The full name matches in any case; a single word only when it
  // starts with a capital, so a name like "Grace" or "Hope" doesn't hide
  // the ordinary word in a sentence.
  function autoNameList(responses) {
    const names = [];
    ["name"].concat(NAME_FIELDS).forEach((id) => {
      const v = responses && responses[id];
      if (typeof v === "string" && v.trim()) names.push(v.trim());
    });
    return names;
  }
  function autoNameRegex(responses) {
    // Escape regex syntax, and let each letter match either case.
    const anyCase = (t) => Array.from(t).map((c) =>
      c.toLowerCase() !== c.toUpperCase() ? `[${c.toLowerCase()}${c.toUpperCase()}]`
        : /\s/.test(c) ? "\\s+"
        : c.replace(/[.*+?^${}()|[\]\\\/]/g, "\\$&")).join("").replace(/(?:\\s\+)+/g, "\\s+");
    const parts = [];
    autoNameList(responses).forEach((n) => {
      parts.push(anyCase(n));
      n.split(/[\s,\-\/&]+/).forEach((w) => {
        w = w.replace(/^[^\p{L}]+|[^\p{L}]+$/gu, "");
        if (Array.from(w).length < 2) return;
        const first = Array.from(w)[0];
        parts.push(first.toUpperCase().replace(/[.*+?^${}()|[\]\\\/]/g, "\\$&") + anyCase(w.slice(first.length)));
      });
    });
    if (!parts.length) return null;
    parts.sort((a, b) => b.length - a.length);
    return new RegExp(`(?<![\\p{L}\\p{N}])(?:${parts.join("|")})(?![\\p{L}\\p{N}])`, "gu");
  }

  function isRedacted(text, phrases) {
    const re = redactionRegex(phrases);
    return !!re && re.test(text);
  }

  // Likely personal info, offered to staff for review - never applied
  // automatically. Skips the Name and Email answers (the portal hides the
  // email, and with opts.autoNames the names are hidden automatically - so
  // anything autoNameRegex already covers isn't suggested either).
  const NAME_FIELDS = ["m_spouseName", "cl_partnerName", "d_spouseName"];
  const RELATION_WORDS = "husband|wife|son|daughter|brother|sister|mother|mom|mum|father|dad|uncle|aunt|auntie|cousin|nephew|niece|friend|boyfriend|girlfriend|fianc[eé]e?|partner|ex|ex-husband|ex-wife|grandmother|grandfather|grandma|grandpa|stepdad|stepmom|stepfather|stepmother|pastor|boss|teacher|coach|named|called";
  const CHILD_STOPWORDS = new Set(("I My Male Female Boy Girl Son Daughter Sons Daughters Living Deceased Alive Passed Died Yes No Age Aged Years Year Months Month Old " +
    "The And Both All One Two Three Four Five She He They Her His Their We Our Is Are Was Were Not Still Has Have Had Married Single Twins Stillborn Miscarriage " +
    "January February March April May June July August September October November December").split(" "));
  function suggestRedactions(responses, opts) {
    const out = [];
    const autoRe = opts && opts.autoNames ? autoNameRegex(responses) : null;
    const autoCovers = (phrase) => {
      if (!autoRe) return false;
      autoRe.lastIndex = 0;
      const m = autoRe.exec(phrase);
      return !!m && m.index === 0 && m[0].length === phrase.length;
    };
    const redactions = responses._redactions || {};
    const seen = new Set();
    const add = (qid, phrase, reason) => {
      phrase = String(phrase).trim().replace(/[.,;:]+$/, "");
      if (phrase.length < 2 || autoCovers(phrase)) return;
      const key = qid + "\u0000" + phrase.toLowerCase();
      if (seen.has(key)) return;
      seen.add(key);
      out.push({ qid, phrase, reason });
    };
    const knownNames = [];
    NAME_FIELDS.forEach((id) => { if (typeof responses[id] === "string" && responses[id].trim()) knownNames.push(responses[id].trim()); });
    for (const section of currentSections()) {
      for (const q of section.questions) {
        if (q.id === "name" || q.id === "email") continue;
        const val = responses[q.id];
        if (val === undefined || val === null || val === "") continue;
        if (typeof val !== "string" && !Array.isArray(val)) continue;
        const text = formatAnswer(q, val);
        if (!text) continue;
        const already = redactions[q.id] || [];
        const tryAdd = (phrase, reason) => { if (!isRedacted(phrase, already)) add(q.id, phrase, reason); };
        if (NAME_FIELDS.includes(q.id)) tryAdd(text, "Name");
        const patterns = [
          [/[\w.+-]+@[\w-]+(?:\.[\w-]+)+/g, "Email address"],
          [/(?:\+?1[\s.-]?)?\(?\b\d{3}\)?[\s.-]?\d{3}[\s.-]?\d{4}\b/g, "Phone number"],
          [/\b[ABCEGHJ-NPRSTVXY]\d[ABCEGHJ-NPRSTV-Z][ -]?\d[ABCEGHJ-NPRSTV-Z]\d\b/gi, "Postal code"],
          [/\b\d{1,5}\s+(?:[A-Z][\w'-]*\s+){1,3}(?:Street|St|Avenue|Ave|Road|Rd|Drive|Dr|Crescent|Cres|Boulevard|Blvd|Court|Ct|Lane|Ln|Way|Place|Pl|Circle|Cir|Trail|Parkway|Pkwy|Line|Concession|Highway|Hwy|Terrace|Square|Gate)\b\.?/g, "Street address"],
          [/\bhttps?:\/\/\S+|\bwww\.\S+/gi, "Web address"],
          [new RegExp("\\b(?:my|our|her|his|their|a|an)\\s+(?:" + RELATION_WORDS + ")\\s+((?:[A-Z][a-z'\u2019-]+)(?:\\s+[A-Z][a-z'\u2019-]+)?)", "g"), "Possible name"]
        ];
        for (const [re, reason] of patterns) {
          let m;
          while ((m = re.exec(text))) tryAdd(m[1] || m[0], reason);
        }
        knownNames.forEach((n) => {
          if (NAME_FIELDS.includes(q.id)) return;
          const re = redactionRegex([n]);
          let m;
          while ((m = re.exec(text))) tryAdd(m[0], "Name");
        });
        if (q.id === "childrenInfo") {
          const re = /\b[A-Z][a-z'\u2019-]{1,}\b/g;
          let m;
          while ((m = re.exec(text))) if (!CHILD_STOPWORDS.has(m[0])) tryAdd(m[0], "Possible child's name");
        }
      }
    }
    return out;
  }

  function escapeHtml(str) {
    return String(str)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;");
  }

  // ---- rich-text intro/liability body -------------------------------
  // INTRO_TEXT.body and LIABILITY_TEXT.body are staff-authored HTML
  // (added 2026-09-26, edited via a contenteditable rich-text box in the
  // Form Editor - see preview/index.html) rendered trusted/unescaped, the
  // same trust level as every other staff-authored piece of the schema
  // (labels, options, etc.) that already ends up on the page unescaped
  // elsewhere in this app. A body is only ever untrusted if it came from
  // somewhere other than a signed-in staff member's own edit, which never
  // happens here - the public form only ever reads this value, never
  // writes it.
  //
  // Two legacy shapes still need to render correctly without a one-time
  // migration step: INTRO_TEXT.body used to be an array of paragraph
  // strings (pre-2026-09-26), and LIABILITY_TEXT.body used to be a single
  // plain-text string relying on CSS white-space:pre-wrap for its line
  // breaks. isHtmlBody() tells "already rich HTML" apart from either -
  // legacy plain text/arrays never contain a real tag, so the heuristic
  // is reliable in practice. The Form Editor calls these same functions
  // when it first loads a legacy value into its editor box, which is what
  // actually performs the one-time migration (the next save writes back
  // real HTML) - the public form's own read path needs the same fallback
  // regardless, in case a schema is ever loaded that hasn't been through
  // the editor since.
  function isHtmlBody(str) {
    return typeof str === "string" && /<[a-z][\s\S]*>/i.test(str);
  }
  function introBodyHtml(body) {
    if (Array.isArray(body)) return body.map((p) => `<p>${escapeHtml(p)}</p>`).join("");
    if (typeof body === "string" && isHtmlBody(body)) return body;
    if (typeof body === "string" && body.trim()) {
      return body
        .split(/\n{2,}/)
        .map((p) => `<p>${escapeHtml(p)}</p>`)
        .join("");
    }
    return "";
  }
  function liabilityBodyHtml(body) {
    if (typeof body === "string" && isHtmlBody(body)) return body;
    if (typeof body === "string" && body.trim()) {
      return body
        .split(/\n{2,}/)
        .map((p) => `<p>${escapeHtml(p).replace(/\n/g, "<br>")}</p>`)
        .join("");
    }
    return "";
  }

  // ---- signature pad ------------------------------------------------
  // A signature is drawn with mouse/touch strokes onto `canvas` (mode
  // "draw"), OR rendered as a cursive rendition of a typed name (mode
  // "type", via drawTyped()) - either way the final pixels live on the
  // same canvas, so isEmpty()/toDataUrl() work identically regardless of
  // which mode the visitor used, and the stored SignatureDataUrl stays a
  // plain PNG data URL either way.
  const SIGNATURE_FONT_FAMILY = "'Dancing Script', cursive";
  let signatureFontLoadPromise = null;
  function ensureSignatureFontLoaded() {
    if (!signatureFontLoadPromise) {
      signatureFontLoadPromise = (document.fonts && document.fonts.load)
        ? document.fonts.load(`48px ${SIGNATURE_FONT_FAMILY}`).catch(() => {})
        : Promise.resolve();
    }
    return signatureFontLoadPromise;
  }

  function attachSignaturePad(canvas) {
    const ctx = canvas.getContext("2d");
    let drawing = false;
    let hasSignature = false;
    let last = null;
    let mode = "draw";
    let typedText = "";

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
      else if (mode === "type" && typedText) renderTyped();
    }
    resize();
    window.addEventListener("resize", resize);

    function pos(evt) {
      const rect = canvas.getBoundingClientRect();
      const point = evt.touches ? evt.touches[0] : evt;
      return { x: point.clientX - rect.left, y: point.clientY - rect.top };
    }
    function start(evt) {
      if (mode !== "draw") return;
      evt.preventDefault();
      drawing = true;
      hasSignature = true;
      last = pos(evt);
    }
    function move(evt) {
      if (mode !== "draw" || !drawing) return;
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

    function renderTyped() {
      const rect = canvas.getBoundingClientRect();
      ctx.clearRect(0, 0, rect.width, rect.height);
      const text = typedText.trim();
      hasSignature = !!text;
      if (!text) return;
      const fontSize = Math.max(20, Math.min(48, rect.height * 0.55));
      ctx.fillStyle = "#1a1a1a";
      ctx.textBaseline = "middle";
      ctx.textAlign = "center";
      ctx.font = `${fontSize}px ${SIGNATURE_FONT_FAMILY}`;
      ctx.fillText(text, rect.width / 2, rect.height / 2, rect.width - 20);
    }

    return {
      clear() {
        ctx.clearRect(0, 0, canvas.width, canvas.height);
        hasSignature = false;
        typedText = "";
      },
      isEmpty() {
        return !hasSignature;
      },
      toDataUrl() {
        return canvas.toDataURL("image/png");
      },
      setMode(m) {
        mode = m === "type" ? "type" : "draw";
        if (mode === "type") renderTyped();
        else {
          ctx.clearRect(0, 0, canvas.width, canvas.height);
          hasSignature = false;
        }
      },
      setTypedText(text) {
        typedText = text || "";
        ensureSignatureFontLoaded().then(renderTyped);
        renderTyped();
      }
    };
  }

  return {
    get SECTIONS() {
      return currentSections();
    },
    applyLiveSchema,
    evalCond,
    isQuestionVisible,
    isSectionVisible,
    getVisibleSections,
    getVisibleQuestions,
    validateSection,
    validateAll,
    formatAnswer,
    renderResponsesHtml,
    renderNotesHtml,
    suggestRedactions,
    isRedacted,
    escapeHtml,
    isHtmlBody,
    introBodyHtml,
    liabilityBodyHtml,
    attachSignaturePad
  };
})();
