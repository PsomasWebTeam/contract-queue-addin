/* Outlook Add In by Psomas — the Contract Review Queue button (task pane).

   The pane reads the selected email through Outlook's own add-in permission, shows the
   same questions as the desktop app (they come from the shared questions.json through the
   Contract Queue Intake flow), and sends everything to that flow's web address. The flow,
   running under Matt's sign-in, adds the row to the SharePoint list and has Paul Atreides
   reply inside the email's conversation. Nothing is stored on this page: the flow address
   lives in the user's Outlook roaming settings, the history in the same place.

   Everything here mirrors the Python package (dates.py, questions.py, emailbuild.py,
   outbox.py) — keep the two in step. The request is sent as text/plain so the browser
   makes no CORS preflight; the flow answers with Access-Control-Allow-Origin: *.
*/
(function () {
  "use strict";
  const D = window.CQ_DEFAULTS;
  const VERSION = "1.6";
  const $ = (id) => document.getElementById(id);
  const state = {
    flowUrl: "", settings: null, questions: [], email: null, values: {}, touched: {},
    plan: null, busy: false, user: { name: "", email: "" }, controls: {}, recent: [],
  };

  // ------------------------------------------------------------------ dates & holidays
  const pad = (n) => String(n).padStart(2, "0");
  const ymd = (d) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
  const fromYmd = (s) => { const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s || ""); return m ? new Date(+m[1], +m[2] - 1, +m[3]) : null; };
  const WEEKDAYS = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
  const MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];

  function psomasHolidays(y) {
    const H = {};
    H["New Year's Day"] = new Date(y, 0, 1);
    let d = new Date(y, 4, 31); while (d.getDay() !== 1) d.setDate(d.getDate() - 1); H["Memorial Day"] = d;
    H["Independence Day"] = new Date(y, 6, 4);
    d = new Date(y, 8, 1); while (d.getDay() !== 1) d.setDate(d.getDate() + 1); H["Labor Day"] = d;
    d = new Date(y, 10, 1); while (d.getDay() !== 4) d.setDate(d.getDate() + 1); d.setDate(d.getDate() + 21);
    H["Thanksgiving Day"] = new Date(d); H["Day after Thanksgiving"] = new Date(y, 10, d.getDate() + 1);
    H["Christmas Day"] = new Date(y, 11, 25);
    return H;
  }
  function observe(name, h) {
    const d = new Date(h);
    if (name === "New Year's Day") { if (d.getDay() === 6) d.setDate(d.getDate() + 2); else if (d.getDay() === 0) d.setDate(d.getDate() + 1); return d; }
    if (d.getDay() === 6) d.setDate(d.getDate() - 1); else if (d.getDay() === 0) d.setDate(d.getDate() + 1);
    return d;
  }
  const holidayCache = {};
  function observedSet(y) {
    if (!holidayCache[y]) { const H = psomasHolidays(y); holidayCache[y] = new Set(Object.keys(H).map((k) => ymd(observe(k, H[k])))); }
    return holidayCache[y];
  }
  const isBusinessDay = (d) => d.getDay() !== 0 && d.getDay() !== 6 && !observedSet(d.getFullYear()).has(ymd(d));
  function addBusinessDays(start, n) { const d = new Date(start); let left = n; while (left > 0) { d.setDate(d.getDate() + 1); if (isBusinessDay(d)) left--; } return d; }
  function businessDaysFor(priority, rule) {
    rule = rule || {}; const want = (priority || "").trim().toLowerCase();
    for (const k of Object.keys(rule)) if (want && k.trim().toLowerCase() === want) { const n = parseInt(rule[k], 10); if (!isNaN(n)) return Math.max(0, n); }
    const dflt = parseInt(rule["default"], 10); return isNaN(dflt) ? 4 : Math.max(0, dflt);
  }
  function completionDate(received, priority, rule) {
    const start = received instanceof Date && !isNaN(received) ? new Date(received.getFullYear(), received.getMonth(), received.getDate()) : new Date();
    start.setHours(0, 0, 0, 0);
    return addBusinessDays(start, businessDaysFor(priority, rule));
  }
  const mmddyyyy = (d) => d ? `${pad(d.getMonth() + 1)}/${pad(d.getDate())}/${d.getFullYear()}` : "";
  const humanDate = (d) => `${WEEKDAYS[d.getDay()]}, ${MONTHS[d.getMonth()]} ${d.getDate()}, ${d.getFullYear()}`;
  const hhmm = (d) => { const h12 = d.getHours() % 12 || 12; return `${h12}:${pad(d.getMinutes())} ${d.getHours() < 12 ? "AM" : "PM"}`; };
  const humanDateTime = (d) => `${mmddyyyy(d)} ${hhmm(d)}`;
  const spDateTime = (d) => d.toISOString().replace(/\.\d{3}Z$/, "Z");          // real instant, UTC
  const spDate = (d) => spDateTime(new Date(d.getFullYear(), d.getMonth(), d.getDate()));  // LOCAL midnight in UTC
  const isoLocal = (d) => { const off = -d.getTimezoneOffset(), s = off >= 0 ? "+" : "-", a = Math.abs(off);
    return `${ymd(d)}T${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}${s}${pad(Math.floor(a / 60))}:${pad(a % 60)}`; };

  // ------------------------------------------------------------------ text helpers (mailinfo / emailbuild)
  const squash = (s) => String(s || "").split(/\s+/).filter(Boolean).join(" ");
  const key = (s) => squash(s).toLowerCase();
  const CHOICE_TYPES = ["choice", "multichoice"];
  function cleanSubject(s) {
    let t = (s || "").trim();
    for (let i = 0; i < 4; i++) { const t2 = t.replace(/^\s*(?:(?:re|fw|fwd|aw|wg|tr)\s*:\s*)+/i, "").replace(/^\s*\[(?:external|ext|test)\]\s*/i, ""); if (t2 === t) break; t = t2; }
    return t.trim();
  }
  const replySubject = (s) => "RE: " + cleanSubject(s);
  function firstName(name, email) {
    let n = (name || "").trim().replace(/^"|"$/g, "");
    if (n.includes(",")) n = n.split(",")[1].trim();
    if (n && !n.includes("@")) return n.split(/\s+/)[0].replace(/^[ .,'"]+|[ .,'"]+$/g, "");
    const local = (email || "").split("@")[0];
    if (local) { const f = local.split(/[._-]/)[0]; return f.charAt(0).toUpperCase() + f.slice(1); }
    return "";
  }
  const esc = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  function inlineToHtml(text) {
    const re = /\[([^\]]*)\]<(https?:\/\/[^>\s]+)>/g; let out = "", last = 0, m;
    const escInline = (t) => esc(t).replace(/\*\*(.+?)\*\*/gs, "<b>$1</b>").replace(/(?<!["=>])(https?:\/\/[^\s<]+)/g, (u) => `<a href="${u}" style="${D.linkStyle}">${u}</a>`);
    while ((m = re.exec(text)) !== null) { out += escInline(text.slice(last, m.index)); out += `<a href="${esc(m[2])}" style="${D.linkStyle}">${escInline(m[1])}</a>`; last = m.index + m[0].length; }
    return out + escInline(text.slice(last));
  }
  function textToHtml(text) {
    const blocks = String(text || "").replace(/\r\n?/g, "\n").trim().split(/\n\s*\n/).map((b) => b.trim()).filter(Boolean);
    return blocks.map((b) => `<p style="${D.prose}">${inlineToHtml(b).replace(/\n/g, "<br>")}</p>`).join("\n");
  }
  const htmlWithSignature = (text) => textToHtml(text) + "\n<p>&nbsp;</p>\n" + D.signaturePaul;
  const testBanner = (to, cc) => `<p style="background:#fff3cd;border:1px solid #b8860b;padding:8px;font-family:Segoe UI,Arial,sans-serif;font-size:10pt;"><b>TEST MODE</b> &mdash; this email was rerouted to you. The real recipients would have been &mdash; To: ${esc(to.join("; ") || "—")}; Cc: ${esc(cc.join("; ") || "—")}</p>\n`;
  const uniqueAddr = (list) => { const seen = new Set(), out = []; for (const a of list) { const k = (a || "").trim().toLowerCase(); if (k && !seen.has(k)) { seen.add(k); out.push(k); } } return out; };
  function recipients(email, settings) {
    const paul = (settings.paulEmail || "").toLowerCase(), sup = (settings.supervisors || []).filter(Boolean);
    if (!email || !email.senderEmail) return { to: [], cc: [] };
    const to = [email.senderEmail.toLowerCase()];
    const cc = uniqueAddr([...email.to.map((p) => p.email), ...email.cc.map((p) => p.email), ...sup]).filter((a) => !to.includes(a) && a !== paul);
    return { to, cc };
  }
  const GENERIC_ERRORS = ["An action failed", "ActionFailed", "The execution of template action"];
  const HIDE_STEPS = new Set(["Try", "Switch_kind", "Switch_email_mode", "Condition_found", "Until_found"]);
  const decodeJson = (f) => { try { return String(JSON.parse('"' + f + '"')); } catch (e) { return f; } };
  function shortError(raw) {
    // mirrors outbox.short_error: the failing step's own message, SharePoint's wording when it has one
    if (raw === null || raw === undefined) return "The flow reported a failure.";
    if (Array.isArray(raw) && raw.length && raw.every((x) => typeof x === "string")) return raw.map((x) => x.trim()).filter(Boolean).join("; ").slice(0, 400) || "The flow reported a failure.";
    const s = typeof raw === "string" ? raw.trim() : JSON.stringify(raw);
    if (!s) return "The flow reported a failure.";
    const sp = /"odata\.error"[\s\S]{0,400}?"value"\s*:\s*"((?:[^"\\]|\\.)+)"/.exec(s) || /"message"\s*:\s*\{[^{}]{0,80}?"value"\s*:\s*"((?:[^"\\]|\\.)+)"/.exec(s);
    if (sp) return decodeJson(sp[1]).trim().slice(0, 400);
    const hits = [...s.matchAll(/"message"\s*:\s*"((?:[^"\\]|\\.)+)"/g)];
    if (!hits.length) return s.slice(0, 400);
    const pick = hits.find((h) => !GENERIC_ERRORS.some((g) => decodeJson(h[1]).startsWith(g))) || hits[0];
    let text = decodeJson(pick[1]).trim();
    const names = [...s.matchAll(/"name"\s*:\s*"([A-Za-z0-9_]+)"/g)].filter((n) => n.index < pick.index);
    if (names.length) {
      const step = names[names.length - 1][1];
      if (!HIDE_STEPS.has(step) && !text.toLowerCase().includes(step.replace(/_/g, " ").toLowerCase())) text = `${step.replace(/_/g, " ")}: ${text}`;
    }
    return text.slice(0, 400);
  }

  // ------------------------------------------------------------------ questions (questions.py)
  function normalizeQ(q) {
    const out = Object.assign({}, q);
    out.label = squash(out.label) || "Untitled"; out.id = out.id || out.label.replace(/\W+/g, "").toLowerCase();
    out.type = D.typeNames[out.type] ? out.type : "text"; out.column = out.column || out.label.replace(/[^A-Za-z0-9]/g, (c) => `_x${c.charCodeAt(0).toString(16).padStart(4, "0")}_`).slice(0, 32);
    out.required = !!out.required; out.default = D.defaultKinds[out.default] ? out.default : "blank"; out.fixed = out.fixed || ""; out.help = out.help || "";
    if (CHOICE_TYPES.includes(out.type)) { const seen = new Set(); out.options = (out.options || []).map(squash).filter((o) => { const k = key(o); if (!k || seen.has(k)) return false; seen.add(k); return true; }); out.allowAdd = out.allowAdd !== false; if (out.store !== "text") delete out.store; } else { delete out.store; }
    return out;
  }
  const isEmpty = (v) => v === null || v === undefined || v === "" || (Array.isArray(v) && v.length === 0);
  function handlerName(settings, options) {
    const aliases = settings.handlerAliases || {}; const e = (state.user.email || "").toLowerCase();
    for (const k of Object.keys(aliases)) if (k.toLowerCase() === e) return aliases[k];
    const first = (state.user.name || "").split(" ")[0].toLowerCase();
    for (const o of options || []) if (first && o.toLowerCase() === first) return o;
    return "";
  }
  function defaultValue(q, ctx) {
    const t = q.type, k = q.default, now = new Date();
    if (k === "today") return t === "datetime" ? now : t === "date" ? ymd(now) : mmddyyyy(now);
    if (k === "emailReceived") { const r = (ctx.email && ctx.email.received) || now; return t === "datetime" ? r : t === "date" ? ymd(r) : humanDateTime(r); }
    if (k === "emailSubject") return ctx.email ? cleanSubject(ctx.email.subject) : "";
    if (k === "emailSender") return ctx.email ? ctx.email.senderName : "";
    if (k === "currentUser") return handlerName(state.settings, q.options);
    if (k === "byPriority") { const d = completionDate(ctx.received, ctx.priority, state.settings.completionRule); return t === "date" ? ymd(d) : t === "datetime" ? d : mmddyyyy(d); }
    if (k === "fixed") return coerce(q, q.fixed);
    return t === "multichoice" ? [] : null;
  }
  function coerce(q, raw) {
    const t = q.type;
    if (raw === null || raw === undefined) return t === "multichoice" ? [] : null;
    if (t === "multichoice") return (Array.isArray(raw) ? raw : String(raw).split(";")).map(squash).filter(Boolean);
    if (t === "number") { const f = parseFloat(String(raw).replace(/,/g, "")); return isNaN(f) ? null : f; }
    if (t === "yesno") return typeof raw === "boolean" ? raw : ["yes", "true", "1", "y"].includes(String(raw).trim().toLowerCase());
    return raw;
  }
  function toSharePoint(q, v) {
    if (isEmpty(v)) return null;
    switch (q.type) {
      case "multichoice": return q.store === "text" ? v.map(String).join("; ") : v.map(String);
      case "date": { const d = fromYmd(v); return d ? spDate(d) : null; }
      case "datetime": return v instanceof Date ? spDateTime(v) : null;
      case "number": return Number.isInteger(v) ? v : v;
      case "yesno": return !!v;
      default: return String(v);
    }
  }
  function displayValue(q, v) {
    if (isEmpty(v)) return "";
    switch (q.type) {
      case "multichoice": return v.join("; ");
      case "date": { const d = fromYmd(v); return d ? mmddyyyy(d) : ""; }
      case "datetime": return v instanceof Date ? humanDateTime(v) : "";
      case "yesno": return v ? "Yes" : "No";
      default: return String(v);
    }
  }

  // --- the list's real columns (apply_schema)
  const SP_TYPES = { Text: "text", Note: "note", Choice: "choice", MultiChoice: "multichoice", DateTime: "datetime", Number: "number", Currency: "number", Boolean: "yesno", URL: "text" };
  const SKIP = new Set(["ContentType", "Attachments", "_UIVersionString", "Edit", "LinkTitleNoMenu", "LinkTitle", "DocIcon", "ItemChildCount", "FolderChildCount", "AppAuthor", "AppEditor", "ComplianceAssetId", "Author", "Editor", "Created", "Modified", "ID", "SelectTitle", "LinkFilename", "Order", "GUID", "MetaInfo"]);
  function schemaFields(schema) {
    const raw = (schema && (Array.isArray(schema.value) ? schema.value : schema.fields)) || [];
    const out = [];
    for (const f of raw) {
      const internal = f.InternalName || ""; if (!internal || SKIP.has(internal) || internal.startsWith("_") || f.Hidden || f.ReadOnlyField) continue;
      const sp = f.TypeAsString || ""; if (["Computed", "Counter", "User", "UserMulti", "Lookup", "LookupMulti", "Attachments", "File", "Guid", "ContentTypeId", "WorkflowStatus", "ModStat", "Thumbnail", "Location", "Image", "Geolocation"].includes(sp)) continue;
      let choices = f.Choices; if (choices && !Array.isArray(choices)) choices = choices.results || []; choices = (choices || []).map(String);
      let ours = SP_TYPES[sp] || "text"; if (ours === "datetime" && (f.DateFormat === 0 || f.DisplayFormat === 0)) ours = "date";
      out.push({ title: String(f.Title || internal), internal, type: sp, ourType: ours, choices, required: !!f.Required });
    }
    return out;
  }
  function applySchema(questions, schema) {
    const fields = schemaFields(schema); if (!fields.length) return { questions, notes: ["The list's columns couldn't be read."], unmatched: [] };
    const byI = {}, byT = {}; fields.forEach((f) => { byI[f.internal.toLowerCase()] = f; byT[f.title.trim().toLowerCase()] = f; });
    const notes = [], used = new Set(), out = [];
    for (let q of questions) {
      q = Object.assign({}, q);
      const f = byI[(q.column || "").toLowerCase()] || byT[q.label.trim().toLowerCase()];
      if (!f) { notes.push(`'${q.label}' has no matching column on the list (column name ${q.column}).`); out.push(q); continue; }
      used.add(f.internal.toLowerCase());
      if (q.column !== f.internal) { notes.push(`'${q.label}': column name corrected to ${f.internal}.`); q.column = f.internal; }
      if (CHOICE_TYPES.includes(q.type) && CHOICE_TYPES.includes(f.ourType)) {
        if (q.store) { delete q.store; notes.push(`'${q.label}': the list column is a Choice column now, so the choices are stored as choices again.`); }
        if (q.type === "choice" && f.ourType === "multichoice") { notes.push(`'${q.label}' now allows several choices, like the list column.`); q.type = "multichoice"; }
        else if (q.type === "multichoice" && f.ourType === "choice") notes.push(`'${q.label}' allows several choices here, but the list column allows only one — in SharePoint, edit the column and tick 'Allow multiple selections'.`);
        if (f.choices.length) { const merged = normalizeQ(Object.assign({}, q, { options: [...f.choices, ...(q.options || [])] })).options; if (JSON.stringify(merged) !== JSON.stringify(q.options)) { notes.push(`'${q.label}': options now match the list.`); q.options = merged; } }
      } else if (CHOICE_TYPES.includes(q.type) && (f.ourType === "text" || f.ourType === "note")) {
        if (q.store !== "text") { q.store = "text"; notes.push(`'${q.label}' is a plain text column on the list, so the chosen options are stored as text ('A; B' when several). For a filterable column, change it in SharePoint to a Choice column with 'Allow multiple selections', then check the columns again.`); }
      } else if (f.ourType !== q.type && !(q.type === "date" && f.ourType === "datetime")) notes.push(`'${q.label}' is a ${D.typeNames[f.ourType] || f.type} column on the list but a ${D.typeNames[q.type]} question here.`);
      out.push(q);
    }
    return { questions: out, notes, unmatched: fields.filter((f) => !used.has(f.internal.toLowerCase())) };
  }

  // ------------------------------------------------------------------ the flow
  async function callFlow(payload, timeoutMs) {
    if (!state.flowUrl) throw new Error("The flow address isn't set — open Settings.");
    const ctrl = new AbortController(); const t = setTimeout(() => ctrl.abort(), timeoutMs || 115000);
    let res;
    try {
      res = await fetch(state.flowUrl, { method: "POST", headers: { "Content-Type": "text/plain;charset=UTF-8" }, body: JSON.stringify(payload), signal: ctrl.signal });
    } catch (e) {
      throw new Error(e.name === "AbortError" ? "The flow didn't answer in time (it may still finish — check the list)." : "Couldn't reach the flow (" + e.message + "). Is the address right and the flow turned on?");
    } finally { clearTimeout(t); }
    const text = await res.text(); let data;
    try { data = JSON.parse(text); } catch (e) { data = { ok: res.ok, error: text.slice(0, 300) || ("HTTP " + res.status) }; }
    if (!res.ok && data.ok !== true) throw new Error(shortError(data.error || ("HTTP " + res.status)));
    if (data.ok === false) throw new Error(shortError(data.error));
    return data;
  }
  const rand = (n) => Array.from(crypto.getRandomValues(new Uint8Array(n))).map((b) => (b % 16).toString(16)).join("");
  function stamp() { const d = new Date(); return `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}${pad(d.getHours())}${pad(d.getMinutes())}${pad(d.getSeconds())}`; }
  function basePayload(kind) {
    const s = state.settings; const user = (state.user.email || "user").split("@")[0].replace(/[^A-Za-z0-9.]+/g, "-");
    return { schema: "contract-queue/2", kind, id: `cq-${kind}-${stamp()}-${user}-${rand(4)}`, marker: "CQ" + rand(10), app: `Outlook Add In (add-in) ${VERSION}`,
      submittedBy: state.user.email, submittedByName: state.user.name, submittedAt: isoLocal(new Date()), testMode: !!s.testMode, siteUrl: s.siteUrl,
      listTitle: s.listTitle };
  }

  // ------------------------------------------------------------------ roaming settings
  function roam(k, v) { try { const rs = Office.context.roamingSettings; if (v === undefined) return rs.get(k); rs.set(k, v); rs.saveAsync(() => {}); } catch (e) { return undefined; } }

  // ------------------------------------------------------------------ the email
  function readItem() {
    const item = Office.context.mailbox && Office.context.mailbox.item; if (!item) return null;
    const pair = (p) => ({ name: p && p.displayName || "", email: (p && p.emailAddress || "").toLowerCase() });
    let restId = ""; try { restId = Office.context.mailbox.convertToRestId(item.itemId, Office.MailboxEnums.RestVersion.v2_0); } catch (e) { restId = ""; }
    return { subject: item.subject || "", senderName: item.from ? item.from.displayName || "" : "", senderEmail: item.from ? (item.from.emailAddress || "").toLowerCase() : "",
      received: item.dateTimeCreated ? new Date(item.dateTimeCreated) : null, to: (item.to || []).map(pair), cc: (item.cc || []).map(pair),
      itemId: item.itemId || "", restId, internetMessageId: item.internetMessageId || "", conversationId: item.conversationId || "" };
  }
  function emailSummary(e) {
    const who = e.senderName || e.senderEmail || "(unknown sender)"; const when = e.received ? humanDateTime(e.received) : "(date unknown)";
    return `${e.subject || "(no subject)"} — from ${who} — ${when}`;
  }

  // ------------------------------------------------------------------ UI plumbing
  function show(id, on) { $(id).classList.toggle("hidden", !on); }
  function setStatus(text, kind) { const el = $("status"); if (!text) { el.classList.add("hidden"); return; } el.className = "status " + (kind || ""); el.textContent = text; el.classList.remove("hidden"); }
  function page(name) { ["sec-setup", "sec-email", "sec-form", "sec-preview", "sec-result", "sec-settings"].forEach((id) => show(id, false));
    if (name === "setup") show("sec-setup", true);
    else if (name === "settings") { show("sec-settings", true); renderSettings(); }
    else if (name === "preview") { show("sec-email", true); show("sec-preview", true); }
    else if (name === "result") { show("sec-result", true); }
    else { show("sec-email", true); show("sec-form", true); } }

  function renderHeader() {
    const s = state.settings; const pill = $("testpill");
    if (s && s.testMode) { pill.textContent = `TEST MODE — emails go to ${s.testRecipient}; rows go to the live list (delete test rows by hand)`; pill.classList.remove("hidden"); } else pill.classList.add("hidden");
    $("nav-list").href = s ? `${(s.siteUrl || "").replace(/\/$/, "")}/Lists/${encodeURIComponent(s.listTitle)}/AllItems.aspx` : "#";
  }
  function renderEmail() {
    const e = state.email;
    if (!e) { $("email-summary").textContent = "Open or select an email, then press the button again."; $("email-detail").textContent = ""; show("email-pill", false); return; }
    $("email-summary").textContent = emailSummary(e);
    const who = [...e.to, ...e.cc].map((p) => p.email).filter(Boolean).join("; ") || "—";
    $("email-detail").textContent = `From ${e.senderEmail || "?"} · To/CC: ${who}`;
    $("email-pill").textContent = "selected in Outlook"; show("email-pill", true);
  }

  // --- the form
  function renderForm() {
    const host = $("form"); host.innerHTML = ""; state.controls = {};
    $("list-pill").textContent = "→ " + state.settings.listTitle;
    for (const q of state.questions) {
      const wrap = document.createElement("div"); wrap.className = "q";
      const lab = document.createElement("div"); lab.className = "qlabel"; lab.innerHTML = esc(q.label) + (q.required ? ' <span class="req">*</span>' : ""); wrap.appendChild(lab);
      const ctl = makeControl(q); wrap.appendChild(ctl.el);
      if (q.help) { const h = document.createElement("div"); h.className = "help"; h.textContent = q.help; wrap.appendChild(h); }
      state.controls[q.id] = ctl; host.appendChild(wrap);
    }
    applyDefaults();
  }
  function touched(q) { state.touched[q.id] = true; if (q.id === "priority" || q.id === "dateReceived") recalcDerived(false); }
  function makeControl(q) {
    const t = q.type;
    if (t === "choice") {
      const el = document.createElement("div"); el.className = "inline";
      const sel = document.createElement("select"); (q.options || []).forEach((o) => { const op = document.createElement("option"); op.value = o; op.textContent = o; sel.appendChild(op); });
      sel.addEventListener("change", () => touched(q)); el.appendChild(sel);
      if (q.allowAdd) el.appendChild(addOptionButton(q, (v) => { sel.value = v; touched(q); }));
      return { el, get: () => sel.value || null, set: (v) => { if (v && [...sel.options].some((o) => o.value === v)) sel.value = v; else sel.selectedIndex = 0; }, setOptions: (opts, pick) => { sel.innerHTML = ""; opts.forEach((o) => { const op = document.createElement("option"); op.value = o; op.textContent = o; sel.appendChild(op); }); if (pick) sel.value = pick; } };
    }
    if (t === "multichoice") return multiControl(q);
    if (t === "date") {
      const el = document.createElement("div"); el.className = "inline"; const inp = document.createElement("input"); inp.type = "date"; inp.addEventListener("change", () => touched(q)); inp.addEventListener("input", () => touched(q)); el.appendChild(inp);
      let note = null; if (q.default === "byPriority") { note = document.createElement("div"); note.className = "rule-note"; }
      const outer = document.createElement("div"); outer.appendChild(el); if (note) outer.appendChild(note);
      return { el: outer, get: () => inp.value || null, set: (v) => { inp.value = v || ""; }, note };
    }
    if (t === "datetime") {
      const el = document.createElement("div"); el.className = "inline";
      const d = document.createElement("input"); d.type = "date"; const tm = document.createElement("input"); tm.type = "time";
      [d, tm].forEach((i) => { i.addEventListener("change", () => touched(q)); i.addEventListener("input", () => touched(q)); });
      const now = document.createElement("button"); now.type = "button"; now.className = "mini secondary"; now.textContent = "Now"; now.addEventListener("click", () => { setDT(new Date()); touched(q); });
      const em = document.createElement("button"); em.type = "button"; em.className = "mini secondary"; em.textContent = "Email's date"; em.addEventListener("click", () => { if (state.email && state.email.received) { setDT(state.email.received); touched(q); } });
      el.append(d, tm, now, em);
      const setDT = (v) => { if (v instanceof Date && !isNaN(v)) { d.value = ymd(v); tm.value = `${pad(v.getHours())}:${pad(v.getMinutes())}`; } else { d.value = ""; tm.value = ""; } };
      return { el, get: () => { const dd = fromYmd(d.value); if (!dd) return null; const m = /^(\d{2}):(\d{2})/.exec(tm.value || ""); if (m) dd.setHours(+m[1], +m[2], 0, 0); return dd; }, set: setDT };
    }
    if (t === "note") { const ta = document.createElement("textarea"); ta.rows = 3; ta.addEventListener("input", () => touched(q)); return { el: ta, get: () => ta.value.trim() || null, set: (v) => { ta.value = v || ""; } }; }
    if (t === "yesno") { const el = document.createElement("label"); const cb = document.createElement("input"); cb.type = "checkbox"; cb.addEventListener("change", () => touched(q)); el.append(cb, " Yes"); return { el, get: () => cb.checked, set: (v) => { cb.checked = !!v; } }; }
    const inp = document.createElement("input"); inp.type = t === "number" ? "number" : "text"; inp.addEventListener("input", () => touched(q));
    return { el: inp, get: () => { const v = inp.value.trim(); return v ? coerce(q, v) : null; }, set: (v) => { inp.value = v === null || v === undefined ? "" : String(v); } };
  }
  function addOptionButton(q, onAdded) {
    const b = document.createElement("button"); b.type = "button"; b.className = "mini secondary"; b.textContent = "＋ Add option";
    b.addEventListener("click", () => {
      if (b.nextSibling && b.nextSibling.className === "addrow") { b.nextSibling.remove(); return; }
      const row = document.createElement("div"); row.className = "addrow"; const inp = document.createElement("input"); inp.type = "text"; inp.placeholder = `New option for ${q.label}`;
      const ok = document.createElement("button"); ok.type = "button"; ok.className = "mini primary"; ok.textContent = "Add";
      ok.addEventListener("click", async () => { const v = squash(inp.value); if (!v) return; ok.disabled = true; try { await addOption(q, v); onAdded(v); row.remove(); } catch (e) { setStatus(e.message, "err"); ok.disabled = false; } });
      row.append(inp, ok); b.parentNode.insertBefore(row, b.nextSibling); inp.focus();
    });
    return b;
  }
  function multiControl(q) {
    const det = document.createElement("details"); det.className = "multi";
    const sum = document.createElement("summary"); const sumText = document.createElement("span"); const chev = document.createElement("span"); chev.className = "chev"; chev.textContent = "▾"; sum.append(sumText, chev); det.appendChild(sum);
    const filt = document.createElement("input"); filt.type = "text"; filt.className = "filter"; filt.placeholder = "Type to filter…"; det.appendChild(filt);
    const opts = document.createElement("div"); opts.className = "opts"; det.appendChild(opts);
    const boxes = {};
    const render = () => { opts.innerHTML = ""; const needle = filt.value.trim().toLowerCase(); for (const o of q.options || []) { if (needle && !o.toLowerCase().includes(needle)) continue; const lab = document.createElement("label"); const cb = document.createElement("input"); cb.type = "checkbox"; cb.checked = !!boxes[o]; cb.addEventListener("change", () => { boxes[o] = cb.checked; refreshSummary(); touched(q); }); lab.append(cb, o); opts.appendChild(lab); } };
    const refreshSummary = () => { const chosen = (q.options || []).filter((o) => boxes[o]); sumText.textContent = chosen.length ? chosen.join("; ") : "None chosen"; sumText.style.color = chosen.length ? "" : "var(--muted)"; };
    filt.addEventListener("input", render);
    const outer = document.createElement("div"); outer.appendChild(det);
    if (q.allowAdd) { const row = document.createElement("div"); row.className = "row"; row.style.marginTop = "6px"; row.appendChild(addOptionButton(q, (v) => { boxes[v] = true; render(); refreshSummary(); touched(q); })); outer.appendChild(row); }
    render(); refreshSummary();
    return { el: outer, get: () => (q.options || []).filter((o) => boxes[o]), set: (vals) => { Object.keys(boxes).forEach((k) => delete boxes[k]); for (const v of vals || []) { const match = (q.options || []).find((o) => key(o) === key(v)); if (match) boxes[match] = true; } render(); refreshSummary(); }, setOptions: (newOpts, pick) => { q.options = newOpts; if (pick) boxes[pick] = true; render(); refreshSummary(); } };
  }
  function ctx() { return { email: state.email, priority: state.controls.priority ? state.controls.priority.get() : null, received: state.controls.dateReceived ? state.controls.dateReceived.get() : null }; }
  function applyDefaults() {
    state.touched = {}; const c = ctx();
    for (const q of state.questions) if (q.default !== "byPriority") state.controls[q.id].set(defaultValue(q, c));
    recalcDerived(true);
  }
  function recalcDerived(force) {
    const c = ctx(); const n = businessDaysFor(c.priority, state.settings.completionRule);
    for (const q of state.questions) { if (q.default !== "byPriority") continue; const ctl = state.controls[q.id]; if (force || !state.touched[q.id]) ctl.set(defaultValue(q, c));
      if (ctl.note) ctl.note.textContent = `${n} business day${n === 1 ? "" : "s"} after Date Received for ${c.priority || "this priority"} — change it if you need to`; }
  }
  function collectValues() {
    const values = {}; for (const q of state.questions) values[q.id] = state.controls[q.id].get();
    const missing = state.questions.filter((q) => q.required && isEmpty(values[q.id])).map((q) => q.label);
    if (missing.length) return { error: "Please fill in: " + missing.join(", ") };
    return { values };
  }

  // --- Paul's email
  function placeholders(values) {
    const s = state.settings, e = state.email, ph = {};
    for (const q of state.questions) ph[q.id] = displayValue(q, values[q.id]);
    const comp = values.estimatedCompletion ? fromYmd(values.estimatedCompletion) : null;
    ph.completionDate = comp ? humanDate(comp) : (ph.estimatedCompletion || "");
    ph.title = ph.title || ""; ph.priority = ph.priority || ""; ph.handler = ph.currentHandler || "";
    ph.dateReceived = values.dateReceived instanceof Date ? humanDateTime(values.dateReceived) : (ph.dateReceived || "");
    ph.queueLink = s.queueLink || ""; ph.listTitle = s.listTitle || ""; ph.senderName = e ? e.senderName : "";
    ph.firstName = e ? firstName(e.senderName, e.senderEmail) : ""; if (!ph.firstName) ph.firstName = "there";
    return ph;
  }
  const fill = (text, ph) => String(text || "").replace(/\{([A-Za-z][A-Za-z0-9]*)\}/g, (m, k) => (k in ph ? String(ph[k]) : m));
  function buildPlan(values) {
    const s = state.settings, ph = placeholders(values), r = recipients(state.email, s), t = s.emailTemplate || {};
    const text = [t.greeting, t.body, t.closing].map((p) => fill(p, ph).trim()).filter(Boolean).join("\n\n");
    return { from: s.paulEmail, to: r.to, cc: r.cc, subject: state.email && state.email.subject ? replySubject(state.email.subject) : `Contract received: ${ph.title || "your contract"}`, text, values };
  }
  function finalize(plan, text, to, cc, subject) {
    const s = state.settings; const realTo = uniqueAddr(to), realCc = uniqueAddr(cc).filter((a) => !realTo.includes(a));
    const body = htmlWithSignature(text);
    const out = { mode: "reply", messageId: state.email ? state.email.restId : "", internetMessageId: state.email ? state.email.internetMessageId : "", from: plan.from, to: realTo.join(";"), cc: realCc.join(";"), subject: subject.trim(), bodyHtml: body, realTo, realCc };
    if (!out.messageId) out.mode = "new";
    if (s.testMode) { out.to = (s.testRecipient || "").toLowerCase(); out.cc = ""; out.subject = "[TEST] " + out.subject; out.bodyHtml = testBanner(realTo, realCc) + body; }
    return out;
  }
  const splitAddrs = (s) => s.replace(/,/g, ";").split(";").map((a) => a.trim().toLowerCase()).filter(Boolean);
  const badAddrs = (list) => list.filter((a) => a.includes("<") || a.includes(">") || !a.includes("@") || a.includes(" "));

  function review(listOnly) {
    const got = collectValues(); if (got.error) { $("form-problem").textContent = got.error; return; }
    $("form-problem").textContent = ""; state.plan = buildPlan(got.values);
    if (listOnly) { submit(false); return; }
    $("pv-from").textContent = state.plan.from; $("pv-to").value = state.plan.to.join("; "); $("pv-cc").value = state.plan.cc.join("; ");
    $("pv-subject").value = state.plan.subject; $("pv-body").value = state.plan.text; $("preview-problem").textContent = "";
    const s = state.settings; if (s.testMode) { $("preview-pill").textContent = `TEST MODE: this will go to ${s.testRecipient} only`; show("preview-pill", true); } else show("preview-pill", false);
    page("preview");
  }

  async function submit(withEmail) {
    if (state.busy || !state.plan) return;
    const s = state.settings, values = state.plan.values; let message = null;
    if (withEmail) {
      const to = splitAddrs($("pv-to").value), cc = splitAddrs($("pv-cc").value), subject = $("pv-subject").value.trim(), text = $("pv-body").value;
      const problems = []; if (!to.length) problems.push("Add at least one address in the To line, or use 'Add to the list only'."); if (badAddrs([...to, ...cc]).length) problems.push("Use plain email addresses, separated by semicolons.");
      if (!subject) problems.push("The subject is empty."); if (!text.trim()) problems.push("The email text is empty.");
      if (problems.length) { $("preview-problem").textContent = problems.join(" "); return; }
      message = finalize(state.plan, text, to, cc, subject);
    }
    state.busy = true; ["btn-send", "btn-send-list-only", "btn-review", "btn-list-only"].forEach((id) => { $(id).disabled = true; });
    $("btn-send").textContent = "Sending…"; $("preview-problem").textContent = "";
    const fields = {}, answers = {};
    for (const q of state.questions) { const v = toSharePoint(q, values[q.id]); if (v !== null) fields[q.column] = v; answers[q.label] = displayValue(q, values[q.id]); }
    const payload = Object.assign(basePayload("create-item"), { fields, answers, emailSentVia: message ? "flow" : "none",
      email: message ? { mode: message.mode, messageId: message.messageId, internetMessageId: message.internetMessageId, from: message.from, to: message.to, cc: message.cc, subject: message.subject, bodyHtml: message.bodyHtml } : null,
      source: state.email ? { subject: state.email.subject, from: state.email.senderEmail, fromName: state.email.senderName, receivedAt: state.email.received ? isoLocal(state.email.received) : "", internetMessageId: state.email.internetMessageId, how: "addin" } : { how: "none" } });
    setStatus(message ? "Adding the row and sending Paul's reply… this can take up to a minute." : "Adding the row…", "busy");
    let res;
    try { res = await callFlow(payload); }
    catch (e) {
      setStatus(e.message, "err"); state.busy = false; ["btn-send", "btn-send-list-only", "btn-review", "btn-list-only"].forEach((id) => { $(id).disabled = false; }); $("btn-send").textContent = "Send Paul's email and add to the list"; return;
    }
    state.busy = false; setStatus("", "");
    const itemId = res.itemId ? parseInt(res.itemId, 10) : null;
    $("res-list").textContent = `✓ Added to “${payload.listTitle}”` + (itemId ? ` (item ${itemId})` : "") + ".";
    const sent = res.emailSent || (message ? "flow" : "none");
    const emailErrors = Array.isArray(res.emailErrors) ? res.emailErrors.filter(Boolean).map(String) : [];
    $("res-email").textContent = !message ? "No acknowledgment email was sent." : (emailErrors.length && sent !== "reply" && sent !== "new") ? `✗ The row is on the list, but Paul's email was NOT sent: ${emailErrors.join("; ")}`
      : sent === "reply" ? `✓ Paul replied inside the conversation to ${message.to}` + (s.testMode ? " (TEST MODE — rerouted to you)" : "") + "."
      : sent === "new" ? `✓ Paul's mailbox sent a new email to ${message.to}` + (s.testMode ? " (TEST MODE — rerouted to you)" : "") + "." : `Paul's email: ${sent}.`;
    const openBtn = $("res-open"); if (itemId) { openBtn.href = `${(s.siteUrl || "").replace(/\/$/, "")}/Lists/${encodeURIComponent(payload.listTitle)}/DispForm.aspx?ID=${itemId}`; show("res-open", true); } else show("res-open", false);
    remember({ when: humanDateTime(new Date()), title: answers.Title || "", itemId, list: payload.listTitle, email: sent, to: message ? message.to : "" });
    page("result");
    ["btn-send", "btn-send-list-only", "btn-review", "btn-list-only"].forEach((id) => { $(id).disabled = false; }); $("btn-send").textContent = "Send Paul's email and add to the list";
  }

  // --- shared edits through the flow
  async function fetchConfig() {
    const res = await callFlow(basePayloadLite("config"), 40000);
    const merged = Object.assign({}, D.settings, res.settings || {}); merged.emailTemplate = Object.assign({}, D.settings.emailTemplate, (res.settings && res.settings.emailTemplate) || {});
    const qs = (res.questions && (Array.isArray(res.questions) ? res.questions : res.questions.questions)) || D.questions;
    state.settings = merged; state.questions = qs.map(normalizeQ);
  }
  function basePayloadLite(kind) { const s = state.settings || D.settings; return { schema: "contract-queue/2", kind, id: `cq-${kind}-${stamp()}-${rand(4)}`, submittedBy: state.user.email, submittedAt: isoLocal(new Date()), listTitle: s.listTitle }; }
  async function saveQuestions() { await callFlow(Object.assign(basePayloadLite("save-file"), { file: "questions.json", content: JSON.stringify({ version: 1, questions: state.questions }, null, 2) }), 40000); }
  async function addOption(q, value) {
    await fetchConfig();                                            // re-read before editing (the family rule)
    const target = state.questions.find((x) => x.id === q.id); if (!target) throw new Error("That question no longer exists in the shared file.");
    if ((target.options || []).some((o) => key(o) === key(value))) throw new Error(`“${value}” is already an option.`);
    target.options = [...(target.options || []), value]; q.options = target.options;
    await saveQuestions();
    try { await callFlow(Object.assign(basePayloadLite("add-choice"), { column: target.column, columnLabel: target.label, choices: target.options }), 40000); }
    catch (e) { setStatus(`Saved for the form; the list's own drop-down couldn't be updated: ${e.message}`, "warn"); return; }
    const ctl = state.controls[q.id]; if (ctl && ctl.setOptions) ctl.setOptions(target.options, value);
    setStatus(`Added “${value}” to ${target.label} for everyone.`, "ok");
  }
  async function checkColumns() {
    setStatus("Asking the flow for the list's columns…", "busy"); $("set-msg").textContent = "";
    try {
      const res = await callFlow(basePayloadLite("refresh-schema"), 60000);
      let schema = res.schema; if (typeof schema === "string") schema = JSON.parse(schema);
      const r = applySchema(state.questions, schema); state.questions = r.questions.map(normalizeQ);
      await saveQuestions();
      const lines = [`Read ${schemaFields(schema).length} columns from the list.`, ...r.notes];
      if (r.unmatched.length) lines.push("Columns with no question yet: " + r.unmatched.map((f) => f.title).join(", ") + " — add them as questions in the desktop app's Settings.");
      $("set-msg").textContent = lines.join(" "); setStatus("Columns checked.", "ok"); renderForm();
    } catch (e) { setStatus(e.message, "err"); }
  }

  // --- settings & recent
  function remember(entry) { const list = [entry, ...(state.recent || [])].slice(0, 15); state.recent = list; roam("recent", JSON.stringify(list)); }
  function renderSettings() {
    $("set-url").value = state.flowUrl || "";
    $("about").textContent = `Outlook Add In (add-in) ${VERSION} · signed in as ${state.user.email}`;
    const host = $("recent"); host.innerHTML = ""; if (!state.recent.length) host.textContent = "Nothing yet.";
    for (const r of state.recent) { const row = document.createElement("div"); row.className = "recent-row"; row.textContent = `${r.when} — ${r.title || "(no title)"} → ${r.list}${r.itemId ? " · item " + r.itemId : ""}${r.email && r.email !== "none" ? " · Paul: " + r.email : ""}`; host.appendChild(row); }
  }
  async function connect(url, fromSetup) {
    const msg = fromSetup ? $("setup-msg") : $("set-msg");
    url = (url || "").trim(); if (!/^https:\/\//i.test(url)) { msg.textContent = "That doesn't look like a flow address (it starts with https://)."; return false; }
    state.flowUrl = url; msg.textContent = "Connecting…";
    try { await fetchConfig(); } catch (e) { msg.textContent = e.message; return false; }
    roam("flowUrl", url); msg.textContent = `Connected — ${state.questions.length} questions · test mode ${state.settings.testMode ? "ON" : "off"}.`;
    renderHeader(); renderForm(); return true;
  }

  // ------------------------------------------------------------------ init
  function wire() {
    $("btn-review").addEventListener("click", () => review(false));
    $("btn-list-only").addEventListener("click", () => review(true));
    $("btn-back").addEventListener("click", () => page("main"));
    $("btn-send").addEventListener("click", () => submit(true));
    $("btn-send-list-only").addEventListener("click", () => submit(false));
    $("btn-again").addEventListener("click", () => { state.plan = null; state.email = readItem(); renderEmail(); renderForm(); page("main"); });
    $("setup-save").addEventListener("click", async () => { if (await connect($("setup-url").value, true)) { page("main"); } });
    $("set-save").addEventListener("click", async () => { if (await connect($("set-url").value, false)) setStatus("Flow address saved.", "ok"); });
    $("set-test").addEventListener("click", async () => { $("set-msg").textContent = "Testing…"; try { await fetchConfig(); $("set-msg").textContent = `The flow answered — ${state.questions.length} questions, test mode ${state.settings.testMode ? "ON" : "off"}.`; renderHeader(); renderForm(); } catch (e) { $("set-msg").textContent = e.message; } });
    $("set-schema").addEventListener("click", checkColumns);
    $("nav-main").addEventListener("click", (e) => { e.preventDefault(); page(state.flowUrl ? "main" : "setup"); });
    $("nav-settings").addEventListener("click", (e) => { e.preventDefault(); page("settings"); });
  }

  // a few pure helpers exposed for checking from a browser console (no data, no secrets)
  window.__cq = { completionDate, businessDaysFor, isBusinessDay, spDate, spDateTime, humanDate, cleanSubject, replySubject,
    firstName, textToHtml, shortError, recipients, applySchema, normalizeQ, mmddyyyy, version: VERSION };

  Office.onReady(async () => {
    wire();
    try { const p = Office.context.mailbox.userProfile; state.user = { name: p.displayName || "", email: (p.emailAddress || "").toLowerCase() }; } catch (e) { state.user = { name: "", email: "" }; }
    try { state.recent = JSON.parse(roam("recent") || "[]"); } catch (e) { state.recent = []; }
    state.flowUrl = roam("flowUrl") || "";
    state.settings = Object.assign({}, D.settings); state.questions = D.questions.map(normalizeQ);
    state.email = readItem(); renderHeader(); renderEmail();
    if (!state.flowUrl) { page("setup"); return; }
    setStatus("Connecting to the flow…", "busy");
    try { await fetchConfig(); setStatus("", ""); } catch (e) { setStatus(e.message + " Using the built-in defaults until the flow answers.", "warn"); }
    renderHeader(); renderForm(); page("main");
  });
})();
