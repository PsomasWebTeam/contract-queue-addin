/* Outlook Add In by Psomas — the on-send watcher and the "Mark returned" command (v2.2).

   Classic Outlook runs the on-send watcher from this file alone in a JavaScript-only runtime (no page, no
   DOM, no window globals such as atob), so it is self-contained: nothing from taskpane.js is imported. The
   "Mark returned" button and Outlook on the web / the new Outlook load it through commands.html (a browser
   runtime) — Office.onReady below tells that runtime the page is ready; without it a function command never
   runs and classic Outlook shows "… is working on your request" until it gives up.

   When the reviewer presses Send on a reply that (1) goes to someone outside Legal and (2) carries an
   attachment (signature pictures do not count) and (3) belongs to an email thread with an open entry in the
   Contract Review Queue, a box appears: "Send anyway" sends without touching the queue; "Mark returned & send"
   stamps Date Returned in the Contracts Review Log, removes the row from the queue, and asks for one more
   press of Send. Replies that stay inside Legal (a first pass handed to Chad) never see the box.

   Every step is bounded in time and every command path ends in event.completed(); a synchronous error
   completes the command too and says what happened in the notification area.
*/
(function () {
  "use strict";
  try { Office.onReady(function () { /* the browser runtime (commands.html) waits for this */ }); } catch (e) { /* JS-only runtime */ }

  var LEGAL = ["matthew.pipes@psomas.com", "chad.wilson@psomas.com", "paul.atreides@psomas.com", "legal@psomas.com"];
  var SKIP = "cqSkipMark";                 // set on the email once it has been marked (or the user chose not to)
  var DIAG = true;                         // temporary: a tiny 'thread' probe to the flow at load and at command entry (visible in the run history)

  function roam(key) { try { return Office.context.roamingSettings.get(key); } catch (e) { return undefined; } }
  function settings() { try { var c = roam("configCache"); return c ? (JSON.parse(c).settings || {}) : {}; } catch (e) { return {}; } }
  function legalAddresses(s) {
    var list = LEGAL.concat(s.supervisors || [], [s.paulEmail || ""], s.internalAddresses || []);
    return list.map(function (a) { return String(a || "").trim().toLowerCase(); }).filter(Boolean);
  }
  function pad2(n) { n = String(n); return n.length < 2 ? "0" + n : n; }
  function isoLocal(d) {
    var off = -d.getTimezoneOffset(), sign = off >= 0 ? "+" : "-", a = Math.abs(off), p = pad2;
    return d.getFullYear() + "-" + p(d.getMonth() + 1) + "-" + p(d.getDate()) + "T" + p(d.getHours()) + ":" + p(d.getMinutes()) + ":" + p(d.getSeconds()) + sign + p(Math.floor(a / 60)) + ":" + p(a % 60);
  }
  function stamp() { var d = new Date(), p = pad2; return "" + d.getFullYear() + p(d.getMonth() + 1) + p(d.getDate()) + p(d.getHours()) + p(d.getMinutes()) + p(d.getSeconds()); }
  function rand(n) { var s = ""; while (s.length < n) s += Math.random().toString(36).slice(2); return s.slice(0, n); }
  function errText(e) { return e && e.message ? e.message : String(e); }

  /* Timers that may not exist in a stripped runtime: null when unavailable. */
  function timer(fn, ms) { try { return setTimeout(fn, ms); } catch (e) { return null; } }
  function cancel(t) { if (t !== null && t !== undefined) { try { clearTimeout(t); } catch (e) { /* nothing */ } } }
  /* Every Office callback is bounded: if it has not answered after `ms`, the promise settles with `fallback`.
     (Outlook shows "… is working on your request" until event.completed() — a callback that never fires
     must not hold that bar forever.) */
  function bounded(promise, ms, fallback) {
    return new Promise(function (res) {
      var t = timer(function () { res(fallback); }, ms);
      promise.then(function (v) { cancel(t); res(v); }, function () { cancel(t); res(fallback); });
    });
  }

  var B64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
  function b64bytes(str) {        // base64 → array of byte values (no atob: the JS-only runtime has no DOM globals)
    var out = [], buf = 0, bits = 0;
    for (var i = 0; i < str.length; i++) {
      var c = str.charAt(i); if (c === "=") break;
      var v = B64.indexOf(c); if (v < 0) continue;
      buf = ((buf << 6) | v) & 0xFFFFFF; bits += 6;
      if (bits >= 8) { bits -= 8; out.push((buf >> bits) & 255); }
    }
    return out;
  }
  /* The 16-byte conversation GUID as hex, from Outlook's conversationId (EWS/Graph base64 — the GUID is
     its last 16 bytes; the front part names the mailbox, so it differs between Matt's and Chad's copies). */
  function conversationKey(raw) {
    var s = String(raw || "").trim(); if (!s) return "";
    if (/^[0-9a-f]{32}$/i.test(s)) return s.toLowerCase();
    try {
      var bytes = b64bytes(s.replace(/-/g, "+").replace(/_/g, "/")); if (bytes.length < 16) return "";
      var hex = ""; for (var i = bytes.length - 16; i < bytes.length; i++) hex += pad2(bytes[i].toString(16));
      return hex;
    } catch (e) { return ""; }
  }

  /* POST to the flow; always settles (an answer object, an error object, or a timeout object). */
  function callFlow(url, payload, timeoutMs) {
    var ms = timeoutMs || 20000;
    var ctl = null; try { if (typeof AbortController !== "undefined") ctl = new AbortController(); } catch (e) { ctl = null; }
    var opts = { method: "POST", headers: { "Content-Type": "text/plain;charset=UTF-8" }, body: JSON.stringify(payload) };
    if (ctl) opts.signal = ctl.signal;
    var request;
    try {
      request = fetch(url, opts).then(function (r) {
        return r.text().then(function (t) { var j = {}; try { j = JSON.parse(t); } catch (e) { j = { ok: false, error: "The flow answered oddly (" + r.status + ")." }; } if (!r.ok && !j.error) j.error = "The flow answered " + r.status + "."; return j; });
      });
    } catch (e) { return Promise.resolve({ ok: false, error: "no web requests from here (" + errText(e) + ")" }); }
    var t = null;
    var clock = new Promise(function (res) { t = timer(function () { if (ctl) { try { ctl.abort(); } catch (e) { /* nothing */ } } res({ ok: false, error: "the flow did not answer within " + Math.round(ms / 1000) + " s" }); }, ms); });
    var raced = t === null ? request : Promise.race([request, clock]);
    return raced.then(function (v) { cancel(t); return v; }, function (e) { cancel(t); return { ok: false, error: errText(e) || "the flow could not be reached" }; });
  }
  function lite(kind, s) {
    var me = ""; try { me = (Office.context.mailbox.userProfile.emailAddress || "").toLowerCase(); } catch (e) { me = ""; }
    return { schema: "contract-queue/2", kind: kind, id: "cq-" + kind + "-" + stamp() + "-" + rand(4), submittedBy: me, submittedAt: isoLocal(new Date()), listTitle: (s && s.listTitle) || "Contract Review Queue" };
  }
  function probe(url, s, what) {           // DIAG only: an empty thread lookup carrying a marker, never awaited
    if (!DIAG || !url) return;
    try { callFlow(url, Object.assign(lite("thread", s), { conversationKey: "", probe: what }), 8000).then(function () {}, function () {}); } catch (e) { /* nothing */ }
  }

  // Office.js callbacks → promises, each bounded
  function recipients(field) { return bounded(new Promise(function (res) { try { field.getAsync(function (r) { res(r.status === Office.AsyncResultStatus.Succeeded ? (r.value || []).map(function (p) { return (p.emailAddress || "").toLowerCase(); }).filter(Boolean) : []); }); } catch (e) { res([]); } }), 6000, []); }
  function attachments(item) { return bounded(new Promise(function (res) { try { item.getAttachmentsAsync(function (r) { res(r.status === Office.AsyncResultStatus.Succeeded ? (r.value || []).filter(function (a) { return !a.isInline; }) : []); }); } catch (e) { res([]); } }), 6000, []); }
  function customProps(item) { return bounded(new Promise(function (res) { try { item.loadCustomPropertiesAsync(function (r) { res(r.status === Office.AsyncResultStatus.Succeeded ? r.value : null); }); } catch (e) { res(null); } }), 6000, null); }
  function saveProps(props) { return bounded(new Promise(function (res) { try { props.saveAsync(function () { res(true); }); } catch (e) { res(false); } }), 6000, false); }
  function notify(item, key, message, isError) {
    try {
      item.notificationMessages.replaceAsync(key, isError ? { type: "errorMessage", message: String(message).slice(0, 150) }
        : { type: "informationalMessage", message: String(message).slice(0, 150), icon: "Icon.16", persistent: false }, function () {});
    } catch (e) { /* no notification area here */ }
  }
  function supports(v) { try { return Office.context.requirements.isSetSupported("Mailbox", v); } catch (e) { return false; } }

  /* The thread's open entry, if any: {match, outside, key} — or null when nothing should happen. */
  function openEntryFor(item, s, url) {
    var key = conversationKey(item.conversationId);
    if (!key) return Promise.resolve(null);
    return Promise.all([recipients(item.to), recipients(item.cc), attachments(item)]).then(function (parts) {
      var legal = legalAddresses(s);
      var outside = parts[0].concat(parts[1]).filter(function (a) { return legal.indexOf(a) < 0; });
      if (!outside.length || !parts[2].length) return null;
      return callFlow(url, Object.assign(lite("thread", s), { conversationKey: key }), 12000).then(function (res) {
        var m = (res && res.matches) || [];
        return m.length ? { match: m[0], outside: outside, key: key } : null;
      });
    });
  }

  /* The Smart Alert's contextData: on the event itself, else through getInitializationContextAsync (the
     documented route for the button's task pane or function), else {} — the compose-ribbon button was pressed
     directly and the thread key does the lookup. */
  function contextOf(event, item) {
    var parse = function (v) { if (!v) return null; if (typeof v === "object") return v; try { return JSON.parse(v) || null; } catch (e) { return null; } };
    var direct = parse(event && event.contextData);
    if (direct) return Promise.resolve(direct);
    return bounded(new Promise(function (res) {
      try { item.getInitializationContextAsync(function (r) { res((r.status === Office.AsyncResultStatus.Succeeded && parse(r.value)) || {}); }); }
      catch (e) { res({}); }
    }), 4000, {});
  }

  // --- Send pressed ---------------------------------------------------------------------------
  function onMessageSend(event) {
    var finished = false, watchdog = null;
    var complete = function (opts) { if (finished) return; finished = true; cancel(watchdog); try { event.completed(opts); } catch (e) { /* nothing */ } };
    var allow = function () { complete({ allowEvent: true }); };
    try {
      var item = Office.context.mailbox.item;
      var url = roam("flowUrl");
      var s = settings();
      watchdog = timer(allow, 25000);        // our own trouble never holds an email back
      if (!url || !item) { allow(); return; }
      customProps(item).then(function (props) {
        if (props && props.get(SKIP)) { props.remove(SKIP); return saveProps(props).then(allow); }   // already decided for this email
        return openEntryFor(item, s, url).then(function (found) {
          if (finished) return;
          if (!found) { allow(); return; }
          var m = found.match;
          var cut = function (t, n) { t = String(t || ""); return t.length > n ? t.slice(0, n - 1) + "…" : t; };
          var who = cut(found.outside.slice(0, 2).join(", ") + (found.outside.length > 2 ? " and " + (found.outside.length - 2) + " more" : ""), 90);
          // the Smart Alerts box shows at most 500 characters
          var text = "Going to " + who + ".\nIn the queue: “" + cut(m.Title || "(no title)", 70) + "”" + (m.Assigned ? " (" + cut(m.Assigned, 30) + ")" : "") + ".\n\n" +
            "Send anyway = send WITHOUT marking the contract returned.\n" +
            "Mark returned & send = today's date goes in Date Returned, the row leaves the queue, then press Send once more.";
          var opts = { allowEvent: false, errorMessage: text };
          if (supports("1.14")) {
            opts.cancelLabel = "Mark returned & send";
            opts.commandId = "cqMarkReturned";
            opts.contextData = JSON.stringify({ itemId: m.Id, title: m.Title || "", key: found.key });
          } else {
            opts.errorMessage = text + "\n\n(No second button in this Outlook: use the add-in's Mark returned button, then send.)";
          }
          opts.errorMessage = opts.errorMessage.slice(0, 500);
          complete(opts);
        });
      }).then(null, function () { allow(); });
    } catch (e) { allow(); }
  }

  // --- "Mark returned & send" (also the compose-ribbon button "Mark returned") -----------------
  function cqMarkReturned(event) {
    var item = null, finished = false, watchdog = null;
    var done = function () { if (finished) return; finished = true; cancel(watchdog); try { event.completed(); } catch (e) { /* nothing */ } };
    try {
      item = Office.context.mailbox.item;
      notify(item, "cq", "Mark returned: starting…", false);
      var url = roam("flowUrl");
      var s = settings();
      probe(url, s, "mark-entry");
      watchdog = timer(function () {         // whatever stalled, the command ends within a minute
        notify(item, "cq", "Mark returned took too long here — use the task pane's Mark returned button.", true); done();
      }, 60000);
      if (!url || !item) { notify(item, "cq", "The add-in is not connected to the flow yet — open the task pane once.", true); done(); return; }
      contextOf(event, item).then(function (ctx) {
        var key = ctx.key || conversationKey(item.conversationId);
        var payload = Object.assign(lite("mark-returned", s), { itemId: ctx.itemId || 0, conversationKey: key || "" });
        if (!payload.itemId && !payload.conversationKey) { notify(item, "cq", "This email is not part of a thread the queue knows.", true); return; }
        notify(item, "cq", "Mark returned: asking the flow…", false);
        return callFlow(url, payload, 40000).then(function (res) {
          if (res && res.ok) {
            var r = res.returned, log = (r && typeof r === "object") ? (r.ok ? "Date Returned written." : "Workbook: " + (r.error || "not written.")) : "";
            return customProps(item).then(function (props) {
              if (props) { props.set(SKIP, true); return saveProps(props); }
            }).then(function () {
              // informational messages hold 150 characters: the instruction first, the details after
              notify(item, "cq", "Marked returned. Press Send to send this email. (" + (res.title || ctx.title || "the contract") + " left the queue. " + log + ")", false);
            });
          }
          var why = (res && res.error) || "the flow did not answer";
          if (res && Array.isArray(res.matches) && res.matches.length > 1) why = "Several contracts share this thread — use the task pane to pick one.";
          notify(item, "cq", "Not marked returned: " + why, true);
        });
      }).then(null, function (e) { notify(item, "cq", "Not marked returned: " + errText(e), true); }).then(done, done);
    } catch (e) {
      try { notify(item, "cq", "Mark returned hit an error here: " + errText(e), true); } catch (e2) { /* nothing */ }
      done();
    }
  }

  Office.actions.associate("onMessageSend", onMessageSend);
  Office.actions.associate("cqMarkReturned", cqMarkReturned);
  try { var u0 = roam("flowUrl"); if (u0) probe(u0, settings(), "load"); } catch (e) { /* nothing */ }
})();
