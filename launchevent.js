/* Outlook Add In by Psomas — the on-send watcher and the "Mark returned" command (v2.2).

   Classic Outlook runs this file on its own (no page, no DOM), so it is self-contained: nothing
   from taskpane.js is imported. Outlook on the web and the new Outlook load it through commands.html.

   When the reviewer presses Send on a reply that (1) goes to someone outside Legal and (2) carries an
   attachment (signature pictures do not count) and (3) belongs to an email thread with an open entry in the
   Contract Review Queue, a box
   appears: "Send anyway" sends without touching the queue; "Mark returned & send" stamps Date
   Returned in the Contracts Review Log, removes the row from the queue, and asks for one more press of
   Send. Replies that stay inside Legal (a first pass handed to Chad) never see the box.
*/
(function () {
  "use strict";
  var LEGAL = ["matthew.pipes@psomas.com", "chad.wilson@psomas.com", "paul.atreides@psomas.com", "legal@psomas.com"];
  var SKIP = "cqSkipMark";                 // set on the email once it has been marked (or the user chose not to)

  function roam(key) { try { return Office.context.roamingSettings.get(key); } catch (e) { return undefined; } }
  function settings() { try { var c = roam("configCache"); return c ? (JSON.parse(c).settings || {}) : {}; } catch (e) { return {}; } }
  function legalAddresses(s) {
    var list = LEGAL.concat(s.supervisors || [], [s.paulEmail || ""], s.internalAddresses || []);
    return list.map(function (a) { return String(a || "").trim().toLowerCase(); }).filter(Boolean);
  }
  function isoLocal(d) {
    var off = -d.getTimezoneOffset(), sign = off >= 0 ? "+" : "-", a = Math.abs(off), p = function (n) { return String(n).padStart(2, "0"); };
    return d.getFullYear() + "-" + p(d.getMonth() + 1) + "-" + p(d.getDate()) + "T" + p(d.getHours()) + ":" + p(d.getMinutes()) + ":" + p(d.getSeconds()) + sign + p(Math.floor(a / 60)) + ":" + p(a % 60);
  }
  function stamp() { var d = new Date(), p = function (n) { return String(n).padStart(2, "0"); }; return "" + d.getFullYear() + p(d.getMonth() + 1) + p(d.getDate()) + p(d.getHours()) + p(d.getMinutes()) + p(d.getSeconds()); }
  function rand(n) { var s = ""; while (s.length < n) s += Math.random().toString(36).slice(2); return s.slice(0, n); }

  /* The 16-byte conversation GUID as hex, from Outlook's conversationId (EWS/Graph base64 — the GUID is
     its last 16 bytes; the front part names the mailbox, so it differs between Matt's and Chad's copies). */
  function conversationKey(raw) {
    var s = String(raw || "").trim(); if (!s) return "";
    if (/^[0-9a-f]{32}$/i.test(s)) return s.toLowerCase();
    try {
      var b64 = s.replace(/-/g, "+").replace(/_/g, "/"); b64 += "=".repeat((4 - b64.length % 4) % 4);
      var bin = atob(b64); if (bin.length < 16) return "";
      var hex = ""; for (var i = bin.length - 16; i < bin.length; i++) hex += bin.charCodeAt(i).toString(16).padStart(2, "0");
      return hex;
    } catch (e) { return ""; }
  }

  function callFlow(url, payload, timeoutMs) {
    var ctl = typeof AbortController !== "undefined" ? new AbortController() : null;
    var timer = ctl ? setTimeout(function () { ctl.abort(); }, timeoutMs || 20000) : null;
    return fetch(url, { method: "POST", headers: { "Content-Type": "text/plain;charset=UTF-8" }, body: JSON.stringify(payload), signal: ctl ? ctl.signal : undefined })
      .then(function (r) { return r.text().then(function (t) { var j = {}; try { j = JSON.parse(t); } catch (e) { j = { ok: false, error: "The flow answered oddly (" + r.status + ")." }; } if (!r.ok && !j.error) j.error = "The flow answered " + r.status + "."; return j; }); })
      .finally(function () { if (timer) clearTimeout(timer); });
  }
  function lite(kind, s) {
    var me = ""; try { me = (Office.context.mailbox.userProfile.emailAddress || "").toLowerCase(); } catch (e) { me = ""; }
    return { schema: "contract-queue/2", kind: kind, id: "cq-" + kind + "-" + stamp() + "-" + rand(4), submittedBy: me, submittedAt: isoLocal(new Date()), listTitle: (s && s.listTitle) || "Contract Review Queue" };
  }

  // Office.js callbacks → promises
  function recipients(field) { return new Promise(function (res) { try { field.getAsync(function (r) { res(r.status === Office.AsyncResultStatus.Succeeded ? (r.value || []).map(function (p) { return (p.emailAddress || "").toLowerCase(); }).filter(Boolean) : []); }); } catch (e) { res([]); } }); }
  function attachments(item) { return new Promise(function (res) { try { item.getAttachmentsAsync(function (r) { res(r.status === Office.AsyncResultStatus.Succeeded ? (r.value || []).filter(function (a) { return !a.isInline; }) : []); }); } catch (e) { res([]); } }); }
  function customProps(item) { return new Promise(function (res) { try { item.loadCustomPropertiesAsync(function (r) { res(r.status === Office.AsyncResultStatus.Succeeded ? r.value : null); }); } catch (e) { res(null); } }); }
  function saveProps(props) { return new Promise(function (res) { try { props.saveAsync(function () { res(); }); } catch (e) { res(); } }); }
  function notify(item, key, message, isError) {
    try {
      item.notificationMessages.replaceAsync(key, isError ? { type: "errorMessage", message: String(message).slice(0, 150) }
        : { type: "informationalMessage", message: String(message).slice(0, 150), icon: "Icon.16", persistent: false }, function () {});
    } catch (e) { /* no notification area here */ }
  }
  function supports(v) { try { return Office.context.requirements.isSetSupported("Mailbox", v); } catch (e) { return false; } }

  /* The thread's open entry, if any: {match, outside} — or null when nothing should happen. */
  function openEntryFor(item, s, url) {
    var key = conversationKey(item.conversationId);
    if (!key) return Promise.resolve(null);
    return Promise.all([recipients(item.to), recipients(item.cc), attachments(item)]).then(function (parts) {
      var legal = legalAddresses(s);
      var outside = parts[0].concat(parts[1]).filter(function (a) { return legal.indexOf(a) < 0; });
      if (!outside.length || !parts[2].length) return null;
      return callFlow(url, Object.assign(lite("thread", s), { conversationKey: key }), 20000).then(function (res) {
        var m = (res && res.matches) || [];
        return m.length ? { match: m[0], outside: outside, key: key } : null;
      });
    });
  }

  // --- Send pressed ---------------------------------------------------------------------------
  function onMessageSend(event) {
    var item = Office.context.mailbox.item;
    var url = roam("flowUrl");
    var s = settings();
    var allow = function () { event.completed({ allowEvent: true }); };
    if (!url || !item) return allow();
    customProps(item).then(function (props) {
      if (props && props.get(SKIP)) { props.remove(SKIP); return saveProps(props).then(allow); }   // already decided for this email
      return openEntryFor(item, s, url).then(function (found) {
        if (!found) return allow();
        var m = found.match;
        var cut = function (t, n) { t = String(t || ""); return t.length > n ? t.slice(0, n - 1) + "\u2026" : t; };
        var who = cut(found.outside.slice(0, 2).join(", ") + (found.outside.length > 2 ? " and " + (found.outside.length - 2) + " more" : ""), 90);
        // the Smart Alerts box shows at most 500 characters
        var text = "Going to " + who + ".\nIn the queue: \u201c" + cut(m.Title || "(no title)", 70) + "\u201d" + (m.Assigned ? " (" + cut(m.Assigned, 30) + ")" : "") + ".\n\n" +
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
        event.completed(opts);
      });
    }).catch(function () { allow(); });   // never hold an email back because of our own trouble
  }

  /* The Smart Alert's contextData: on the event itself, else through getInitializationContextAsync (the
     documented route for the button's task pane or function), else {} — the compose-ribbon button was pressed
     directly and the thread key does the lookup. */
  function contextOf(event, item) {
    var parse = function (v) { if (!v) return null; if (typeof v === "object") return v; try { return JSON.parse(v) || null; } catch (e) { return null; } };
    var direct = parse(event && event.contextData);
    if (direct) return Promise.resolve(direct);
    return new Promise(function (res) {
      try { item.getInitializationContextAsync(function (r) { res((r.status === Office.AsyncResultStatus.Succeeded && parse(r.value)) || {}); }); }
      catch (e) { res({}); }
    });
  }

  // --- "Mark returned & send" (also the compose-ribbon button "Mark returned") ---------------
  function cqMarkReturned(event) {
    var item = Office.context.mailbox.item;
    var url = roam("flowUrl");
    var s = settings();
    var done = function () { try { event.completed(); } catch (e) { /* nothing */ } };
    if (!url || !item) { notify(item, "cq", "The add-in is not connected to the flow yet — open the task pane once.", true); return done(); }
    contextOf(event, item).then(function (ctx) {
      var key = ctx.key || conversationKey(item.conversationId);
      var payload = Object.assign(lite("mark-returned", s), { itemId: ctx.itemId || 0, conversationKey: key || "" });
      if (!payload.itemId && !payload.conversationKey) { notify(item, "cq", "This email is not part of a thread the queue knows.", true); return; }
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
    }).catch(function (e) { notify(item, "cq", "Not marked returned: " + (e && e.message ? e.message : e), true); }).then(done);
  }

  Office.actions.associate("onMessageSend", onMessageSend);
  Office.actions.associate("cqMarkReturned", cqMarkReturned);
})();
