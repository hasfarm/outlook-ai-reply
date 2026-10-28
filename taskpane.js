/*
 * AI Reply (Claude) - Outlook add-in task pane
 * Viết bằng ES5 để chạy được trên Outlook 2016 (trình duyệt nhúng IE11):
 * không dùng arrow function, let/const, fetch, Promise, template string.
 */
(function () {
  "use strict";

  var API_URL = "https://api.anthropic.com/v1/messages";
  var MAX_BODY_CHARS = 40000;

  var state = {
    bodyText: "",
    meta: null,
    history: []   // hội thoại với Claude để có thể "chỉnh lại"
  };

  function $(id) { return document.getElementById(id); }

  // ---------- Khởi động ----------
  Office.onReady(function () {
    loadSettings();
    bindEvents();
    loadItem();

    // Khi ghim (pin) pane và chuyển sang email khác (Outlook hỗ trợ Mailbox 1.5+)
    if (Office.context.mailbox.addHandlerAsync && Office.EventType && Office.EventType.ItemChanged) {
      Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, function () {
        resetDraft();
        loadItem();
      });
    }
  });

  function bindEvents() {
    $("btnSettings").onclick = function () {
      var s = $("settings");
      s.style.display = (s.style.display === "none") ? "block" : "none";
    };
    $("btnSaveSettings").onclick = saveSettings;
    $("btnGenerate").onclick = generate;
    $("btnRefine").onclick = refine;
    $("refine").onkeydown = function (e) {
      if ((e.keyCode || e.which) === 13) { refine(); }
    };
    $("btnReplyAll").onclick = function () { openReply(true); };
    $("btnReply").onclick = function () { openReply(false); };
    $("lang").onchange = savePrefs;
    $("tone").onchange = savePrefs;
  }

  // ---------- Cài đặt (lưu trong mailbox, đi theo tài khoản) ----------
  function rs() { return Office.context.roamingSettings; }

  function loadSettings() {
    var key = rs().get("apiKey") || "";
    $("apiKey").value = key;
    $("model").value = rs().get("model") || "claude-sonnet-5-5";
    var sign = rs().get("signName");
    if (sign === undefined || sign === null) {
      sign = (Office.context.mailbox.userProfile && Office.context.mailbox.userProfile.displayName) || "";
    }
    $("signName").value = sign;
    $("persona").value = rs().get("persona") || "";
    $("lang").value = rs().get("lang") || "auto";
    $("tone").value = rs().get("tone") || "professional";
    if (!key) {
      $("settings").style.display = "block";
      $("settingsMsg").innerText = "Nhập API key để bắt đầu";
    }
  }

  function saveSettings() {
    rs().set("apiKey", trim($("apiKey").value));
    rs().set("model", $("model").value);
    rs().set("signName", trim($("signName").value));
    rs().set("persona", trim($("persona").value));
    rs().saveAsync(function (r) {
      if (r.status === Office.AsyncResultStatus.Succeeded) {
        $("settingsMsg").innerText = "Đã lưu";
        setTimeout(function () { $("settings").style.display = "none"; $("settingsMsg").innerText = ""; }, 800);
      } else {
        $("settingsMsg").innerText = "Lưu lỗi: " + r.error.message;
      }
    });
  }

  function savePrefs() {
    rs().set("lang", $("lang").value);
    rs().set("tone", $("tone").value);
    rs().saveAsync(function () {});
  }

  // ---------- Đọc email đang chọn ----------
  function loadItem() {
    var item = Office.context.mailbox.item;
    state.bodyText = "";
    if (!item) {
      $("mailFrom").innerText = "Chưa chọn email";
      $("mailSubject").innerText = "";
      return;
    }
    var from = item.from || item.sender || {};
    state.meta = {
      fromName: from.displayName || "",
      fromEmail: from.emailAddress || "",
      to: joinRecipients(item.to),
      cc: joinRecipients(item.cc),
      subject: item.subject || "",
      date: item.dateTimeCreated ? item.dateTimeCreated.toString() : ""
    };
    $("mailFrom").innerText = state.meta.fromName || state.meta.fromEmail || "(không rõ người gửi)";
    $("mailSubject").innerText = state.meta.subject;

    item.body.getAsync(Office.CoercionType.Text, function (r) {
      if (r.status === Office.AsyncResultStatus.Succeeded) {
        state.bodyText = r.value || "";
      } else {
        showStatus("Không đọc được nội dung email: " + r.error.message, true);
      }
    });
  }

  function joinRecipients(list) {
    if (!list || !list.length) { return ""; }
    var out = [];
    for (var i = 0; i < list.length; i++) {
      out.push(list[i].displayName || list[i].emailAddress);
    }
    return out.join("; ");
  }

  // ---------- Gọi Claude ----------
  var TONES = {
    professional: "professional, clear and courteous",
    friendly: "warm and friendly but still professional",
    brief: "short and to the point (3-6 sentences)",
    formal: "formal and respectful"
  };

  function buildSystemPrompt() {
    var me = (Office.context.mailbox.userProfile && Office.context.mailbox.userProfile.displayName) || "the user";
    var persona = trim($("persona").value);
    var sign = trim($("signName").value);
    var lang = $("lang").value;

    var langRule = lang === "vi" ? "Write the reply in Vietnamese."
      : lang === "en" ? "Write the reply in English."
      : "Write the reply in the same language as the most recent message in the thread.";

    return [
      "You are an email assistant drafting a reply on behalf of " + me + (persona ? " (" + persona + ")" : "") + ".",
      "The user will review and edit your draft before sending.",
      langRule,
      "Tone: " + TONES[$("tone").value] + ".",
      "Reply to the most recent message in the thread; use earlier messages only as context.",
      "Address the sender's actual questions and requests. Do not invent facts, dates, numbers or commitments.",
      "If something must be decided or filled in by the user, put a short placeholder in square brackets, e.g. [ngày giờ].",
      sign ? "End with an appropriate closing and the name: " + sign + "."
           : "End with an appropriate closing line but do NOT add a name or signature (Outlook adds it).",
      "Output ONLY the email body as plain text: greeting, content, closing. No subject line, no markdown, no explanations."
    ].join("\n");
  }

  function buildUserPrompt() {
    var m = state.meta || {};
    var body = state.bodyText;
    if (body.length > MAX_BODY_CHARS) {
      body = body.substring(0, MAX_BODY_CHARS) + "\n[...thread truncated...]";
    }
    var notes = trim($("notes").value);
    return "From: " + m.fromName + " <" + m.fromEmail + ">\n" +
      "To: " + m.to + "\n" +
      (m.cc ? "Cc: " + m.cc + "\n" : "") +
      "Subject: " + m.subject + "\n" +
      "Date: " + m.date + "\n\n" +
      "----- EMAIL THREAD -----\n" + body + "\n----- END -----\n\n" +
      (notes ? "My key points for the reply: " + notes
             : "I have no specific notes; draft the most sensible reply.");
  }

  function generate() {
    if (!checkReady()) { return; }
    state.history = [{ role: "user", content: buildUserPrompt() }];
    callClaude("Claude đang đọc email và soạn trả lời...");
  }

  function refine() {
    var instr = trim($("refine").value);
    if (!instr) { $("refine").focus(); return; }
    if (!checkReady()) { return; }
    if (!state.history.length) { state.history = [{ role: "user", content: buildUserPrompt() }]; }
    // Lấy bản nháp hiện tại (kể cả phần người dùng đã tự sửa)
    var last = state.history[state.history.length - 1];
    if (last.role === "assistant") { last.content = $("draft").value; }
    else { state.history.push({ role: "assistant", content: $("draft").value }); }
    state.history.push({ role: "user", content: "Revise the draft above: " + instr + "\nKeep any edits I made. Output only the revised email body." });
    callClaude("Claude đang chỉnh lại bản nháp...");
  }

  function checkReady() {
    if (!trim($("apiKey").value)) {
      $("settings").style.display = "block";
      showStatus("Chưa có API key. Nhập key trong phần Cài đặt rồi bấm Lưu.", true);
      return false;
    }
    if (!state.bodyText && !(state.meta && state.meta.subject)) {
      showStatus("Chưa đọc được nội dung email. Thử chọn lại email.", true);
      return false;
    }
    return true;
  }

  function callClaude(loadingText) {
    setBusy(true);
    showStatus(loadingText, false);

    var payload = {
      model: $("model").value,
      max_tokens: 4096,
      system: buildSystemPrompt(),
      messages: state.history
    };

    var xhr = new XMLHttpRequest();
    xhr.open("POST", API_URL, true);
    xhr.setRequestHeader("content-type", "application/json");
    xhr.setRequestHeader("x-api-key", trim($("apiKey").value));
    xhr.setRequestHeader("anthropic-version", "2023-06-01");
    xhr.setRequestHeader("anthropic-dangerous-direct-browser-access", "true");
    xhr.timeout = 180000;
    var done = false;

    xhr.onreadystatechange = function () {
      if (xhr.readyState !== 4 || done) { return; }
      done = true;
      setBusy(false);
      var data = null;
      try { data = JSON.parse(xhr.responseText); } catch (e) { data = null; }

      if (xhr.status === 200 && data && data.content) {
        var text = "";
        for (var i = 0; i < data.content.length; i++) {
          if (data.content[i].type === "text") { text += data.content[i].text; }
        }
        state.history.push({ role: "assistant", content: text });
        $("draft").value = trim(text);
        $("draftBox").style.display = "block";
        $("refine").value = "";
        hideStatus();
        return;
      }

      // lượt bị lỗi -> bỏ yêu cầu vừa gửi khỏi lịch sử
      state.history.pop();

      var msg;
      if (xhr.status === 0) {
        msg = "Không kết nối được tới api.anthropic.com. Kiểm tra mạng / tường lửa công ty.";
      } else if (xhr.status === 401) {
        msg = "API key không hợp lệ. Mở Cài đặt và nhập lại key.";
      } else if (xhr.status === 429) {
        msg = "Gọi quá nhiều lần trong thời gian ngắn. Đợi một chút rồi thử lại.";
      } else {
        msg = "Lỗi " + xhr.status + ": " + ((data && data.error && data.error.message) || xhr.responseText);
      }
      showStatus(msg, true);
    };
    xhr.ontimeout = function () {
      if (done) { return; }
      done = true; setBusy(false); state.history.pop(); showStatus("Hết thời gian chờ Claude. Thử lại.", true);
    };

    xhr.send(JSON.stringify(payload));
  }

  // ---------- Mở cửa sổ trả lời của Outlook ----------
  function openReply(all) {
    var text = $("draft").value;
    if (!trim(text)) { return; }
    var html = '<div style="font-family:Calibri,Arial,sans-serif;font-size:11pt">' +
      escapeHtml(text).replace(/\r?\n/g, "<br>") + "</div>";
    var item = Office.context.mailbox.item;
    try {
      if (all) { item.displayReplyAllForm(html); }
      else { item.displayReplyForm(html); }
    } catch (e) {
      showStatus("Không mở được cửa sổ trả lời: " + e.message, true);
    }
  }

  // ---------- Tiện ích ----------
  function resetDraft() {
    state.history = [];
    $("draft").value = "";
    $("notes").value = "";
    $("draftBox").style.display = "none";
    hideStatus();
  }

  function setBusy(b) {
    var ids = ["btnGenerate", "btnRefine", "btnReply", "btnReplyAll"];
    for (var i = 0; i < ids.length; i++) { $(ids[i]).disabled = b; }
  }

  function showStatus(msg, isError) {
    var s = $("status");
    s.innerText = msg;
    s.className = isError ? "status error" : "status";
    s.style.display = "block";
  }
  function hideStatus() { $("status").style.display = "none"; }

  function trim(s) { return (s || "").replace(/^\s+|\s+$/g, ""); }

  function escapeHtml(s) {
    return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  }
})();
