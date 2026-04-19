/* ================================================================
   NDL Issue Tracker — Outlook Add-in Task Pane
   Reads selected email → sends to Power Automate HTTP trigger
   ================================================================ */

// ── CONFIGURATION ─────────────────────────────────────────────────
// Replace this URL with your Power Automate "When an HTTP request is received" trigger URL
const FLOW_ENDPOINT = "https://default5e8309eec8d04bc7b5b85a37a4eb10.7b.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/27e63548f67749afaff088af1e123a8b/triggers/manual/paths/invoke?api-version=1";
// ──────────────────────────────────────────────────────────────────

let emailData = {
  subject: "",
  from: "",
  body: "",
  bodyText: "",
  attachments: [],
  messageId: "",
};

// ── Office ready ──────────────────────────────────────────────────
Office.onReady(function (info) {
  if (info.host === Office.HostType.Outlook) {
    loadEmailData();
    document.getElementById("meetingSelect").addEventListener("change", onMeetingChange);
    document.getElementById("submitBtn").addEventListener("click", onSubmit);
  }
});

// ── Load email data from the selected message ─────────────────────
function loadEmailData() {
  var item = Office.context.mailbox.item;

  // Subject
  emailData.subject = item.subject || "(No subject)";
  document.getElementById("emailSubject").textContent = emailData.subject;
  document.getElementById("issueTitle").value = emailData.subject;

  // From
  if (item.from) {
    emailData.from = item.from.displayName + " <" + item.from.emailAddress + ">";
  }
  document.getElementById("emailFrom").textContent = emailData.from;

  // Message ID for attachment retrieval
  emailData.messageId = item.itemId;

  // Body (async)
  item.body.getAsync(Office.CoercionType.Text, function (result) {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      emailData.bodyText = result.value;
      var preview = result.value.substring(0, 200);
      if (result.value.length > 200) preview += "...";
      document.getElementById("emailBodyPreview").textContent = preview;
      document.getElementById("issueDesc").value = result.value.substring(0, 2000);
    }
  });

  // HTML body (for richer description)
  item.body.getAsync(Office.CoercionType.Html, function (result) {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      emailData.body = result.value;
    }
  });

  // Attachments
  loadAttachments(item);
}

// ── Load attachments ──────────────────────────────────────────────
function loadAttachments(item) {
  var attachments = item.attachments;
  var attachmentsContainer = document.getElementById("emailAttachments");
  var toggleRow = document.getElementById("attachToggleRow");

  if (!attachments || attachments.length === 0) {
    attachmentsContainer.style.display = "none";
    toggleRow.style.display = "none";
    return;
  }

  // Show attachment chips and toggle
  attachmentsContainer.style.display = "flex";
  toggleRow.style.display = "flex";
  attachmentsContainer.innerHTML = "";

  emailData.attachments = [];

  for (var i = 0; i < attachments.length; i++) {
    var att = attachments[i];
    // Skip inline images (embedded in body)
    if (att.isInline) continue;

    emailData.attachments.push({
      id: att.id,
      name: att.name,
      size: att.size,
      contentType: att.contentType,
    });

    var chip = document.createElement("div");
    chip.className = "email-att-chip";
    chip.innerHTML = getFileIcon(att.name) + " " + att.name +
      " <span style='color:#9ca3af;font-size:10px;'>(" + formatSize(att.size) + ")</span>";
    attachmentsContainer.appendChild(chip);
  }

  if (emailData.attachments.length === 0) {
    attachmentsContainer.style.display = "none";
    toggleRow.style.display = "none";
  }
}

// ── Meeting selection change ──────────────────────────────────────
function onMeetingChange() {
  var meeting = document.getElementById("meetingSelect").value;
  var btn = document.getElementById("submitBtn");
  if (meeting) {
    btn.disabled = false;
    btn.textContent = "Create Issue";
  } else {
    btn.disabled = true;
    btn.textContent = "Select a meeting to continue";
  }
}

// ── Submit: send to Power Automate ────────────────────────────────
function onSubmit() {
  var meeting = document.getElementById("meetingSelect").value;
  var title = document.getElementById("issueTitle").value.trim();
  var desc = document.getElementById("issueDesc").value.trim();
  var includeAtt = document.getElementById("includeAttachments").checked;

  if (!meeting || !title) return;

  setStatus("loading", "Creating issue...");
  var btn = document.getElementById("submitBtn");
  btn.disabled = true;
  btn.textContent = "Creating...";

  // If we need to include attachments, fetch their content first
  if (includeAtt && emailData.attachments.length > 0) {
    fetchAttachmentContents(function (attachmentContents) {
      sendToFlow(meeting, title, desc, attachmentContents);
    });
  } else {
    sendToFlow(meeting, title, desc, []);
  }
}

// ── Fetch attachment binary content via Office JS ─────────────────
function fetchAttachmentContents(callback) {
  var item = Office.context.mailbox.item;
  var results = [];
  var remaining = emailData.attachments.length;

  if (remaining === 0) {
    callback([]);
    return;
  }

  for (var i = 0; i < emailData.attachments.length; i++) {
    (function (att) {
      item.getAttachmentContentAsync(att.id, function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          results.push({
            name: att.name,
            contentType: att.contentType,
            content: result.value.content,
            format: result.value.format, // base64 or other
          });
        }
        remaining--;
        if (remaining === 0) {
          callback(results);
        }
      });
    })(emailData.attachments[i]);
  }
}

// ── Send payload to Power Automate HTTP trigger ───────────────────
function sendToFlow(meeting, title, desc, attachments) {
  var payload = {
    meeting: meeting,
    title: title,
    description: desc,
    emailFrom: emailData.from,
    emailSubject: emailData.subject,
    createdBy: Office.context.mailbox.userProfile.displayName,
    createdByEmail: Office.context.mailbox.userProfile.emailAddress,
    attachments: attachments.map(function (a) {
      return {
        fileName: a.name,
        contentType: a.contentType,
        contentBytes: a.content, // base64
      };
    }),
  };

  fetch(FLOW_ENDPOINT, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  })
    .then(function (response) {
      if (response.ok) {
        return response.json().catch(function () { return {}; });
      }
      throw new Error("Flow returned " + response.status);
    })
    .then(function (data) {
      setStatus("success", "Issue created! ✓");
      var btn = document.getElementById("submitBtn");
      btn.textContent = "Issue Created ✓";

      // Auto-close after 2 seconds
      setTimeout(function () {
        setStatus("", "");
        btn.disabled = false;
        btn.textContent = "Create Another";
      }, 3000);
    })
    .catch(function (err) {
      setStatus("error", "Failed: " + err.message);
      var btn = document.getElementById("submitBtn");
      btn.disabled = false;
      btn.textContent = "Retry";
    });
}

// ── Helpers ────────────────────────────────────────────────────────
function setStatus(type, msg) {
  var el = document.getElementById("statusMsg");
  el.className = "status" + (type ? " " + type : "");
  el.textContent = msg;
  el.style.display = msg ? "block" : "none";
}

function getFileIcon(filename) {
  var ext = (filename || "").split(".").pop().toLowerCase();
  var icons = {
    pdf: "📄", doc: "📝", docx: "📝", xls: "📊", xlsx: "📊",
    ppt: "📽️", pptx: "📽️", png: "🖼️", jpg: "🖼️", jpeg: "🖼️",
    gif: "🖼️", zip: "📦", rar: "📦", txt: "📃", csv: "📊",
  };
  return icons[ext] || "📎";
}

function formatSize(bytes) {
  if (bytes < 1024) return bytes + " B";
  if (bytes < 1048576) return (bytes / 1024).toFixed(0) + " KB";
  return (bytes / 1048576).toFixed(1) + " MB";
}
