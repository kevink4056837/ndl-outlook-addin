/* ================================================================
   NDL Issue Tracker — Outlook Add-in Task Pane
   Reads selected email → sends to Power Automate HTTP trigger
   ================================================================ */

// ── CONFIGURATION ─────────────────────────────────────────────────
// Replace this URL with your Power Automate "When an HTTP request is received" trigger URL
const FLOW_ENDPOINT = "https://default5e8309eec8d04bc7b5b85a37a4eb10.7b.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/27e63548f67749afaff088af1e123a8b/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=sIJbuWBl4vFYUkY-4bUvl93tWZlEY6WL3y7dHpB-mLw";
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

  // Get HTML body (preserves formatting + images)
  item.body.getAsync(Office.CoercionType.Html, function (result) {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      emailData.body = result.value;
      // Build plain text preview for the task pane UI
      var plainText = htmlToCleanText(result.value);
      emailData.bodyText = plainText;
      var preview = plainText.substring(0, 200);
      if (plainText.length > 200) preview += "...";
      document.getElementById("emailBodyPreview").textContent = preview;
      document.getElementById("issueDesc").value = plainText.substring(0, 2000);

      // Resolve inline images (cid: references → base64 data URLs)
      resolveInlineImages(item, function (resolvedHtml) {
        emailData.body = resolvedHtml;
      });
    } else {
      item.body.getAsync(Office.CoercionType.Text, function (textResult) {
        if (textResult.status === Office.AsyncResultStatus.Succeeded) {
          emailData.bodyText = textResult.value;
          var preview = textResult.value.substring(0, 200);
          if (textResult.value.length > 200) preview += "...";
          document.getElementById("emailBodyPreview").textContent = preview;
          document.getElementById("issueDesc").value = textResult.value.substring(0, 2000);
        }
      });
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
    htmlDescription: emailData.body || desc,
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

// ── Resolve inline images: replace cid: with base64 data URLs ─────
function resolveInlineImages(item, callback) {
  var html = emailData.body;
  var attachments = item.attachments;
  if (!attachments || attachments.length === 0) {
    callback(html);
    return;
  }

  // Find inline attachments
  var inlineAtts = [];
  for (var i = 0; i < attachments.length; i++) {
    if (attachments[i].isInline && attachments[i].contentType && attachments[i].contentType.indexOf("image") === 0) {
      inlineAtts.push(attachments[i]);
    }
  }

  if (inlineAtts.length === 0) {
    callback(html);
    return;
  }

  var remaining = inlineAtts.length;
  var replacements = {};

  for (var j = 0; j < inlineAtts.length; j++) {
    (function (att) {
      item.getAttachmentContentAsync(att.id, function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          var dataUrl = "data:" + att.contentType + ";base64," + result.value.content;
          // Replace cid references — Outlook uses cid:filename or cid:contentId
          html = html.replace(new RegExp('src=["\']cid:' + escapeRegex(att.name) + '["\']', 'gi'), 'src="' + dataUrl + '"');
          // Also try matching by content ID (without the name)
          if (att.id) {
            html = html.replace(new RegExp('src=["\']cid:[^"\']*' + escapeRegex(att.name.split('.')[0]) + '[^"\']*["\']', 'gi'), 'src="' + dataUrl + '"');
          }
        }
        remaining--;
        if (remaining === 0) {
          callback(html);
        }
      });
    })(inlineAtts[j]);
  }
}

function escapeRegex(str) {
  return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function htmlToCleanText(html) {
  if (!html) return "";
  var text = html;
  // Convert block elements to line breaks
  text = text.replace(/<br\s*\/?>/gi, "\n");
  text = text.replace(/<\/p>/gi, "\n\n");
  text = text.replace(/<\/div>/gi, "\n");
  text = text.replace(/<\/h[1-6]>/gi, "\n\n");
  text = text.replace(/<\/li>/gi, "\n");
  text = text.replace(/<\/tr>/gi, "\n");
  text = text.replace(/<hr[^>]*>/gi, "\n---\n");
  // Strip all remaining HTML tags
  text = text.replace(/<[^>]+>/g, "");
  // Decode common HTML entities
  text = text.replace(/&nbsp;/gi, " ");
  text = text.replace(/&amp;/gi, "&");
  text = text.replace(/&lt;/gi, "<");
  text = text.replace(/&gt;/gi, ">");
  text = text.replace(/&quot;/gi, '"');
  text = text.replace(/&#39;/gi, "'");
  text = text.replace(/&rsquo;/gi, "'");
  text = text.replace(/&lsquo;/gi, "'");
  text = text.replace(/&rdquo;/gi, '"');
  text = text.replace(/&ldquo;/gi, '"');
  text = text.replace(/&mdash;/gi, "—");
  text = text.replace(/&ndash;/gi, "–");
  // Clean up excessive whitespace but preserve intentional line breaks
  text = text.replace(/[ \t]+/g, " ");
  text = text.replace(/\n /g, "\n");
  text = text.replace(/ \n/g, "\n");
  text = text.replace(/\n{3,}/g, "\n\n");
  return text.trim();
}

function cleanEmailBody(text) {
  if (!text) return "";
  var body = text;
  // Split on common reply/forward markers (with or without leading newline)
  var markers = [
    /\n?\s*From:\s*.+\n?\s*Sent:\s*.+\n?\s*To:/i,   // Outlook reply header
    /From:\s*[^\n]*<[^>]+>\s*Sent:/i,                  // Inline "From: Name <email>Sent:"
    /-{2,}\s*Original Message\s*-{2,}/i,               // ---- Original Message ----
    /On .{10,80}wrote:\s*/i,                            // Gmail-style "On ... wrote:"
    /Get Outlook for (Android|iOS|Mobile)/i,            // "Get Outlook for Android/iOS"
    /_{10,}/,                                           // long underscores (signature divider)
    /-{10,}/,                                           // long dashes (signature divider)
    /Sent from my (iPhone|iPad|Galaxy|Android)/i,       // mobile signatures
  ];
  for (var i = 0; i < markers.length; i++) {
    var match = body.search(markers[i]);
    if (match > 0) {
      body = body.substring(0, match);
    }
  }
  // Trim signature block — cut at common sign-off patterns
  var sigMarkers = [
    /Thanks[,!]?\s*\n?\s*[\w].*\|/i,                   // "Thanks, Name | Title"
    /Thanks[,!]?\s*\n?\s*[\w].*\nD:/i,                 // "Thanks!\nKevin Kimball\nD: 603..."
    /\nRegards[,]?\s*\n/i,
    /\nBest[,]?\s*\n/i,
    /\nCheers[,]?\s*\n/i,
    /\nThank you[,]?\s*\n/i,
  ];
  for (var j = 0; j < sigMarkers.length; j++) {
    var sigMatch = body.search(sigMarkers[j]);
    if (sigMatch > 0) {
      body = body.substring(0, sigMatch);
    }
  }
  // Clean up whitespace
  body = body.replace(/\r\n/g, "\n").replace(/\n{3,}/g, "\n\n").trim();
  return body;
}

function formatSize(bytes) {
  if (bytes < 1024) return bytes + " B";
  if (bytes < 1048576) return (bytes / 1024).toFixed(0) + " KB";
  return (bytes / 1048576).toFixed(1) + " MB";
}
