/*
 * Follow-Up Reminder add-in
 * Flow: user opens an email -> clicks "Set Follow-Up" ribbon button ->
 * this task pane opens -> user picks a date -> we create a calendar
 * appointment pre-filled with the email subject + a CLICKABLE link back to the email.
 */

let currentSubject = "";
let emailLink = "";

Office.onReady(function (info) {
  if (info.host !== Office.HostType.Outlook) {
    return;
  }

  const item = Office.context.mailbox.item;

  // 1. Grab the subject of the email we're looking at
  currentSubject = item.subject || "(no subject)";
  document.getElementById("subjectText").textContent = currentSubject;

  // 2. Build a clickable link back to this exact email (opens in Outlook on the web)
  try {
    const encodedId = encodeURIComponent(item.itemId);
    emailLink =
      "https://outlook.office365.com/owa/?ItemID=" +
      encodedId +
      "&exvsurl=1&viewmodel=ReadMessageItem";
  } catch (e) {
    emailLink = "";
  }

  // 3. Default the date box to tomorrow
  setDateOffset(1);

  // Quick-pick buttons (Tomorrow / In 3 days / Next week)
  const quickButtons = document.querySelectorAll(".quick-row button");
  quickButtons.forEach(function (btn) {
    btn.addEventListener("click", function () {
      setDateOffset(parseInt(btn.getAttribute("data-days"), 10));
    });
  });

  // Main action
  document.getElementById("setBtn").addEventListener("click", createReminder);
});

// Set the date input to today + n days
function setDateOffset(days) {
  const d = new Date();
  d.setDate(d.getDate() + days);
  const yyyy = d.getFullYear();
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const dd = String(d.getDate()).padStart(2, "0");
  document.getElementById("followDate").value = yyyy + "-" + mm + "-" + dd;
}

// Escape any special characters in the subject so the HTML body stays valid
function escapeHtml(text) {
  return String(text)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function createReminder() {
  const dateVal = document.getElementById("followDate").value;
  if (!dateVal) {
    showStatus("err", "Pick a date first.");
    return;
  }

  // Parse the chosen date as a local date
  const parts = dateVal.split("-");
  const start = new Date(
    parseInt(parts[0], 10),
    parseInt(parts[1], 10) - 1,
    parseInt(parts[2], 10),
    8, 0, 0 // 8:00 AM on that day
  );
  const end = new Date(start.getTime());
  end.setMinutes(end.getMinutes() + 30);

  const safeSubject = escapeHtml(currentSubject);

  // ---- HTML body so the link is actually CLICKABLE ----
  let body =
    "<div style=\"font-family:Segoe UI,Arial,sans-serif;font-size:14px;\">" +
    "<p>Follow up on this email.</p>" +
    "<p><strong>Original subject:</strong> " + safeSubject + "</p>";

  if (emailLink) {
    body +=
      "<p>\uD83D\uDC49 <a href=\"" + emailLink + "\">Open the original email</a></p>";
  } else {
    body +=
      "<p>(Open Outlook and search the subject line above to find the email.)</p>";
  }
  body += "</div>";

  try {
    // Opens a pre-filled appointment form. User just clicks Save.
    Office.context.mailbox.displayNewAppointmentForm({
      requiredAttendees: [],
      optionalAttendees: [],
      start: start,
      end: end,
      subject: "Follow-Up: " + currentSubject,
      body: body,
      location: ""
    });
    showStatus(
      "ok",
      "Reminder ready \u2014 just click Save on the appointment that opened. It'll sit on your calendar for " +
        dateVal + " with a clickable link back to this email."
    );
  } catch (e) {
    showStatus("err", "Couldn't open the reminder. Try again or check the console.");
  }
}

function showStatus(type, msg) {
  const el = document.getElementById("status");
  el.className = type;
  el.textContent = msg;
}
