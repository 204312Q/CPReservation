const RESERVATION_SHEET_NAME = "Joo Chiat Reservation";
const BLOCKING_SHEET_NAME = "Joo Chiat_Blocking";
const PUBLIC_HOLIDAY_SHEET_NAME = "Public Holiday";
const SPREADSHEET_ID = "1KMhTLxhmrvz-ili7oj8NMrC-wTSxT-x7fqjEtNeHdNo";
const MAX_PAX_PER_SLOT = 60;
const STAFF_NOTIFICATION_EMAIL = "chillipadinonyarestaurant63@gmail.com";

function getSheets_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  return {
    reservation: ss.getSheetByName(RESERVATION_SHEET_NAME),
    blocking: ss.getSheetByName(BLOCKING_SHEET_NAME),
    publicHoliday: ss.getSheetByName(PUBLIC_HOLIDAY_SHEET_NAME),
  };
}

// Generate Reservation ID
function getReservationID() {
  const chars = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789";
  let id = "";

  for (let i = 0; i < 7; i++) {
    const index = Math.floor(Math.random() * chars.length);
    id += chars[index];
  }
  return id;
}

function isActive_(v) {
  if (v === true) return true;
  const s = String(v || "").trim().toLowerCase();
  return ["active", "true", "yes", "y", "1"].includes(s);
}

function escapeHtml_(text) {
  return String(text ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function normalizeDateStr_(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return String(v || "").trim();
}

function normalizeTimeStr_(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), "HH:mm");
  }
  const s = String(v || "").trim();
  return s.length >= 5 ? s.slice(0, 5) : s;
}

function getWebAppUrl_() {
  let webAppUrl = ScriptApp.getService().getUrl();
  return webAppUrl.replace(/\/u\/\d+\//, "/");
}

function getManageButtonHtml_(reservationId) {
  const manageLink = `${getWebAppUrl_()}?resId=${encodeURIComponent(reservationId)}`;
  return `
<div style="padding:20px 0;display:block;text-align:left;">
  <a href="${manageLink}" 
     style="background-color:#8B0000;
            color:#ffffff;
            padding:14px 28px;
            text-decoration:none;
            border-radius:5px;
            display:inline-block;
            font-weight:bold;
            font-size:15px;
            line-height:1;">
     Manage / Cancel Reservation
  </a>
</div>`;
}

function getNewReservationButtonHtml_() {
  const newReservationLink = getWebAppUrl_();
  return `
<div style="padding:20px 0;display:block;text-align:left;">
  <a href="${newReservationLink}" 
     style="background-color:#8B0000;
            color:#ffffff;
            padding:14px 28px;
            text-decoration:none;
            border-radius:5px;
            display:inline-block;
            font-weight:bold;
            font-size:15px;
            line-height:1;">
     Make a New Reservation
  </a>
</div>`;
}

function getBlockStateForDate_(dateStr) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(BLOCKING_SHEET_NAME);
  if (!sh) return { closedAllDay: false, blockedSlots: {} };

  const values = sh.getDataRange().getValues();
  const tz = Session.getScriptTimeZone();

  const toDateStr = (d) =>
    d instanceof Date
      ? Utilities.formatDate(d, tz, "yyyy-MM-dd")
      : String(d || "").trim();

  const toTimeStr = (t) => {
    if (!t) return "";
    if (t instanceof Date) return Utilities.formatDate(t, tz, "HH:mm");
    const s = String(t).trim();
    return s.length >= 5 ? s.slice(0, 5) : s;
  };

  const toMinutes = (hhmm) => {
    const [h, m] = hhmm.split(":").map(Number);
    return h * 60 + m;
  };

  const interval = 15;
  const blockedSlots = {};
  let closedAllDay = false;

  for (let i = 1; i < values.length; i++) {
    const rowDate = toDateStr(values[i][0]); // A Date
    const start = toTimeStr(values[i][1]);   // B Start
    const end = toTimeStr(values[i][2]);     // C End
    const active = values[i][4];             // E Active

    if (!isActive_(active)) continue;
    if (rowDate !== dateStr) continue;

    // Whole-day block when Start/End are blank
    if (!start && !end) {
      closedAllDay = true;
      break;
    }

    // Time-range block when both Start and End exist
    if (start && end) {
      const startMin = toMinutes(start);
      const endMin = toMinutes(end);

      for (let t = startMin; t <= endMin; t += interval) {
        const hh = String(Math.floor(t / 60)).padStart(2, "0");
        const mm = String(t % 60).padStart(2, "0");
        blockedSlots[`${dateStr}|${hh}:${mm}`] = true;
      }
    }
  }

  return { closedAllDay, blockedSlots };
}

function getPaxByTimeForDate_(dateStr) {
  const { reservation } = getSheets_();
  const lastRow = reservation.getLastRow();
  if (lastRow < 2) return {};

  const values = reservation.getRange(2, 1, lastRow - 1, 11).getValues();
  const paxByTime = {};

  for (const r of values) {
    const status = String(r[1] || "").toUpperCase().trim(); // B Status
    const dateCell = r[6]; // G Date
    const timeCell = r[7]; // H Time
    if (!dateCell || !timeCell) continue;

    if (status === "CANCELLED" || status === "NO-SHOW") continue;

    const rowDateStr = (dateCell instanceof Date)
      ? Utilities.formatDate(dateCell, Session.getScriptTimeZone(), "yyyy-MM-dd")
      : String(dateCell).trim();

    if (rowDateStr !== dateStr) continue;

    let timeStr;
    if (timeCell instanceof Date) {
      timeStr = Utilities.formatDate(timeCell, Session.getScriptTimeZone(), "HH:mm");
    } else {
      const s = String(timeCell).trim();
      timeStr = s.length >= 5 ? s.slice(0, 5) : s;
    }

    const adults = Number(r[8] || 0);    // I Adults
    const children = Number(r[9] || 0);  // J Children
    const pax = adults + children;

    paxByTime[timeStr] = (paxByTime[timeStr] || 0) + pax;
  }

  return paxByTime;
}

// Block table expected columns: Date | Time | Reason
function getBlockedMap() {
  const { blocking } = getSheets_();
  const values = blocking.getDataRange().getValues();
  const map = {};

  for (let i = 1; i < values.length; i++) {
    const date = values[i][0]; // col A
    const time = values[i][1]; // col B
    if (!date || !time) continue;

    const dateStr = (date instanceof Date)
      ? Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd")
      : String(date).trim();

    const timeStr = String(time).trim();
    map[`${dateStr}|${timeStr}`] = true;
  }
  return map;
}

// Fetch all public holidays from Google Sheet
function getPublicHolidayMap_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(PUBLIC_HOLIDAY_SHEET_NAME);
  if (!sh) return {};

  const values = sh.getDataRange().getValues();
  const map = {};

  for (let i = 1; i < values.length; i++) {
    const d = values[i][0];
    const name = values[i][1];
    if (!d) continue;

    const dateStr = (d instanceof Date)
      ? Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd")
      : String(d).trim();

    map[dateStr] = name ? String(name).trim() : "Public Holiday";
  }
  return map;
}

function isPublicHoliday(dateStr) {
  const phMap = getPublicHolidayMap_();
  return Boolean(phMap[dateStr]);
}

function getAvailableTimes(dateStr) {
  const blockState = getBlockStateForDate_(dateStr);
  if (blockState.closedAllDay) return [];

  const blockedMap = blockState.blockedSlots;
  const paxByTime = getPaxByTimeForDate_(dateStr);

  const dateObj = new Date(dateStr + "T00:00:00");
  let day = dateObj.getDay(); // 0=Sun ... 6=Sat

  // Public holiday treated as weekend
  if (isPublicHoliday(dateStr)) {
    day = 0;
  }

  const isWeekend = (day === 0 || day === 6);

  const toMinutes = (t) => {
    const [h, m] = t.split(":").map(Number);
    return h * 60 + m;
  };

  const lunch = { start: "11:30", end: "14:15" };
  const dinner = isWeekend
    ? { start: "17:30", end: "21:30" }
    : { start: "17:30", end: "21:00" };

  const periods = [lunch, dinner];
  const interval = 15;
  const results = [];

  for (const p of periods) {
    const startMinutes = toMinutes(p.start);
    const endMinutes = toMinutes(p.end);

    for (let t = startMinutes; t <= endMinutes; t += interval) {
      const hh = String(Math.floor(t / 60)).padStart(2, "0");
      const mm = String(t % 60).padStart(2, "0");
      const timeStr = `${hh}:${mm}`;

      if (blockedMap[`${dateStr}|${timeStr}`]) continue;

      const currentPax = paxByTime[timeStr] || 0;
      if (currentPax >= MAX_PAX_PER_SLOT) continue;

      results.push(timeStr);
    }
  }

  return results;
}

// Submit Reservation
function submitReservation(payload) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const reservation = ss.getSheetByName(RESERVATION_SHEET_NAME);

  if (!payload.firstName || !payload.lastName || !payload.phone || !payload.date || !payload.time) {
    throw new Error("Missing required fields.");
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);

  try {
    const blockState = getBlockStateForDate_(payload.date);

    if (blockState.closedAllDay) {
      throw new Error("This date is unavailable. Please choose another date.");
    }

    if (blockState.blockedSlots[`${payload.date}|${payload.time}`]) {
      throw new Error("This time slot is blocked. Please choose another time.");
    }

    const paxByTime = getPaxByTimeForDate_(payload.date);
    const currentPax = paxByTime[payload.time] || 0;
    const incomingPax = Number(payload.adults || 0) + Number(payload.children || 0);

    if (currentPax + incomingPax > MAX_PAX_PER_SLOT) {
      throw new Error("This time slot is fully booked (capacity reached). Please choose another time.");
    }

    const reservationId = getReservationID();

    reservation.appendRow([
      reservationId,                 // A Reservation ID
      "PENDING",                     // B Status
      payload.firstName,             // C First Name
      payload.lastName,              // D Last Name
      payload.email || "",           // E Email
      payload.phone,                 // F Phone
      payload.date,                  // G Date
      payload.time,                  // H Time
      Number(payload.adults || 0),   // I Adults
      Number(payload.children || 0), // J Children
      payload.notes || ""            // K Additional Request
    ]);

    sendReservationEmail_(reservationId, payload);
    sendStaffNotificationEmail_("NEW", reservationId, payload);

    return { ok: true, reservationId };
  } finally {
    lock.releaseLock();
  }
}

function sendReservationEmail_(reservationId, payload) {
  if (!payload.email) return;

  const adults = Number(payload.adults || 0);
  const children = Number(payload.children || 0);
  const manageLink = `${getWebAppUrl_()}?resId=${encodeURIComponent(reservationId)}`;
  const buttonHtml = getManageButtonHtml_(reservationId);

  const subject = `Chilli Padi Reservation (${reservationId})`;

  const plainBody =
    `Dear ${payload.firstName || "Guest"}${payload.lastName ? " " + payload.lastName : ""},

Thanks for choosing Chilli Padi Nonya Restaurant. We're super excited to have you.

Here are your reservation details:
Reservation ID: ${reservationId}
Date: ${payload.date}
Time: ${payload.time}
No. Of Adults: ${adults}
No. Of Children: ${children}
Additional Request: ${payload.notes || "None"}

Our Location:
11 Joo Chiat Place #01-03 Singapore 427744
Tel: +65 6275 1002

To modify or cancel your reservation:
${manageLink}

Chilli Padi`;

  const htmlBody = `
<div style="margin:0;padding:0;background:#f4f6f8;font-family:Arial,Helvetica,sans-serif;">
  <table width="100%" cellpadding="0" cellspacing="0" style="background:#f4f6f8;padding:30px 10px;">
    <tr>
      <td align="center">
        <table width="100%" cellpadding="0" cellspacing="0"
          style="max-width:600px;background:#ffffff;border-radius:12px;
                 box-shadow:0 8px 24px rgba(0,0,0,0.06);overflow:hidden;">

          <tr>
            <td style="background:#8B0000;padding:20px 24px;color:#ffffff;">
              <h2 style="margin:0;font-size:20px;">
                Chilli Padi Nonya Restaurant
              </h2>
              <p style="margin:6px 0 0 0;font-size:13px;opacity:0.9;">
                Reservation Confirmation
              </p>
            </td>
          </tr>

          <tr>
            <td style="padding:24px;color:#333333;font-size:14px;line-height:1.6;">
              <p style="margin-top:0;">
                Dear ${escapeHtml_(payload.firstName || "Guest")}${payload.lastName ? " " + escapeHtml_(payload.lastName) : ""},
              </p>

              <p>
                Thank you for choosing <b>Chilli Padi Nonya Restaurant</b>.
                We’re excited to welcome you!
              </p>

              <table width="100%" cellpadding="8" cellspacing="0"
                style="margin:20px 0;border:1px solid #eeeeee;border-radius:8px;">
                <tr style="background:#fafafa;">
                  <td colspan="2" style="font-weight:bold;">
                    Reservation Details
                  </td>
                </tr>

                <tr>
                  <td width="40%" style="color:#666;"><b>Reservation ID</b></td>
                  <td>${escapeHtml_(reservationId)}</td>
                </tr>
                <tr>
                  <td style="color:#666;"><b>Date</b></td>
                  <td>${escapeHtml_(payload.date)}</td>
                </tr>
                <tr>
                  <td style="color:#666;"><b>Time</b></td>
                  <td>${escapeHtml_(payload.time)}</td>
                </tr>
                <tr>
                  <td style="color:#666;"><b>Adults</b></td>
                  <td>${adults}</td>
                </tr>
                <tr>
                  <td style="color:#666;"><b>Children</b></td>
                  <td>${children}</td>
                </tr>
                <tr>
                  <td style="color:#666;"><b>Additional Request</b></td>
                  <td>${escapeHtml_(payload.notes || "None")}</td>
                </tr>
              </table>

              <p>
                <b>Location</b><br>
                11 Joo Chiat Place #01-03 Singapore 427744<br>
                Tel: +65 6275 1002
              </p>

              <p>
                If you need to modify or cancel your reservation,
                please click on the button below:
                <br>
                ${buttonHtml}
              </p>

              <p style="margin-bottom:0;">
                We look forward to serving you.
              </p>
            </td>
          </tr>

          <tr>
            <td style="background:#fafafa;padding:16px;text-align:center;
                       font-size:12px;color:#888;">
              © ${new Date().getFullYear()} Chilli Padi Nonya Restaurant<br>
              This is an automated confirmation email. Please do not reply.
            </td>
          </tr>

        </table>
      </td>
    </tr>
  </table>
</div>`;

  MailApp.sendEmail({
    to: payload.email,
    subject,
    body: plainBody,
    htmlBody,
    name: "No Reply - Chilli Padi Nonya Restaurant"
  });
}

function sendReservationUpdatedEmail_(reservationId, payload, oldData) {
  if (!payload.email) return;

  const adults = Number(payload.adults || 0);
  const children = Number(payload.children || 0);
  const manageLink = `${getWebAppUrl_()}?resId=${encodeURIComponent(reservationId)}`;
  const buttonHtml = getManageButtonHtml_(reservationId);

  const subject = `Chilli Padi Reservation Updated (${reservationId})`;

  const plainBody =
    `Dear ${payload.firstName || "Guest"}${payload.lastName ? " " + payload.lastName : ""},

Your reservation at Chilli Padi Nonya Restaurant has been successfully updated.

Previous details:
Reservation ID: ${reservationId}
Date: ${oldData.date}
Time: ${oldData.time}
Adults: ${oldData.adults}
Children: ${oldData.children}
Additional Request: ${oldData.notes || "None"}

Updated details:
Reservation ID: ${reservationId}
Date: ${payload.date}
Time: ${payload.time}
Adults: ${adults}
Children: ${children}
Additional Request: ${payload.notes || "None"}

Manage your reservation:
${manageLink}

Chilli Padi`;

  const htmlBody = `
<div style="margin:0;padding:0;background:#f4f6f8;font-family:Arial,Helvetica,sans-serif;">
  <table width="100%" cellpadding="0" cellspacing="0" style="background:#f4f6f8;padding:30px 10px;">
    <tr>
      <td align="center">
        <table width="100%" cellpadding="0" cellspacing="0"
          style="max-width:600px;background:#ffffff;border-radius:12px;
                 box-shadow:0 8px 24px rgba(0,0,0,0.06);overflow:hidden;">

          <tr>
            <td style="background:#8B0000;padding:20px 24px;color:#ffffff;">
              <h2 style="margin:0;font-size:20px;">Chilli Padi Nonya Restaurant</h2>
              <p style="margin:6px 0 0 0;font-size:13px;opacity:0.9;">Reservation Updated</p>
            </td>
          </tr>

          <tr>
            <td style="padding:24px;color:#333333;font-size:14px;line-height:1.6;">
              <p style="margin-top:0;">
                Dear ${escapeHtml_(payload.firstName || "Guest")}${payload.lastName ? " " + escapeHtml_(payload.lastName) : ""},
              </p>

              <p>Your reservation has been successfully updated.</p>

              <table width="100%" cellpadding="8" cellspacing="0"
                style="margin:20px 0;border:1px solid #eeeeee;border-radius:8px;">
                <tr style="background:#fafafa;">
                  <td colspan="2" style="font-weight:bold;">Previous Details</td>
                </tr>
                <tr><td width="40%"><b>Reservation ID</b></td><td>${escapeHtml_(reservationId)}</td></tr>
                <tr><td width="40%"><b>Date</b></td><td>${escapeHtml_(oldData.date)}</td></tr>
                <tr><td><b>Time</b></td><td>${escapeHtml_(oldData.time)}</td></tr>
                <tr><td><b>Adults</b></td><td>${Number(oldData.adults || 0)}</td></tr>
                <tr><td><b>Children</b></td><td>${Number(oldData.children || 0)}</td></tr>
                <tr><td><b>Additional Request</b></td><td>${escapeHtml_(oldData.notes || "None")}</td></tr>
              </table>

              <table width="100%" cellpadding="8" cellspacing="0"
                style="margin:20px 0;border:1px solid #eeeeee;border-radius:8px;">
                <tr style="background:#fafafa;">
                  <td colspan="2" style="font-weight:bold;">Updated Details</td>
                </tr>
                <tr><td width="40%"><b>Reservation ID</b></td><td>${escapeHtml_(reservationId)}</td></tr>
                <tr><td><b>Date</b></td><td>${escapeHtml_(payload.date)}</td></tr>
                <tr><td><b>Time</b></td><td>${escapeHtml_(payload.time)}</td></tr>
                <tr><td><b>Adults</b></td><td>${adults}</td></tr>
                <tr><td><b>Children</b></td><td>${children}</td></tr>
                <tr><td><b>Additional Request</b></td><td>${escapeHtml_(payload.notes || "None")}</td></tr>
              </table>

              <p>
                <b>Location</b><br>
                11 Joo Chiat Place #01-03 Singapore 427744<br>
                Tel: +65 6275 1002
              </p>

              <p>
                If you need to modify or cancel your reservation,
                please click on the button below:
                <br>
                ${buttonHtml}
              </p>

              <p style="margin-bottom:0;">We look forward to serving you.</p>
            </td>
          </tr>

          <tr>
            <td style="background:#fafafa;padding:16px;text-align:center;font-size:12px;color:#888;">
              © ${new Date().getFullYear()} Chilli Padi Nonya Restaurant<br>
              This is an automated email. Please do not reply.
            </td>
          </tr>

        </table>
      </td>
    </tr>
  </table>
</div>`;

  MailApp.sendEmail({
    to: payload.email,
    subject,
    body: plainBody,
    htmlBody,
    name: "No Reply - Chilli Padi Nonya Restaurant"
  });
}

function sendReservationCancelledEmail_(reservationId, payload) {
  if (!payload.email) return;

  const newReservationLink = getWebAppUrl_();
  const newReservationButton = getNewReservationButtonHtml_();

  const subject = `Chilli Padi Reservation Cancelled (${reservationId})`;

  const plainBody =
    `Dear ${payload.firstName || "Guest"}${payload.lastName ? " " + payload.lastName : ""},

Your reservation at Chilli Padi Nonya Restaurant has been cancelled.

Cancelled reservation details:
Reservation ID: ${reservationId}
Date: ${payload.date}
Time: ${payload.time}
Adults: ${Number(payload.adults || 0)}
Children: ${Number(payload.children || 0)}
Additional Request: ${payload.notes || "None"}

If this was a mistake, you can make a new reservation here:
${newReservationLink}

Chilli Padi`;

  const htmlBody = `
<div style="margin:0;padding:0;background:#f4f6f8;font-family:Arial,Helvetica,sans-serif;">
  <table width="100%" cellpadding="0" cellspacing="0" style="background:#f4f6f8;padding:30px 10px;">
    <tr>
      <td align="center">
        <table width="100%" cellpadding="0" cellspacing="0"
          style="max-width:600px;background:#ffffff;border-radius:12px;
                 box-shadow:0 8px 24px rgba(0,0,0,0.06);overflow:hidden;">

          <tr>
            <td style="background:#8B0000;padding:20px 24px;color:#ffffff;">
              <h2 style="margin:0;font-size:20px;">Chilli Padi Nonya Restaurant</h2>
              <p style="margin:6px 0 0 0;font-size:13px;opacity:0.9;">Reservation Cancelled</p>
            </td>
          </tr>

          <tr>
            <td style="padding:24px;color:#333333;font-size:14px;line-height:1.6;">
              <p style="margin-top:0;">
                Dear ${escapeHtml_(payload.firstName || "Guest")}${payload.lastName ? " " + escapeHtml_(payload.lastName) : ""},
              </p>

              <p>Your reservation has been cancelled successfully.</p>

              <table width="100%" cellpadding="8" cellspacing="0"
                style="margin:20px 0;border:1px solid #eeeeee;border-radius:8px;">
                <tr style="background:#fafafa;">
                  <td colspan="2" style="font-weight:bold;">Cancelled Reservation Details</td>
                </tr>
                <tr><td width="40%"><b>Reservation ID</b></td><td>${escapeHtml_(reservationId)}</td></tr>
                <tr><td><b>Date</b></td><td>${escapeHtml_(payload.date)}</td></tr>
                <tr><td><b>Time</b></td><td>${escapeHtml_(payload.time)}</td></tr>
                <tr><td><b>Adults</b></td><td>${Number(payload.adults || 0)}</td></tr>
                <tr><td><b>Children</b></td><td>${Number(payload.children || 0)}</td></tr>
                <tr><td><b>Additional Request</b></td><td>${escapeHtml_(payload.notes || "None")}</td></tr>
              </table>

                            <p>
                <b>Location</b><br>
                11 Joo Chiat Place #01-03 Singapore 427744<br>
                Tel: +65 6275 1002
              </p>

              <p>
                If this was a mistake, you can create a new reservation below:
              </p>

              ${newReservationButton}
            </td>
          </tr>

          <tr>
            <td style="background:#fafafa;padding:16px;text-align:center;font-size:12px;color:#888;">
              © ${new Date().getFullYear()} Chilli Padi Nonya Restaurant<br>
              This is an automated email. Please do not reply.
            </td>
          </tr>

        </table>
      </td>
    </tr>
  </table>
</div>`;

  MailApp.sendEmail({
    to: payload.email,
    subject,
    body: plainBody,
    htmlBody,
    name: "No Reply - Chilli Padi Nonya Restaurant"
  });
}

function sendStaffNotificationEmail_(type, reservationId, payload, oldData) {
  if (!STAFF_NOTIFICATION_EMAIL) return;

  const name = [payload.firstName || "", payload.lastName || ""].join(" ").trim();
  const adults = Number(payload.adults || 0);
  const children = Number(payload.children || 0);
  const totalPax = adults + children;

  let subject = "";
  let body = "";

  if (type === "NEW") {
    subject = `NEW Reservation - ${reservationId}`;
    body =
      `NEW RESERVATION

Reservation ID: ${reservationId}
Name: ${name}
Phone: ${payload.phone || ""}
Email: ${payload.email || ""}

Date: ${payload.date}
Time: ${payload.time}

Adults: ${adults}
Children: ${children}
Total Pax: ${totalPax}

Notes: ${payload.notes || "None"}`;
  }

  if (type === "UPDATE") {
    subject = `UPDATED Reservation - ${reservationId}`;
    body =
      `RESERVATION UPDATED

Reservation ID: ${reservationId}
Name: ${name}
Phone: ${payload.phone || ""}
Email: ${payload.email || ""}

Previous
Date: ${oldData.date}
Time: ${oldData.time}
Adults: ${oldData.adults}
Children: ${oldData.children}
Notes: ${oldData.notes || "None"}

Updated
Date: ${payload.date}
Time: ${payload.time}
Adults: ${adults}
Children: ${children}
Total Pax: ${totalPax}
Notes: ${payload.notes || "None"}`;
  }

  if (type === "CANCEL") {
    subject = `CANCELLED Reservation - ${reservationId}`;
    body =
      `RESERVATION CANCELLED

Reservation ID: ${reservationId}
Name: ${name}
Phone: ${payload.phone || ""}
Email: ${payload.email || ""}

Date: ${payload.date}
Time: ${payload.time}

Adults: ${adults}
Children: ${children}
Total Pax: ${totalPax}

Notes: ${payload.notes || "None"}`;
  }

  if (!subject) return;

  MailApp.sendEmail({
    to: STAFF_NOTIFICATION_EMAIL,
    subject: subject,
    body: body,
    name: "Chilli Padi Reservation System"
  });
}

function doGet(e) {
  const resId = e?.parameter?.resId || null;
  const template = HtmlService.createTemplateFromFile("Index");

  if (resId) {
    const sheet = SpreadsheetApp
      .openById(SPREADSHEET_ID)
      .getSheetByName(RESERVATION_SHEET_NAME);

    const data = sheet.getDataRange().getValues();
    let resData = null;

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0] || "") === String(resId)) {
        resData = {
          id: data[i][0],
          firstName: data[i][2],
          lastName: data[i][3],
          email: data[i][4],
          phone: data[i][5],
          date: data[i][6] instanceof Date
            ? Utilities.formatDate(data[i][6], Session.getScriptTimeZone(), "yyyy-MM-dd")
            : String(data[i][6] || "").trim(),
          time: data[i][7] instanceof Date
            ? Utilities.formatDate(data[i][7], Session.getScriptTimeZone(), "HH:mm")
            : String(data[i][7] || "").trim().slice(0, 5),
          adults: data[i][8],
          children: data[i][9],
          notes: data[i][10],
          status: data[i][1],
        };
        break;
      }
    }

    template.res = resData;
    template.resId = resId;
  } else {
    template.res = null;
    template.resId = null;
  }

  return template.evaluate()
    .setTitle("Restaurant Reservation")
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

function updateReservation(form) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(RESERVATION_SHEET_NAME);
  const data = sheet.getDataRange().getValues();

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);

  try {
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(form.resId)) {
        const row = i + 1;
        const currentStatus = String(data[i][1] || "").toUpperCase().trim();

        if (currentStatus === "CANCELLED") {
          throw new Error("This reservation has already been cancelled.");
        }

        const oldData = {
          date: normalizeDateStr_(data[i][6]),
          time: normalizeTimeStr_(data[i][7]),
          adults: Number(data[i][8] || 0),
          children: Number(data[i][9] || 0),
          notes: data[i][10] || ""
        };

        const updatedPayload = {
          firstName: data[i][2],
          lastName: data[i][3],
          email: data[i][4],
          phone: data[i][5],
          date: String(form.date || "").trim(),
          time: String(form.time || "").trim(),
          adults: Number(form.adults != null ? form.adults : data[i][8] || 0),
          children: Number(form.children != null ? form.children : data[i][9] || 0),
          notes: form.notes != null ? form.notes : (data[i][10] || "")
        };

        if (!updatedPayload.date || !updatedPayload.time) {
          throw new Error("Date and time are required.");
        }

        const blockState = getBlockStateForDate_(updatedPayload.date);
        if (blockState.closedAllDay) {
          throw new Error("This date is unavailable. Please choose another date.");
        }

        if (blockState.blockedSlots[`${updatedPayload.date}|${updatedPayload.time}`]) {
          throw new Error("This time slot is blocked. Please choose another time.");
        }

        const paxByTime = getPaxByTimeForDate_(updatedPayload.date);

        // Subtract this reservation's existing pax if recalculating same slot
        if (oldData.date === updatedPayload.date && oldData.time === updatedPayload.time) {
          paxByTime[updatedPayload.time] =
            Math.max(0, (paxByTime[updatedPayload.time] || 0) - (oldData.adults + oldData.children));
        }

        const newPax = updatedPayload.adults + updatedPayload.children;
        const currentPax = paxByTime[updatedPayload.time] || 0;

        if (currentPax + newPax > MAX_PAX_PER_SLOT) {
          throw new Error("This time slot is fully booked (capacity reached). Please choose another time.");
        }

        sheet.getRange(row, 7).setValue(updatedPayload.date);      // G Date
        sheet.getRange(row, 8).setValue(updatedPayload.time);      // H Time
        sheet.getRange(row, 9).setValue(updatedPayload.adults);    // I Adults
        sheet.getRange(row, 10).setValue(updatedPayload.children); // J Children
        sheet.getRange(row, 11).setValue(updatedPayload.notes);    // K Notes

        sendReservationUpdatedEmail_(form.resId, updatedPayload, oldData);
        sendStaffNotificationEmail_("UPDATE", form.resId, updatedPayload, oldData);

        return { ok: true, message: "Reservation updated successfully!" };
      }
    }

    throw new Error("Reservation not found.");
  } finally {
    lock.releaseLock();
  }
}

function cancelReservation(resId) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(RESERVATION_SHEET_NAME);
  const data = sheet.getDataRange().getValues();

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);

  try {
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(resId)) {
        const row = i + 1;
        const currentStatus = String(data[i][1] || "").toUpperCase().trim();

        if (currentStatus === "CANCELLED") {
          return "This reservation is already cancelled.";
        }

        const payload = {
          firstName: data[i][2],
          lastName: data[i][3],
          email: data[i][4],
          phone: data[i][5],
          date: normalizeDateStr_(data[i][6]),
          time: normalizeTimeStr_(data[i][7]),
          adults: Number(data[i][8] || 0),
          children: Number(data[i][9] || 0),
          notes: data[i][10] || ""
        };

        sheet.getRange(row, 2).setValue("CANCELLED"); // B Status

        sendReservationCancelledEmail_(resId, payload);
        sendStaffNotificationEmail_("CANCEL", resId, payload);

        return "Success: Your reservation is now cancelled.";
      }
    }

    throw new Error("Reservation not found.");
  } finally {
    lock.releaseLock();
  }
}