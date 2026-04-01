const RESERVATION_SHEET_NAME = "Joo Chiat Reservation";
const BLOCKING_SHEET_NAME = "Joo Chiat_Blocking";
const PUBLIC_HOLIDAY_SHEET_NAME = "Public Holiday";
const SPREADSHEET_ID = "1KMhTLxhmrvz-ili7oj8NMrC-wTSxT-x7fqjEtNeHdNo";
const MAX_PAX_PER_SLOT = 60;
const MAX_PAX_PER_RESERVATION = 10;
const STAFF_NOTIFICATION_EMAIL = "chillipadinonyarestaurant63@gmail.com";

// Use your exact LIVE /exec URL here
const WEB_APP_URL = "https://script.google.com/macros/s/AKfycbwdZqRTsorlzVAvg4Otryuy8XvNlE60Sc6W4z2ajhsAGC7RkdM7bdj_AJzAuKiUpbM/exec";

const RL = {
  // Existing reservation / manage flow
  MANAGE_MAX_ATTEMPTS: 10,
  MANAGE_WINDOW_SECONDS: 10 * 60, // 10 min

  // New reservation anti-spam
  NEW_RES_MAX_ATTEMPTS: 3,
  NEW_RES_WINDOW_SECONDS: 10 * 60, // 10 min

  // Duplicate submit guard
  DUPLICATE_WINDOW_SECONDS: 90 // 1.5 min
};

// Reservation sheet column map
const COL = {
  RES_ID: 1,           // A
  STATUS: 2,           // B
  FIRST_NAME: 3,       // C
  LAST_NAME: 4,        // D
  EMAIL: 5,            // E
  PHONE: 6,            // F
  DATE: 7,             // G
  TIME: 8,             // H
  PAX: 9,              // I
  NOTES: 10,           // J
  MANAGE_TOKEN: 11,    // K
  CREATED_AT: 12,      // L
  UPDATED_AT: 13,      // M
  CANCELLED_AT: 14,    // N
  TOKEN_EXPIRES_AT: 15 // O
};

function getSheets_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  return {
    reservation: ss.getSheetByName(RESERVATION_SHEET_NAME),
    blocking: ss.getSheetByName(BLOCKING_SHEET_NAME),
    publicHoliday: ss.getSheetByName(PUBLIC_HOLIDAY_SHEET_NAME),
  };
}

function getReservationSheet_() {
  const { reservation } = getSheets_();
  if (!reservation) throw new Error("Reservation sheet not found.");
  return reservation;
}

function getWebAppUrl_() {
  return WEB_APP_URL;
}

// Generate human-friendly Reservation ID
function getReservationID() {
  const chars = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789";
  let id = "";
  for (let i = 0; i < 7; i++) {
    const index = Math.floor(Math.random() * chars.length);
    id += chars[index];
  }
  return id;
}

// Generate secure manage token
function getManageToken_() {
  const raw = Utilities.getUuid().replace(/-/g, "") + Utilities.getUuid().replace(/-/g, "");
  return raw;
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

function normalizeEmailForLimit_(email) {
  return String(email || "").trim().toLowerCase();
}

function normalizePhoneForLimit_(phone) {
  return String(phone || "").replace(/\D/g, "");
}

function sha256_(text) {
  const raw = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(text || ""),
    Utilities.Charset.UTF_8
  );

  return raw.map(function (b) {
    const v = (b < 0 ? b + 256 : b).toString(16);
    return v.length === 1 ? "0" + v : v;
  }).join("");
}

function getRateLimitCount_(key) {
  const cache = CacheService.getScriptCache();
  return Number(cache.get(key) || 0);
}

function incrementRateLimit_(key, windowSeconds) {
  const cache = CacheService.getScriptCache();
  const current = Number(cache.get(key) || 0) + 1;
  cache.put(key, String(current), windowSeconds);
  return current;
}

function clearRateLimit_(key) {
  CacheService.getScriptCache().remove(key);
}

function assertWithinRateLimit_(key, maxAttempts, windowSeconds, message) {
  const current = getRateLimitCount_(key);
  if (current >= maxAttempts) {
    throw new Error(message || "Too many attempts. Please try again later.");
  }
  incrementRateLimit_(key, windowSeconds);
}

function getManageRateLimitKey_(resId) {
  return "rl_manage_" + sha256_(String(resId || "").trim());
}

function consumeManageAttempt_(resId) {
  const key = getManageRateLimitKey_(resId);
  assertWithinRateLimit_(
    key,
    RL.MANAGE_MAX_ATTEMPTS,
    RL.MANAGE_WINDOW_SECONDS,
    "Too many invalid reservation access attempts. Please try again later."
  );
}

function resetManageAttempts_(resId) {
  clearRateLimit_(getManageRateLimitKey_(resId));
}

function getNewReservationRateLimitKey_(payload) {
  const email = normalizeEmailForLimit_(payload.email);
  const phone = normalizePhoneForLimit_(payload.phone);
  const identity = phone || email || "unknown";
  return "rl_new_" + sha256_(identity);
}

function consumeNewReservationAttempt_(payload) {
  const key = getNewReservationRateLimitKey_(payload);

  assertWithinRateLimit_(
    key,
    RL.NEW_RES_MAX_ATTEMPTS,
    RL.NEW_RES_WINDOW_SECONDS,
    "Too many reservation attempts in a short time. Please wait a few minutes and try again."
  );
}

function getDuplicateReservationKey_(payload) {
  const signature = [
    normalizePhoneForLimit_(payload.phone),
    normalizeEmailForLimit_(payload.email),
    String(payload.date || "").trim(),
    String(payload.time || "").trim(),
    Number(payload.pax || 0)
  ].join("|");

  return "dup_new_" + sha256_(signature);
}

function assertNotDuplicateReservation_(payload) {
  const cache = CacheService.getScriptCache();
  const key = getDuplicateReservationKey_(payload);

  if (cache.get(key)) {
    throw new Error("A similar reservation was just submitted. Please wait a moment before trying again.");
  }

  cache.put(key, "1", RL.DUPLICATE_WINDOW_SECONDS);
}

function getNow_() {
  return new Date();
}

function buildManageLink_(reservationId, manageToken) {
  return `${getWebAppUrl_()}?resId=${encodeURIComponent(reservationId)}&token=${encodeURIComponent(manageToken)}`;
}

function getManageButtonHtml_(reservationId, manageToken) {
  const manageLink = buildManageLink_(reservationId, manageToken);
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

// Token Expiration Logic
function getTokenExpiryForReservation_(dateStr, timeStr) {
  return new Date(`${dateStr}T${timeStr}:00`);
}

function isTokenExpired_(expiryValue) {
  if (!expiryValue) return true;

  const expiry = expiryValue instanceof Date ? expiryValue : new Date(expiryValue);
  if (isNaN(expiry.getTime())) return true;

  return Date.now() > expiry.getTime();
}

// Reservation Time Validation
function toComparableDateTime_(dateStr, timeStr) {
  return new Date(`${dateStr}T${timeStr}:00`);
}

function isPastReservation_(dateStr, timeStr) {
  const dt = toComparableDateTime_(dateStr, timeStr);
  return dt.getTime() < Date.now();
}

// Form Validation
function validateReservationPayload_(payload) {
  if (!payload.firstName || !payload.lastName || !payload.phone || !payload.date || !payload.time) {
    throw new Error("Missing required fields.");
  }

  const pax = Number(payload.pax || 0);

  if (pax <= 0) {
    throw new Error("Please select at least 1 guest.");
  }

  if (pax > MAX_PAX_PER_RESERVATION) {
    throw new Error("Guest count exceeds allowed maximum, Max 10 Pax.");
  }
}

// Find reservation by ID and token, return null if not found or token expired
function findReservationByIdAndToken_(resId, token) {
  const sheet = getReservationSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  const data = sheet.getRange(2, 1, lastRow - 1, COL.TOKEN_EXPIRES_AT).getValues();

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const rowResId = String(row[COL.RES_ID - 1] || "").trim();
    const rowToken = String(row[COL.MANAGE_TOKEN - 1] || "").trim();
    const tokenExpiresAt = row[COL.TOKEN_EXPIRES_AT - 1];

    if (rowResId !== String(resId).trim()) continue;
    if (rowToken !== String(token).trim()) continue;
    if (!rowToken) return null;
    if (isTokenExpired_(tokenExpiresAt)) return null;

    return {
      rowNumber: i + 2,
      values: row
    };
  }

  return null;
}

// Convert a sheet row to a reservation record object
function getReservationRecord_(row) {
  return {
    id: row[COL.RES_ID - 1],
    status: row[COL.STATUS - 1],
    firstName: row[COL.FIRST_NAME - 1],
    lastName: row[COL.LAST_NAME - 1],
    email: row[COL.EMAIL - 1],
    phone: row[COL.PHONE - 1],
    date: normalizeDateStr_(row[COL.DATE - 1]),
    time: normalizeTimeStr_(row[COL.TIME - 1]),
    pax: Number(row[COL.PAX - 1] || 0),
    notes: row[COL.NOTES - 1] || "",
    manageToken: row[COL.MANAGE_TOKEN - 1] || "",
    createdAt: row[COL.CREATED_AT - 1] || "",
    updatedAt: row[COL.UPDATED_AT - 1] || "",
    cancelledAt: row[COL.CANCELLED_AT - 1] || "",
    tokenExpiresAt: row[COL.TOKEN_EXPIRES_AT - 1] || ""
  };
}

// Get blocking state for a specific date, including all-day closure and blocked time slots
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
    const rowDate = toDateStr(values[i][0]);
    const start = toTimeStr(values[i][1]);
    const end = toTimeStr(values[i][2]);
    const active = values[i][4];

    if (!isActive_(active)) continue;
    if (rowDate !== dateStr) continue;

    if (!start && !end) {
      closedAllDay = true;
      break;
    }

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

// Get total pax count by time slot for a specific date, excluding cancelled/no-show reservations
function getPaxByTimeForDate_(dateStr) {
  const reservation = getReservationSheet_();
  const lastRow = reservation.getLastRow();
  if (lastRow < 2) return {};

  const values = reservation.getRange(2, 1, lastRow - 1, COL.TOKEN_EXPIRES_AT).getValues();
  const paxByTime = {};

  for (const r of values) {
    const status = String(r[COL.STATUS - 1] || "").toUpperCase().trim();
    const dateCell = r[COL.DATE - 1];
    const timeCell = r[COL.TIME - 1];
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

    const pax = Number(r[COL.PAX - 1] || 0);
    paxByTime[timeStr] = (paxByTime[timeStr] || 0) + pax;
  }

  return paxByTime;
}

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
  let day = dateObj.getDay();

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

function submitReservation(payload) {
  validateReservationPayload_(payload);

  consumeNewReservationAttempt_(payload);
  assertNotDuplicateReservation_(payload);

  const reservation = getReservationSheet_();
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

    if (isPastReservation_(payload.date, payload.time)) {
      throw new Error("You cannot create a reservation in the past.");
    }

    const paxByTime = getPaxByTimeForDate_(payload.date);
    const currentPax = paxByTime[payload.time] || 0;
    const incomingPax = Number(payload.pax || 0);

    if (currentPax + incomingPax > MAX_PAX_PER_SLOT) {
      throw new Error("This time slot is fully booked (capacity reached). Please choose another time.");
    }

    const reservationId = getReservationID();
    const manageToken = getManageToken_();
    const now = getNow_();
    const tokenExpiresAt = getTokenExpiryForReservation_(payload.date, payload.time);

    reservation.appendRow([
      reservationId,               // A Reservation ID
      "PENDING",                   // B Status
      payload.firstName,           // C First Name
      payload.lastName,            // D Last Name
      payload.email || "",         // E Email
      payload.phone,               // F Phone
      payload.date,                // G Date
      payload.time,                // H Time
      Number(payload.pax || 0),    // I Total Pax
      payload.notes || "",         // J Additional Request
      manageToken,                 // K Manage Token
      now,                         // L Created At
      "",                          // M Updated At
      "",                          // N Cancelled At
      tokenExpiresAt               // O Token Expires At
    ]);

    sendReservationEmail_(reservationId, manageToken, payload);
    sendStaffNotificationEmail_("NEW", reservationId, payload);

    return { ok: true, reservationId };
  } finally {
    lock.releaseLock();
  }
}

function sendReservationEmail_(reservationId, manageToken, payload) {
  if (!payload.email) return;

  const pax = Number(payload.pax || 0);
  const manageLink = buildManageLink_(reservationId, manageToken);
  const buttonHtml = getManageButtonHtml_(reservationId, manageToken);

  const subject = `Chilli Padi Reservation (${reservationId})`;

  const plainBody =
    `Dear ${payload.firstName || "Guest"}${payload.lastName ? " " + payload.lastName : ""},

Thanks for choosing Chilli Padi Nonya Restaurant. We're super excited to have you.

Here are your reservation details:
Reservation ID: ${reservationId}
Date: ${payload.date}
Time: ${payload.time}
Total Pax: ${pax}
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
                  <td style="color:#666;"><b>Total Pax</b></td>
                  <td>${pax}</td>
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

function sendReservationUpdatedEmail_(reservationId, manageToken, payload, oldData) {
  if (!payload.email) return;

  const pax = Number(payload.pax || 0);
  const manageLink = buildManageLink_(reservationId, manageToken);
  const buttonHtml = getManageButtonHtml_(reservationId, manageToken);

  const subject = `Chilli Padi Reservation Updated (${reservationId})`;

  const plainBody =
    `Dear ${payload.firstName || "Guest"}${payload.lastName ? " " + payload.lastName : ""},

Your reservation at Chilli Padi Nonya Restaurant has been successfully updated.

Previous details:
Reservation ID: ${reservationId}
Date: ${oldData.date}
Time: ${oldData.time}
Total Pax: ${oldData.pax}
Additional Request: ${oldData.notes || "None"}

Updated details:
Reservation ID: ${reservationId}
Date: ${payload.date}
Time: ${payload.time}
Total Pax: ${pax}
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
                <tr><td><b>Total Pax</b></td><td>${Number(oldData.pax || 0)}</td></tr>
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
                <tr><td><b>Total Pax</b></td><td>${pax}</td></tr>
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
Total Pax: ${Number(payload.pax || 0)}
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
                <tr><td><b>Total Pax</b></td><td>${Number(payload.pax || 0)}</td></tr>
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
  const pax = Number(payload.pax || 0);

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

Total Pax: ${pax}

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
Total Pax: ${oldData.pax}
Notes: ${oldData.notes || "None"}

Updated
Date: ${payload.date}
Time: ${payload.time}
Total Pax: ${pax}
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

Total Pax: ${pax}

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
  try {
    const resId = String(e?.parameter?.resId || "").trim();
    const token = String(e?.parameter?.token || "").trim();
    const template = HtmlService.createTemplateFromFile("Index");

    template.res = null;
    template.resId = null;
    template.token = null;
    template.errorMessage = "";

    if (resId || token) {
      if (!resId || !token) {
        if (resId) consumeManageAttempt_(resId);
        template.errorMessage = "Invalid reservation link.";
      } else {
        const found = findReservationByIdAndToken_(resId, token);

        if (!found) {
          consumeManageAttempt_(resId);
          template.errorMessage = "This reservation link is invalid or has expired.";
        } else {
          resetManageAttempts_(resId);

          const resData = getReservationRecord_(found.values);
          template.res = resData;
          template.resId = resId;
          template.token = token;
        }
      }
    }

    return template.evaluate()
      .setTitle("Restaurant Reservation")
      .addMetaTag("viewport", "width=device-width, initial-scale=1");
  } catch (err) {
    return HtmlService.createHtmlOutput(`
      <html>
        <head>
          <meta name="viewport" content="width=device-width, initial-scale=1">
          <title>Restaurant Reservation</title>
        </head>
        <body style="font-family:Arial,sans-serif;padding:24px;">
          <h2>Unable to open reservation page</h2>
          <p>Please try again later or contact the restaurant directly.</p>
          <p style="color:#666;font-size:14px;">${escapeHtml_(err.message || err)}</p>
        </body>
      </html>
    `);
  }
}

function updateReservation(form) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);

  try {
    const resId = String(form.resId || "").trim();
    const token = String(form.token || "").trim();

    if (!resId || !token) {
      if (resId) consumeManageAttempt_(resId);
      throw new Error("Invalid reservation request.");
    }

    const found = findReservationByIdAndToken_(resId, token);
    if (!found) {
      consumeManageAttempt_(resId);
      throw new Error("Reservation not found or invalid access token.");
    }

    resetManageAttempts_(resId);

    const row = found.rowNumber;
    const current = getReservationRecord_(found.values);
    const currentStatus = String(current.status || "").toUpperCase().trim();

    if (currentStatus === "CANCELLED") {
      throw new Error("This reservation has already been cancelled.");
    }

    if (isPastReservation_(current.date, current.time)) {
      throw new Error("Past reservations can no longer be modified.");
    }

    const updatedPayload = {
      firstName: current.firstName,
      lastName: current.lastName,
      email: current.email,
      phone: current.phone,
      date: String(form.date || "").trim(),
      time: String(form.time || "").trim(),
      pax: Number(form.pax != null ? form.pax : current.pax),
      notes: form.notes != null ? form.notes : current.notes
    };

    if (!updatedPayload.date || !updatedPayload.time) {
      throw new Error("Date and time are required.");
    }

    if (isPastReservation_(updatedPayload.date, updatedPayload.time)) {
      throw new Error("You cannot move a reservation to a past date/time.");
    }

    const blockState = getBlockStateForDate_(updatedPayload.date);
    if (blockState.closedAllDay) {
      throw new Error("This date is unavailable. Please choose another date.");
    }

    if (blockState.blockedSlots[`${updatedPayload.date}|${updatedPayload.time}`]) {
      throw new Error("This time slot is blocked. Please choose another time.");
    }

    const paxByTime = getPaxByTimeForDate_(updatedPayload.date);

    if (current.date === updatedPayload.date && current.time === updatedPayload.time) {
      paxByTime[updatedPayload.time] =
        Math.max(0, (paxByTime[updatedPayload.time] || 0) - current.pax);
    }

    const newPax = Number(updatedPayload.pax || 0);
    const currentPax = paxByTime[updatedPayload.time] || 0;

    if (newPax <= 0) {
      throw new Error("Please select at least 1 guest.");
    }

    if (newPax > MAX_PAX_PER_RESERVATION) {
      throw new Error("Guest count exceeds allowed maximum, Max 10 Pax.");
    }

    if (currentPax + newPax > MAX_PAX_PER_SLOT) {
      throw new Error("This time slot is fully booked (capacity reached). Please choose another time.");
    }

    const sheet = getReservationSheet_();
    const now = getNow_();
    const newManageToken = getManageToken_();
    const newTokenExpiresAt = getTokenExpiryForReservation_(updatedPayload.date, updatedPayload.time);

    sheet.getRange(row, COL.DATE).setValue(updatedPayload.date);
    sheet.getRange(row, COL.TIME).setValue(updatedPayload.time);
    sheet.getRange(row, COL.PAX).setValue(updatedPayload.pax);
    sheet.getRange(row, COL.NOTES).setValue(updatedPayload.notes);
    sheet.getRange(row, COL.MANAGE_TOKEN).setValue(newManageToken);
    sheet.getRange(row, COL.UPDATED_AT).setValue(now);
    sheet.getRange(row, COL.TOKEN_EXPIRES_AT).setValue(newTokenExpiresAt);

    sendReservationUpdatedEmail_(resId, newManageToken, updatedPayload, {
      date: current.date,
      time: current.time,
      pax: current.pax,
      notes: current.notes
    });

    sendStaffNotificationEmail_("UPDATE", resId, updatedPayload, {
      date: current.date,
      time: current.time,
      pax: current.pax,
      notes: current.notes
    });

    return { ok: true, message: "Reservation updated successfully!" };
  } finally {
    lock.releaseLock();
  }
}

function cancelReservation(resId, token) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);

  try {
    const cleanResId = String(resId || "").trim();
    const cleanToken = String(token || "").trim();

    if (!cleanResId || !cleanToken) {
      if (cleanResId) consumeManageAttempt_(cleanResId);
      throw new Error("Invalid cancellation request.");
    }

    const found = findReservationByIdAndToken_(cleanResId, cleanToken);
    if (!found) {
      consumeManageAttempt_(cleanResId);
      throw new Error("Reservation not found or invalid access token.");
    }

    resetManageAttempts_(cleanResId);

    const row = found.rowNumber;
    const current = getReservationRecord_(found.values);
    const currentStatus = String(current.status || "").toUpperCase().trim();

    if (currentStatus === "CANCELLED") {
      return "This reservation is already cancelled.";
    }

    if (isPastReservation_(current.date, current.time)) {
      throw new Error("Past reservations can no longer be cancelled online.");
    }

    const sheet = getReservationSheet_();
    const now = getNow_();

    sheet.getRange(row, COL.STATUS).setValue("CANCELLED");
    sheet.getRange(row, COL.CANCELLED_AT).setValue(now);
    sheet.getRange(row, COL.UPDATED_AT).setValue(now);
    sheet.getRange(row, COL.MANAGE_TOKEN).setValue("");
    sheet.getRange(row, COL.TOKEN_EXPIRES_AT).setValue("");

    const payload = {
      firstName: current.firstName,
      lastName: current.lastName,
      email: current.email,
      phone: current.phone,
      date: current.date,
      time: current.time,
      pax: current.pax,
      notes: current.notes
    };

    sendReservationCancelledEmail_(cleanResId, payload);
    sendStaffNotificationEmail_("CANCEL", cleanResId, payload);

    return "Success: Your reservation is now cancelled.";
  } finally {
    lock.releaseLock();
  }
}

//================================================================
// Offline reservation creation for walk-in customers or phone reservations
//================================================================

//================================================================
// Spreadsheet UI / staff sidebars
//================================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Reservations")
    .addItem("New Offline Reservation", "openOfflineReservationSidebar")
    .addItem("Edit Reservation", "openEditReservationSidebar")
    .addItem("Block Date/Time", "openBlockingSidebar")
    .addToUi();
}

function openOfflineReservationSidebar() {
  const html = HtmlService
    .createHtmlOutputFromFile("OfflineReservation")
    .setTitle("New Offline Reservation");
  SpreadsheetApp.getUi().showSidebar(html);
}

function openEditReservationSidebar() {
  const html = HtmlService
    .createHtmlOutputFromFile("EditReservation")
    .setTitle("Edit Reservation");
  SpreadsheetApp.getUi().showSidebar(html);
}

function openBlockingSidebar() {
  const html = HtmlService
    .createHtmlOutputFromFile("BlockDateTime")
    .setTitle("Block Date / Time");
  SpreadsheetApp.getUi().showSidebar(html);
}

//================================================================
// Offline reservation creation for walk-in / phone reservations
//================================================================

function createOfflineReservation(form) {
  const payload = {
    firstName: String(form.firstName || "").trim(),
    lastName: String(form.lastName || "").trim(),
    email: String(form.email || "").trim(),
    phone: String(form.phone || "").trim(),
    date: String(form.date || "").trim(),
    time: String(form.time || "").trim(),
    pax: Number(form.pax || 0),
    notes: String(form.notes || "").trim()
  };

  validateReservationPayload_(payload);

  const reservation = getReservationSheet_();
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

    if (isPastReservation_(payload.date, payload.time)) {
      throw new Error("You cannot create a reservation in the past.");
    }

    const paxByTime = getPaxByTimeForDate_(payload.date);
    const currentPax = paxByTime[payload.time] || 0;
    const incomingPax = Number(payload.pax || 0);

    if (currentPax + incomingPax > MAX_PAX_PER_SLOT) {
      throw new Error("This time slot is fully booked (capacity reached). Please choose another time.");
    }

    const reservationId = getReservationID();
    const manageToken = getManageToken_();
    const now = getNow_();
    const tokenExpiresAt = getTokenExpiryForReservation_(payload.date, payload.time);

    reservation.appendRow([
      reservationId,
      "PENDING",
      payload.firstName,
      payload.lastName,
      payload.email || "",
      payload.phone,
      payload.date,
      payload.time,
      Number(payload.pax || 0),
      payload.notes || "",
      manageToken,
      now,
      "",
      "",
      tokenExpiresAt
    ]);

    if (payload.email) {
      sendReservationEmail_(reservationId, manageToken, payload);
    }

    sendStaffNotificationEmail_("NEW", reservationId, payload);

    return {
      ok: true,
      reservationId: reservationId
    };
  } finally {
    lock.releaseLock();
  }
}

function getAvailableTimesForSidebar(dateStr) {
  const cleanDate = String(dateStr || "").trim();
  if (!cleanDate) return [];
  return getAvailableTimes(cleanDate);
}

//================================================================
// Staff edit reservation from spreadsheet sidebar
//================================================================

function findReservationForStaffEdit(searchValue) {
  const keyword = String(searchValue || "").trim();
  if (!keyword) {
    throw new Error("Please enter a Reservation ID or phone number.");
  }

  const sheet = getReservationSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    throw new Error("No reservations found.");
  }

  const values = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const normalizedKeyword = keyword.toLowerCase();
  const normalizedPhoneKeyword = normalizePhoneForLimit_(keyword);

  for (let i = values.length - 1; i >= 0; i--) {
    const row = values[i];

    const rowResId = String(row[COL.RES_ID - 1] || "").trim().toLowerCase();
    const rowPhone = normalizePhoneForLimit_(row[COL.PHONE - 1] || "");

    if (
      rowResId === normalizedKeyword ||
      (normalizedPhoneKeyword && rowPhone === normalizedPhoneKeyword)
    ) {
      return {
        rowNumber: i + 2,
        reservation: {
          id: String(row[COL.RES_ID - 1] || "").trim(),
          status: String(row[COL.STATUS - 1] || "").trim(),
          firstName: String(row[COL.FIRST_NAME - 1] || "").trim(),
          lastName: String(row[COL.LAST_NAME - 1] || "").trim(),
          email: String(row[COL.EMAIL - 1] || "").trim(),
          phone: String(row[COL.PHONE - 1] || "").trim(),
          date: normalizeDateStr_(row[COL.DATE - 1]),
          time: normalizeTimeStr_(row[COL.TIME - 1]),
          pax: Number(row[COL.PAX - 1] || 0),
          notes: String(row[COL.NOTES - 1] || "").trim()
        }
      };
    }
  }

  throw new Error("Reservation not found.");
}

function updateReservationByStaff(form) {
  const rowNumber = Number(form.rowNumber || 0);
  if (!rowNumber) {
    throw new Error("Missing reservation row.");
  }

  const sheet = getReservationSheet_();
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);

  try {
    const currentRow = sheet.getRange(rowNumber, 1, 1, COL.TOKEN_EXPIRES_AT).getValues()[0];
    const current = getReservationRecord_(currentRow);

    if (!current.id) {
      throw new Error("Reservation not found.");
    }

    const currentStatus = String(current.status || "").toUpperCase().trim();
    if (currentStatus === "CANCELLED") {
      throw new Error("Cancelled reservations cannot be edited.");
    }

    const updatedStatus = String(form.status || current.status || "PENDING").toUpperCase().trim();

    // If staff changes to CANCELLED, use the cancel flow
    if (updatedStatus === "CANCELLED") {
      const now = getNow_();

      sheet.getRange(rowNumber, COL.STATUS).setValue("CANCELLED");
      sheet.getRange(rowNumber, COL.CANCELLED_AT).setValue(now);
      sheet.getRange(rowNumber, COL.UPDATED_AT).setValue(now);
      sheet.getRange(rowNumber, COL.MANAGE_TOKEN).setValue("");
      sheet.getRange(rowNumber, COL.TOKEN_EXPIRES_AT).setValue("");

      const payload = {
        firstName: current.firstName,
        lastName: current.lastName,
        email: current.email,
        phone: current.phone,
        date: current.date,
        time: current.time,
        pax: current.pax,
        notes: current.notes
      };

      if (payload.email) {
        sendReservationCancelledEmail_(current.id, payload);
      }

      sendStaffNotificationEmail_("CANCEL", current.id, payload);

      return {
        ok: true,
        reservationId: current.id,
        message: "Reservation cancelled successfully."
      };
    }

    const updatedPayload = {
      firstName: String(form.firstName || "").trim(),
      lastName: String(form.lastName || "").trim(),
      email: String(form.email || "").trim(),
      phone: String(form.phone || "").trim(),
      date: String(form.date || "").trim(),
      time: String(form.time || "").trim(),
      pax: Number(form.pax || 0),
      notes: String(form.notes || "").trim(),
      status: updatedStatus
    };

    validateReservationPayload_(updatedPayload);

    if (isPastReservation_(updatedPayload.date, updatedPayload.time)) {
      throw new Error("You cannot move a reservation to a past date/time.");
    }

    const blockState = getBlockStateForDate_(updatedPayload.date);
    if (blockState.closedAllDay) {
      throw new Error("This date is unavailable. Please choose another date.");
    }

    if (blockState.blockedSlots[`${updatedPayload.date}|${updatedPayload.time}`]) {
      throw new Error("This time slot is blocked. Please choose another time.");
    }

    const paxByTime = getPaxByTimeForDate_(updatedPayload.date);

    if (current.date === updatedPayload.date && current.time === updatedPayload.time) {
      paxByTime[updatedPayload.time] =
        Math.max(0, (paxByTime[updatedPayload.time] || 0) - current.pax);
    }

    const newPax = Number(updatedPayload.pax || 0);
    const currentPax = paxByTime[updatedPayload.time] || 0;

    if (currentPax + newPax > MAX_PAX_PER_SLOT) {
      throw new Error("This time slot is fully booked (capacity reached). Please choose another time.");
    }

    const now = getNow_();
    const newManageToken = getManageToken_();
    const newTokenExpiresAt = getTokenExpiryForReservation_(updatedPayload.date, updatedPayload.time);

    sheet.getRange(rowNumber, COL.STATUS).setValue(updatedPayload.status);
    sheet.getRange(rowNumber, COL.FIRST_NAME).setValue(updatedPayload.firstName);
    sheet.getRange(rowNumber, COL.LAST_NAME).setValue(updatedPayload.lastName);
    sheet.getRange(rowNumber, COL.EMAIL).setValue(updatedPayload.email);
    sheet.getRange(rowNumber, COL.PHONE).setValue(updatedPayload.phone);
    sheet.getRange(rowNumber, COL.DATE).setValue(updatedPayload.date);
    sheet.getRange(rowNumber, COL.TIME).setValue(updatedPayload.time);
    sheet.getRange(rowNumber, COL.PAX).setValue(updatedPayload.pax);
    sheet.getRange(rowNumber, COL.NOTES).setValue(updatedPayload.notes);
    sheet.getRange(rowNumber, COL.UPDATED_AT).setValue(now);

    // DEPARTED / NO-SHOW = update only, no customer/staff emails
    if (updatedPayload.status === "DEPARTED" || updatedPayload.status === "NO-SHOW") {
      if (updatedPayload.status === "NO-SHOW") {
        sheet.getRange(rowNumber, COL.MANAGE_TOKEN).setValue("");
        sheet.getRange(rowNumber, COL.TOKEN_EXPIRES_AT).setValue("");
      } else {
        sheet.getRange(rowNumber, COL.MANAGE_TOKEN).setValue(newManageToken);
        sheet.getRange(rowNumber, COL.TOKEN_EXPIRES_AT).setValue(newTokenExpiresAt);
      }

      return {
        ok: true,
        reservationId: current.id,
        message: "Reservation updated successfully."
      };
    }

    // Normal edit flow
    sheet.getRange(rowNumber, COL.MANAGE_TOKEN).setValue(newManageToken);
    sheet.getRange(rowNumber, COL.TOKEN_EXPIRES_AT).setValue(newTokenExpiresAt);

    const oldData = {
      date: current.date,
      time: current.time,
      pax: current.pax,
      notes: current.notes
    };

    if (updatedPayload.email) {
      sendReservationUpdatedEmail_(current.id, newManageToken, updatedPayload, oldData);
    }

    sendStaffNotificationEmail_("UPDATE", current.id, updatedPayload, oldData);

    return {
      ok: true,
      reservationId: current.id,
      message: "Reservation updated successfully."
    };
  } finally {
    lock.releaseLock();
  }
}

function cancelReservationByStaff(rowNumber) {
  const row = Number(rowNumber || 0);
  if (!row) throw new Error("Invalid reservation row.");

  const sheet = getReservationSheet_();
  const values = sheet.getRange(row, 1, 1, COL.TOKEN_EXPIRES_AT).getValues()[0];
  const record = getReservationRecord_(values);

  if (!record.id) {
    throw new Error("Reservation not found.");
  }

  if (String(record.status || "").toUpperCase() === "CANCELLED") {
    throw new Error("Reservation already cancelled.");
  }

  const now = getNow_();

  sheet.getRange(row, COL.STATUS).setValue("CANCELLED");
  sheet.getRange(row, COL.CANCELLED_AT).setValue(now);
  sheet.getRange(row, COL.UPDATED_AT).setValue(now);
  sheet.getRange(row, COL.MANAGE_TOKEN).setValue("");
  sheet.getRange(row, COL.TOKEN_EXPIRES_AT).setValue("");

  const payload = {
    firstName: record.firstName,
    lastName: record.lastName,
    email: record.email,
    phone: record.phone,
    date: record.date,
    time: record.time,
    pax: record.pax,
    notes: record.notes
  };

  if (payload.email) {
    sendReservationCancelledEmail_(record.id, payload);
  }

  sendStaffNotificationEmail_("CANCEL", record.id, payload);

  return {
    ok: true,
    message: "Reservation cancelled successfully."
  };
}

//================================================================
// Blocking sidebar
//================================================================

function createBlockingEntry(form) {
  const blockingSheet = getSheets_().blocking;
  if (!blockingSheet) {
    throw new Error("Blocking sheet not found.");
  }

  const date = String(form.date || "").trim();
  const blockType = String(form.blockType || "").trim();
  const startTime = String(form.startTime || "").trim();
  const endTime = String(form.endTime || "").trim();
  const reason = String(form.reason || "").trim() || "Blocked";

  if (!date) {
    throw new Error("Date is required.");
  }

  let finalStart = "";
  let finalEnd = "";

  if (blockType === "TIME_RANGE") {
    if (!startTime || !endTime) {
      throw new Error("Start time and end time are required.");
    }
    if (startTime >= endTime) {
      throw new Error("End time must be later than start time.");
    }
    finalStart = startTime;
    finalEnd = endTime;
  }

  blockingSheet.appendRow([
    date,        // A Date
    finalStart,  // B Start
    finalEnd,    // C End
    reason,      // D Reason
    "Active"     // E Active
  ]);

  return {
    ok: true,
    message: blockType === "FULL_DAY"
      ? "Full-day block added successfully."
      : "Time block added successfully."
  };
}

function getBlockingEntriesByDate(dateStr) {
  const cleanDate = String(dateStr || "").trim();
  if (!cleanDate) return [];

  const blockingSheet = getSheets_().blocking;
  if (!blockingSheet) return [];

  const lastRow = blockingSheet.getLastRow();
  if (lastRow < 2) return [];

  const values = blockingSheet.getRange(2, 1, lastRow - 1, 5).getValues();
  const tz = Session.getScriptTimeZone();

  return values
    .map(function(row, index) {
      const rowDate = row[0] instanceof Date
        ? Utilities.formatDate(row[0], tz, "yyyy-MM-dd")
        : String(row[0] || "").trim();

      const start = row[1] instanceof Date
        ? Utilities.formatDate(row[1], tz, "HH:mm")
        : String(row[1] || "").trim();

      const end = row[2] instanceof Date
        ? Utilities.formatDate(row[2], tz, "HH:mm")
        : String(row[2] || "").trim();

      return {
        rowNumber: index + 2,
        date: rowDate,
        startTime: start,
        endTime: end,
        reason: String(row[3] || "").trim(),
        active: String(row[4] || "").trim()
      };
    })
    .filter(function(item) {
      return item.date === cleanDate && String(item.active).toLowerCase() === "active";
    });
}

function deactivateBlockingEntry(rowNumber) {
  const row = Number(rowNumber || 0);
  if (!row) {
    throw new Error("Invalid blocking row.");
  }

  const blockingSheet = getSheets_().blocking;
  if (!blockingSheet) {
    throw new Error("Blocking sheet not found.");
  }

  blockingSheet.getRange(row, 5).setValue("Inactive");

  return {
    ok: true,
    message: "Blocking entry deactivated."
  };
}

// Debug
function debugFindReservation() {
  const result = debugFindReservationInternal_("F8MK7B2");
  Logger.log(result);
}

function debugFindReservationByValue(value) {
  const result = debugFindReservationInternal_(value);
  Logger.log(JSON.stringify(result));
}

function debugTestF8MK7B2() {
  debugFindReservationByValue("F8MK7B2");
}

function debugTest84729273() {
  debugFindReservationByValue("84729273");
}

function debugFindReservationInternal_(searchValue) {
  const keyword = String(searchValue || "").trim();
  const sheet = getReservationSheet_();
  const lastRow = sheet.getLastRow();

  const result = {
    searchValue: keyword,
    reservationSheetName: sheet.getName(),
    lastRow: lastRow,
    matches: []
  };

  if (lastRow < 2) return result;

  const values = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const normalizedKeyword = keyword.toLowerCase();
  const normalizedPhoneKeyword = normalizePhoneForLimit_(keyword);

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const resId = String(row[COL.RES_ID - 1] || "").trim();
    const phone = String(row[COL.PHONE - 1] || "").trim();

    const rowResId = resId.toLowerCase();
    const rowPhone = normalizePhoneForLimit_(phone);

    if (
      rowResId === normalizedKeyword ||
      (normalizedPhoneKeyword && rowPhone === normalizedPhoneKeyword)
    ) {
      result.matches.push({
        rowNumber: i + 2,
        reservationId: resId,
        phone: phone,
        status: String(row[COL.STATUS - 1] || "").trim(),
        date: String(row[COL.DATE - 1] || "").trim(),
        time: String(row[COL.TIME - 1] || "").trim(),
        pax: Number(row[COL.PAX - 1] || 0)
      });
    }
  }

  return result;
}