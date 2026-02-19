const RESERVATION_SHEET_NAME = "Joo Chiat Reservation";
const BLOCKING_SHEET_NAME = "Joo Chiat_Blocking";
const PUBLIC_HOLIDAY_SHEET_NAME = "Public Holiday";

const MAX_PAX_PER_SLOT = 60;


// function doGet() {
//   return HtmlService.createHtmlOutputFromFile("Index")
//     .setTitle("Restaurant Reservation");
// }

function getSheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return {
    reservation: ss.getSheetByName(RESERVATION_SHEET_NAME),
    blocking: ss.getSheetByName(BLOCKING_SHEET_NAME),
  };
}

//Generate Reservation ID
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

function getBlockStateForDate_(dateStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
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
    // handle "12:00:00" => "12:00"
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

    // ✅ Whole-day block when Start/End are blank
    if (!start && !end) {
      closedAllDay = true;
      break;
    }

    // ✅ Time-range block when both Start and End exist
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

    // Count only active reservations (adjust if needed)
    if (status === "CANCELLED" || status === "NO-SHOW") continue;

    const rowDateStr = (dateCell instanceof Date)
      ? Utilities.formatDate(dateCell, Session.getScriptTimeZone(), "yyyy-MM-dd")
      : String(dateCell).trim();

    if (rowDateStr !== dateStr) continue;

    // ✅ Normalize time to "HH:mm"
    let timeStr;
    if (timeCell instanceof Date) {
      timeStr = Utilities.formatDate(timeCell, Session.getScriptTimeZone(), "HH:mm");
    } else {
      // handle "19:15", "19:15:00", etc.
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



// Block table expected columns: Date | Time | Reason (you can have more columns, we read first 2)
function getBlockedMap() {
  const { blocking } = getSheets_();
  const values = blocking.getDataRange().getValues();
  const map = {};

  for (let i = 1; i < values.length; i++) {
    const date = values[i][0]; // col A
    const time = values[i][1]; // col B
    if (!date || !time) continue;

    // normalize to YYYY-MM-DD if it's a Date object
    const dateStr = (date instanceof Date)
      ? Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd")
      : String(date).trim();

    const timeStr = String(time).trim();
    map[`${dateStr}|${timeStr}`] = true;
  }
  return map;
}

//Fetch all the public holidays from Google Sheet
function getPublicHolidayMap_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
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
    day = 0; // treat as Sunday
  }

  const isWeekend = (day === 0 || day === 6);

  const toMinutes = (t) => {
    const [h, m] = t.split(":").map(Number);
    return h * 60 + m;
  };

  // Lunch is same for all days
  const lunch = { start: "11:30", end: "14:15" };

  // Dinner last reservable time depends on weekday/weekend
  const dinner = isWeekend
    ? { start: "17:30", end: "21:30" }  // weekend last order
    : { start: "17:30", end: "21:00" }; // weekday last order

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

      // 1) hide if backend blocked
      if (blockedMap[`${dateStr}|${timeStr}`]) continue;

      // 2) hide if fully booked (Adults + Children >= 60)
      const currentPax = paxByTime[timeStr] || 0;
      if (currentPax >= MAX_PAX_PER_SLOT) continue;

      results.push(timeStr);
    }
  }

  return results;
}

//Submit Reservation 
function submitReservation(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const reservation = ss.getSheetByName("Joo Chiat Reservation");

  // Basic validation
  if (!payload.firstName || !payload.lastName || !payload.phone || !payload.date || !payload.time) {
    throw new Error("Missing required fields.");
  }

  // 🔐 Lock to prevent double booking
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);

  try {
    // Re-check blocked slots (backend safety)
    const blockedMap = getBlockedMap();
    if (blockedMap[`${payload.date}|${payload.time}`]) {
      throw new Error("This time slot is blocked. Please choose another time.");
    }

    // Capacity check
    const paxByTime = getPaxByTimeForDate_(payload.date);
    const currentPax = paxByTime[payload.time] || 0;
    const incomingPax =
      Number(payload.adults || 0) + Number(payload.children || 0);

    if (currentPax + incomingPax > MAX_PAX_PER_SLOT) {
      throw new Error(
        "This time slot is fully booked (capacity reached). Please choose another time."
      );
    }

    // Generate Reservation ID
    const reservationId = getReservationID();

    // ✅ Safe to append row
    reservation.appendRow([
      reservationId,                 // A Reservation ID
      "PENDING",                      // B Status
      payload.firstName,              // C First Name
      payload.lastName,               // D Last Name
      payload.email || "",            // E Email
      payload.phone,                  // F Phone
      payload.date,                   // G Date
      payload.time,                   // H Time
      Number(payload.adults || 0),    // I Adults
      Number(payload.children || 0),  // J Children
      payload.notes || ""             // K Additional Request
    ]);

    sendReservationEmail_(reservationId, payload);

    return { ok: true, reservationId };

  } finally {
    // 🔓 Always release lock
    lock.releaseLock();
  }
}

function sendReservationEmail_(reservationId, payload) {
  const adults = Number(payload.adults || 0);
  const children = Number(payload.children || 0);

const webAppUrl = ScriptApp.getService().getUrl(); 
  const manageLink = `${webAppUrl}?resId=${reservationId}`;

  const buttonHtml = `<a href="${manageLink}" style="background:#8B0000; color:white; padding:10px; text-decoration:none;">Manage Booking</a>`;


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

Need to change or cancel your reservation? Please contact us and include your Reservation ID.

Chilli Padi`;

  const htmlBody = `
<div style="margin:0;padding:0;background:#f4f6f8;font-family:Arial,Helvetica,sans-serif;">
  <table width="100%" cellpadding="0" cellspacing="0" style="background:#f4f6f8;padding:30px 10px;">
    <tr>
      <td align="center">
        
        <!-- Card Container -->
        <table width="100%" max-width="600" cellpadding="0" cellspacing="0"
          style="max-width:600px;background:#ffffff;border-radius:12px;
                 box-shadow:0 8px 24px rgba(0,0,0,0.06);overflow:hidden;">

          <!-- Header -->
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

          <!-- Body -->
          <tr>
            <td style="padding:24px;color:#333333;font-size:14px;line-height:1.6;">
              
              <p style="margin-top:0;">
                Dear ${escapeHtml_(payload.firstName || "Guest")}${payload.lastName ? " " + escapeHtml_(payload.lastName) : ""},
              </p>

              <p>
                Thank you for choosing <b>Chilli Padi Nonya Restaurant</b>. 
                We’re excited to welcome you!
              </p>

              <!-- Reservation Details Box -->
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
                please click on the button below.
                </br>
                ${buttonHtml}
              </p>

              <p style="margin-bottom:0;">
                We look forward to serving you
              </p>

            </td>
          </tr>

          <!-- Footer -->
          <tr>
            <td style="background:#fafafa;padding:16px;text-align:center;
                       font-size:12px;color:#888;">
              © ${new Date().getFullYear()} Chilli Padi Nonya Restaurant<br>
              This is an automated confirmation email. Please do not reply.
            </td>
          </tr>

        </table>
        <!-- End Card -->

      </td>
    </tr>
  </table>
</div>
`;


  MailApp.sendEmail({
    to: payload.email,
    subject,
    body: plainBody,
    htmlBody,
    name: "No Reply - Chilli Padi Nonya Restaurant"
  });
}

function escapeHtml_(text) {
  return String(text ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function doGet(e) {
  // 1. Check if the URL has a Reservation ID
  const resId = e && e.parameter ? e.parameter.resId : null;

  if (resId) {
    try {
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RESERVATION_SHEET_NAME);
      const data = sheet.getDataRange().getValues();
      let resData = null;

      // Find the row
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] && data[i][0].toString() === resId.toString()) {
          resData = {
            id: data[i][0],
            firstName: data[i][2],
            // Format date specifically so the HTML <input type="date"> can read it
            date: data[i][6] instanceof Date ? Utilities.formatDate(data[i][6], Session.getScriptTimeZone(), "yyyy-MM-dd") : data[i][6],
            adults: data[i][8],
            notes: data[i][10]
          };
          break;
        }
      }

      if (!resData) {
        return HtmlService.createHtmlOutput("<h3>Reservation Not Found</h3><p>Could not find ID: " + resId + "</p>");
      }

      // 2. Try to load AmendPage
      const template = HtmlService.createTemplateFromFile('AmendPage');
      template.res = resData;
      return template.evaluate()
        .setTitle("Manage Reservation | Chilli Padi")
        .addMetaTag('viewport', 'width=device-width, initial-scale=1')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

    } catch (err) {
      // If AmendPage.html has a syntax error, this will show it!
      return HtmlService.createHtmlOutput("<h3>Template Error</h3><p>" + err.message + "</p>");
    }
  }

  // 3. If no ID, show the standard booking form
  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("Restaurant Reservation")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}