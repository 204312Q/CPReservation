const RESERVATION_SHEET_NAME = "Joo Chiat Reservation";
const BLOCKING_SHEET_NAME = "Joo Chiat_Blocking";
const PUBLIC_HOLIDAY_SHEET_NAME = "Public Holiday";

const MAX_PAX_PER_SLOT = 60;


function doGet() {
  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("Restaurant Reservation");
}

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

    return { ok: true, reservationId };

  } finally {
    // 🔓 Always release lock
    lock.releaseLock();
  }
}


