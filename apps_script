function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('自習室預約系統')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 1. Get all reservations from Sheet
function getReservations() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Reservations");
  const data = sheet.getDataRange().getValues();
  const headers = data.shift(); // Remove headers
  
  const reservations = data.map((row, index) => ({
    rowId: index + 2, // Row index in Sheet
    id: row[0],
    seat_number: row[1],
    period: row[2],
    nickname: row[3],
    status: row[4] || 'active'
  }));

  // Sort: Active first, then Done
  return reservations.sort((a, b) => {
    if (a.status === 'active' && b.status === 'done') return -1;
    if (a.status === 'done' && b.status === 'active') return 1;
    return 0;
  });
}

// 2. Add a new reservation (with Anti-duplication)
function addReservation(seatNum, period, nickname) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Reservations");
  const data = sheet.getDataRange().getValues();

  // Check if seat is already taken for that period
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] == seatNum && data[i][2] == period) {
      throw new Error("該時段此座位已被預約！");
    }
  }

  const id = new Date().getTime(); // Unique ID
  sheet.appendRow([id, seatNum, period, nickname, 'active']);
  return { success: true };
}

// 3. Toggle Status (Active <-> Done)
function toggleStatus(rowId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Reservations");
  const currentStatus = sheet.getRange(rowId, 5).getValue();
  const newStatus = currentStatus === 'done' ? 'active' : 'done';
  sheet.getRange(rowId, 5).setValue(newStatus);
  return newStatus;
}