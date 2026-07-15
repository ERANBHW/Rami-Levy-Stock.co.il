// ============================================================
//  Google Apps Script — מערכת קריאות שירות
//  הוראות התקנה:
//  1. צור Google Sheet חדש, העתק את ה-ID מהכתובת (ה-/d/.../edit)
//  2. כנס ל-Extensions > Apps Script, מחק הכל, הדבק קובץ זה
//  3. שנה את SHEET_ID ו-ADMIN_TOKEN לערכים שלך
//  4. הרץ initSheet() פעם אחת → מגדיר כותרות
//  5. הרץ setupTrigger() פעם אחת → מגדיר תזכורות יומיות
//  6. Deploy > New Deployment > Web App
//     Execute as: Me | Who has access: Anyone, even anonymous
//  7. העתק את ה-URL לתוך APPS_SCRIPT_URL ב-support.html
// ============================================================

var SHEET_ID      = 'YOUR_GOOGLE_SHEET_ID_HERE';
var SHEET_NAME    = 'קריאות שירות';
var SUPPORT_EMAIL = 'support@rami-levy-stock.co.il';
var MANAGER_EMAIL = 'eran@rami-levy-stock.co.il';
var ADMIN_TOKEN   = 'REPLACE_WITH_YOUR_SECRET_TOKEN';  // כל מחרוזת סודית, למשל: RL2024admin

// ── helpers ──────────────────────────────────────────────────
function getSheet() {
  return SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
}

function fmtDate(d) {
  return Utilities.formatDate(d || new Date(), 'Asia/Jerusalem', 'dd/MM/yyyy HH:mm');
}

// ── doPost: פתיחת קריאה חדשה או עדכון ────────────────────────────────
function doPost(e) {
  try {
    var data  = JSON.parse(e.postData.contents);
    var sheet = getSheet();
    var now   = new Date();

    // עדכון קריאה קיימת
    if (data.action === 'update') {
      var rows = sheet.getDataRange().getValues();
      for (var i = 1; i < rows.length; i++) {
        if (rows[i][0] === data.ticketNumber) {
          sheet.getRange(i + 1, 16).setValue(data.notes);
          sheet.getRange(i + 1, 15).setValue(fmtDate(now));

          // מייל הודעה על עדכון
          try {
            MailApp.sendEmail(
              data.email,
              'קריאה ' + data.ticketNumber + ' — עודכנה',
              'קריאתך עודכנה עם הערות חדשות:\n\n' + data.notes + '\n\nנחזור אליך בהקדם.'
            );
          } catch(me) { Logger.log('Email error: ' + me.message); }

          return ContentService
            .createTextOutput(JSON.stringify({ success: true, message: 'updated' }))
            .setMimeType(ContentService.MimeType.JSON);
        }
      }
      return ContentService
        .createTextOutput(JSON.stringify({ success: false, error: 'ticket not found' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // מספר קריאה: TK-NNNN (ספרור פשוט)
    var rowNum      = sheet.getLastRow();   // row 1 = header
    var ticketNum   = 'TK-' + String(rowNum).padStart(4, '0');

    sheet.appendRow([
      ticketNum,                // A: מספר קריאה
      fmtDate(now),             // B: תאריך פתיחה
      'פתוח',                   // C: סטטוס
      data.name        || '',   // D: שם עובד
      data.email       || '',   // E: מייל
      data.phone       || '',   // F: טלפון
      data.branch      || '',   // G: סניף
      data.computerName|| '',   // H: שם מחשב
      data.ip          || '',   // I: IP
      data.printer     || '',   // J: מדפסת
      data.anyDesk     || '',   // K: AnyDesk
      data.category    || '',   // L: קטגוריה
      data.urgency     || '',   // M: דחיפות
      data.description || '',   // N: תיאור
      fmtDate(now),             // O: תאריך עדכון
      ''                        // P: הערות
    ]);

    // מייל אישור לעובד
    try {
      MailApp.sendEmail(
        data.email,
        'קריאת שירות ' + ticketNum + ' — אישור קבלה',
        'שלום ' + data.name + ',\n\n' +
        'קריאת השירות שלך נפתחה בהצלחה.\n\n' +
        'מספר קריאה: ' + ticketNum + '\n' +
        'תאריך: '      + fmtDate(now) + '\n' +
        'קטגוריה: '    + data.category + '\n' +
        'דחיפות: '     + data.urgency  + '\n\n' +
        'תיאור:\n'     + data.description + '\n\n' +
        'נחזור אליך בהקדם האפשרי.\n' +
        'בברכה, צוות המחשוב'
      );
    } catch(e) {
      Logger.log('Error sending employee email: ' + e.message);
    }

    // התראה לצוות תמיכה
    try {
      MailApp.sendEmail(
        SUPPORT_EMAIL,
        '[' + data.urgency + '] קריאה חדשה: ' + ticketNum + ' | ' + data.category + ' | ' + data.name,
        'קריאה חדשה נפתחה\n\n' +
        'מספר קריאה: ' + ticketNum      + '\n' +
        'תאריך: '      + fmtDate(now)   + '\n\n' +
        'עובד: '        + data.name     + '\n' +
        'טלפון: '       + data.phone    + '\n' +
        'מייל: '        + data.email    + '\n' +
        'סניף: '        + data.branch   + '\n' +
        'מחשב: '        + data.computerName + '\n' +
        'IP: '          + data.ip       + '\n' +
        'מדפסת: '       + data.printer  + '\n' +
        'AnyDesk: '     + data.anyDesk  + '\n\n' +
        'קטגוריה: '     + data.category + '\n' +
        'דחיפות: '      + data.urgency  + '\n\n' +
        'תיאור:\n'      + data.description,
        { cc: MANAGER_EMAIL }
      );
    } catch(e) {
      Logger.log('Error sending support email: ' + e.message);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ success: true, ticketNumber: ticketNum }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── doGet: קריאת קריאות לתצוגה ──────────────────────────────
function doGet(e) {
  try {
    var action  = e.parameter.action || '';
    var email   = e.parameter.email  || '';
    var token   = e.parameter.token  || '';
    var isAdmin = (token === ADMIN_TOKEN);
    var sheet   = getSheet();
    var rows    = sheet.getDataRange().getValues();

    if (rows.length <= 1) {
      if (action === 'stats') {
        return ContentService
          .createTextOutput(JSON.stringify({ open: 0, progress: 0, closed: 0 }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      return ContentService.createTextOutput('[]').setMimeType(ContentService.MimeType.JSON);
    }

    // סטטיסטיקות כוללות
    if (action === 'stats') {
      var stats = { open: 0, progress: 0, closed: 0 };
      rows.slice(1).forEach(function(r) {
        var status = String(r[2] || '').trim().toLowerCase();
        if (status === 'פתוח') stats.open++;
        else if (status === 'בטיפול') stats.progress++;
        else if (status === 'סגור') stats.closed++;
      });
      return ContentService
        .createTextOutput(JSON.stringify(stats))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var tickets = rows.slice(1)
      .filter(function(r) {
        return isAdmin || (email && r[4] === email);
      })
      .map(function(r) {
        return {
          ticketNumber: r[0],
          date:         r[1],
          status:       r[2],
          name:         r[3],
          email:        r[4],
          branch:       r[6],
          computerName: r[7],
          category:     r[11],
          urgency:      r[12],
          description:  r[13],
          updated:      r[14],
          notes:        r[15]
        };
      })
      .reverse();

    return ContentService
      .createTextOutput(JSON.stringify(tickets))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── sendReminders: תזכורות יומיות ───────────────────────────
function sendReminders() {
  var sheet = getSheet();
  var rows  = sheet.getDataRange().getValues();
  if (rows.length <= 1) return;

  var cutoff = new Date(Date.now() - 24 * 60 * 60 * 1000);

  var open = rows.slice(1).filter(function(r) {
    if (r[2] !== 'פתוח') return false;
    try {
      var parts = String(r[1]).split(' ');
      var d = parts[0].split('/'); var t = (parts[1] || '00:00').split(':');
      return new Date(d[2], d[1]-1, d[0], t[0], t[1]) < cutoff;
    } catch(e) { return false; }
  });

  if (!open.length) return;

  var body = 'תזכורת: ' + open.length + ' קריאות פתוחות מעל 24 שעות\n\n';
  open.forEach(function(r) {
    body += r[0] + ' | ' + r[1] + ' | ' + r[12] + ' | ' + r[3] + ' | ' + r[11] + '\n';
  });

  MailApp.sendEmail({
    to:      SUPPORT_EMAIL,
    cc:      MANAGER_EMAIL,
    subject: 'תזכורת: ' + open.length + ' קריאות פתוחות מעל 24 שעות',
    body:    body
  });
}

// ── ONE-TIME SETUP FUNCTIONS ─────────────────────────────────

// הרץ פעם אחת: מגדיר את הגיליון
function initSheet() {
  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);

  var headers = [
    'מספר קריאה','תאריך פתיחה','סטטוס','שם עובד','מייל עובד',
    'טלפון','סניף','שם מחשב','IP','מדפסת','AnyDesk',
    'קטגוריה','דחיפות','תיאור','תאריך עדכון','הערות'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 160);
  sheet.setColumnWidth(14, 300);

  // עיצוב עמודת סטטוס
  var statusRange = sheet.getRange('C2:C1000');
  var rule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('פתוח').setBackground('#fff3cd').setFontColor('#856404').setRanges([statusRange]).build();
  sheet.setConditionalFormatRules([rule]);

  SpreadsheetApp.getUi().alert('הגיליון הוגדר בהצלחה!');
}

// הרץ פעם אחת: מגדיר טריגר יומי ב-09:00
function setupTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'sendReminders') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sendReminders')
    .timeBased().atHour(9).everyDays(1).create();
  SpreadsheetApp.getUi().alert('טריגר יומי הוגדר בהצלחה (09:00 בכל יום)');
}
