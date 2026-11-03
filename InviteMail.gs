/**
 * שליחת ההזמנה לכנס "עיתונות בעידן ה־AI" לרשימת תפוצה, מייל אישי לכל נמען.
 *
 * הכנה (פעם אחת):
 * 1. העלו לגיטהאב (לדף הראשי של הריפו) את התמונה invite-header.png.
 *    בדקו שהיא נפתחת בכתובת: https://aijournalism.co.il/invite-header.png
 * 2. (רשות) העלו את קובץ ה-PDF של ההזמנה ל-Google Drive של החשבון השולח,
 *    ופתחו אותו. מזהה הקובץ הוא החלק בכתובת שבין /d/ לבין /view.
 *    הדביקו אותו ב-PDF_FILE_ID למטה. אם משאירים ריק, נשלח בלי קובץ מצורף.
 * 3. בגיליון (בחשבון השולח) צרו לשונית בשם: הפצה
 *    שורה 1 = כותרות:  שם | אימייל | נשלח
 *    "שם" לא חובה: אם יש שם, המייל ייפתח ב"שלום <שם>,". אם אין, ב"שלום,".
 * 4. Extensions > Apps Script > קובץ חדש (+ > Script), הדביקו את הקובץ הזה ושמרו.
 * 5. הריצו sendTest (שולח רק אליכם). אם הכול נראה טוב, הריצו sendInvitations.
 *    בהרצה הראשונה Google תבקש הרשאה: Advanced > Go to... > Select all > Continue.
 *
 * מגבלה: חשבון Workspace שולח עד 1,500 נמענים ביום (Gmail פרטי: כ-100).
 * הסקריפט מסמן "נשלח" ליד כל שורה, ובהרצה חוזרת ממשיך רק עם מי שעוד לא קיבל.
 */

const INVITE_SHEET   = "הפצה";
const INVITE_SUBJECT = "הזמנה: עיתונות בעידן ה־AI | 2.12.2026, אוניברסיטת תל אביב";
const INVITE_FROM    = "";                    // למשל "info@aijournalism.co.il", רק אחרי שהוגדר ב-Send mail as
const INVITE_NAME    = "AI Journalism 2026";  // שם השולח שהנמענים רואים
const PDF_FILE_ID    = "";                    // מזהה ה-PDF ב-Drive (רשות)
const SITE           = "https://aijournalism.co.il";
const HEADER_IMG     = SITE + "/invite-header.png";

const SPEAKERS = [
  "פרופ' תמר אשורי", "פרופ' ליאור זלמנסון", "דובב גולדשטיין", "ד\"ר אבנר קנטור", "אור הירשאוגה",
  "איתי זילבר", "אלוף בן", "אמיר בן-דוד", "ד\"ר אסף הורוביץ", "בירנית גורן", "גיא זהר", "דין ליבנה",
  "יהונתן קלינגר", "פרופ' ניבה אלקין-קורן", "עומר בן יעקב", "ענת סרגוסטי", "ד\"ר עידו באום",
  "עמרי מרכוס", "פרופ' ענת בן-דוד", "פרופ' צבי רייך", "פרופ' רועי דודזון", "שוקי טאוסיג",
  "ד\"ר תהילה שוורץ אלטשולר",
];

// ---------- הרצה ----------
function sendTest() {
  const me = Session.getActiveUser().getEmail();
  sendOne_(me, "נדב");
  SpreadsheetApp.getUi().alert("נשלחה הזמנת בדיקה אל " + me);
}

function sendInvitations() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(INVITE_SHEET);
  if (!sheet) throw new Error('לא נמצאה לשונית בשם "' + INVITE_SHEET + '"');
  const data = sheet.getDataRange().getValues();
  const head = data[0].map(h => String(h).trim());
  const cName = head.indexOf("שם"), cMail = head.indexOf("אימייל"), cSent = head.indexOf("נשלח");
  if (cMail < 0 || cSent < 0) throw new Error("חסרות כותרות: אימייל / נשלח");

  let quota = MailApp.getRemainingDailyQuota(), sent = 0, left = 0;
  for (let r = 1; r < data.length; r++) {
    const email = String(data[r][cMail]).trim();
    if (!email || data[r][cSent]) continue;
    if (quota <= 0) { left++; continue; }
    try {
      sendOne_(email, cName >= 0 ? String(data[r][cName]).trim() : "");
      sheet.getRange(r + 1, cSent + 1).setValue(new Date());
      sent++; quota--;
    } catch (e) {
      sheet.getRange(r + 1, cSent + 1).setValue("שגיאה: " + e.message);
    }
  }
  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert("נשלחו " + sent + " הזמנות." +
    (left ? "\nנותרו " + left + ": הגעתם למכסה היומית, הריצו שוב מחר." : "\nכל ההזמנות נשלחו."));
}

// ---------- שליחה ----------
function sendOne_(to, name) {
  const opts = { htmlBody: inviteHtml_(name), name: INVITE_NAME };
  if (INVITE_FROM) opts.from = INVITE_FROM;
  if (PDF_FILE_ID) opts.attachments = [DriveApp.getFileById(PDF_FILE_ID).getBlob().setName("הזמנה - עיתונות בעידן ה-AI.pdf")];
  GmailApp.sendEmail(to, INVITE_SUBJECT, inviteText_(name), opts);
}

function calendarUrl_() {
  return "https://calendar.google.com/calendar/render?action=TEMPLATE" +
    "&text=" + encodeURIComponent("עיתונות בעידן ה־AI: מעצבים את כללי המשחק") +
    "&dates=20261202T130000Z/20261202T170000Z" +
    "&location=" + encodeURIComponent("אוניברסיטת תל אביב, בניין נפתלי, אולם אפטר") +
    "&details=" + encodeURIComponent(SITE);
}

// גרסת טקסט פשוט (לתוכנות מייל שלא מציגות עיצוב, ולקוראי מסך)
function inviteText_(name) {
  return (name ? "שלום " + name + "," : "שלום,") + "\n\n" +
    "החוג לתקשורת ע\"ש דן והפקולטה לניהול ע\"ש קולר באוניברסיטת תל אביב מתכבדים להזמינכם לכנס\n" +
    "\"עיתונות בעידן ה־AI: מעצבים את כללי המשחק\".\n\n" +
    "יחד עם בכירי העיתונות, אנשי טכנולוגיה וחוקרי תקשורת נדון בשאלה כיצד להטמיע כלי בינה מלאכותית במערכות החדשות בישראל: איך לממש את יתרונותיהם ואיך לצמצם את הסיכונים.\n\n" +
    "מתי: יום רביעי, 2.12.2026, בין 15:00 ל-19:00\n" +
    "איפה: אוניברסיטת תל אביב, בניין נפתלי, אולם אפטר\n\n" +
    "בין הדוברים והדוברות: " + SPEAKERS.join(", ") + ".\n\n" +
    "ההשתתפות בהרשמה מראש באתר הכנס: " + SITE + "\n\n" +
    "נשמח לראותכם,\nהחוג לתקשורת ע\"ש דן והפקולטה לניהול ע\"ש קולר, אוניברסיטת תל אביב";
}

function esc_(s) {
  return String(s).replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
}

// גרסת HTML מעוצבת, באותה שפה עיצובית כמו ההזמנה
function inviteHtml_(name) {
  const NAVY = "#16245c", SOFT = "#cdd5f0", BG = "#eef0f5", INK = "#0a0c12", MUTED = "#4f5869", LINE = "#d3d8e3";
  const FONT = "'IBM Plex Sans Hebrew',Assistant,'Segoe UI',Arial,sans-serif";
  const MONO = "'IBM Plex Mono',Menlo,Consolas,'Courier New',monospace";
  const label = t => '<div lang="en" style="font-family:' + MONO + ';font-size:11px;letter-spacing:.14em;color:' + SOFT + ';direction:ltr">' + t + '</div>';
  const fact = (lab, main, small, last) =>
    '<td width="33%" valign="top" align="center" style="padding:18px 8px;' + (last ? '' : 'border-left:1px solid rgba(255,255,255,.28);') + '">' +
      label(lab) +
      '<div style="font-size:18px;font-weight:700;color:#ffffff;line-height:1.3;padding-top:4px">' + main + '</div>' +
      (small ? '<div style="font-size:14px;color:' + SOFT + ';padding-top:2px">' + small + '</div>' : '') +
    '</td>';
  // שורות קבועות, נקודה רק בין שמות (לא בתחילת שורה ולא בסופה)
  const ROWS = [4, 5, 5, 5, 4];
  let i = 0;
  const names = ROWS.map(k => {
    const row = SPEAKERS.slice(i, i += k).map(n => '<span style="white-space:nowrap">' + esc_(n) + '</span>')
      .join('<span style="color:' + NAVY + ';font-weight:700;padding:0 7px">·</span>');
    return '<div style="margin:0 0 6px">' + row + '</div>';
  }).join('') + (i < SPEAKERS.length ? '<div>' + SPEAKERS.slice(i).map(esc_).join(' · ') + '</div>' : '');
  const cal = calendarUrl_();

  return '<!doctype html><html dir="rtl" lang="he"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">' +
  '<title>הזמנה: עיתונות בעידן ה־AI</title></head>' +
  '<body style="margin:0;padding:0;background:' + BG + '">' +
  '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="background:' + BG + '"><tr><td align="center" style="padding:24px 12px">' +
  '<table role="presentation" width="600" cellpadding="0" cellspacing="0" border="0" dir="rtl" style="width:100%;max-width:600px;font-family:' + FONT + ';color:' + INK + '">' +

    // כותרת: תמונה מהאתר, עם טקסט חלופי מלא
    '<tr><td style="background:' + NAVY + ';border-radius:16px 16px 0 0">' +
      '<a href="' + SITE + '"><img src="' + HEADER_IMG + '" width="600" alt="אוניברסיטת תל אביב, החוג לתקשורת והפקולטה לניהול. עיתונות בעידן ה־AI: מעצבים את כללי המשחק" ' +
      'style="display:block;width:100%;max-width:600px;height:auto;border:0;border-radius:16px 16px 0 0;color:#ffffff;font-size:22px;font-weight:700;background:' + NAVY + '"></a>' +
    '</td></tr>' +

    // פרטי הכנס על רקע כחול
    '<tr><td style="background:' + NAVY + ';padding:0 16px 8px">' +
      '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" dir="rtl"><tr>' +
        fact('DATE', '2.12.2026', 'יום רביעי') +
        fact('TIME', '<span dir="ltr" style="white-space:nowrap">15:00–19:00</span>', '') +
        fact('PLACE', 'אוניברסיטת תל אביב', 'בניין נפתלי, אולם&nbsp;אפטר', true) +
      '</tr></table>' +
    '</td></tr>' +

    // נוסח ההזמנה על רקע בהיר
    '<tr><td align="center" style="background:#ffffff;border-inline:1px solid ' + LINE + ';padding:32px 36px 8px">' +
      '<p style="margin:0 0 14px;font-size:17px;line-height:1.6;text-align:right">' + (name ? 'שלום ' + esc_(name) + ',' : 'שלום,') + '</p>' +
      '<p style="margin:0 0 12px;font-size:19px;font-weight:700;line-height:1.5;color:' + NAVY + '">החוג לתקשורת והפקולטה לניהול באוניברסיטת תל&nbsp;אביב מתכבדים להזמינכם לכנס</p>' +
      '<p style="margin:0 0 22px;font-size:16px;line-height:1.7;color:' + MUTED + '">יחד עם בכירי העיתונות, אנשי טכנולוגיה וחוקרי תקשורת נדון בשאלה כיצד להטמיע כלי בינה מלאכותית במערכות החדשות בישראל: איך לממש את יתרונותיהם ואיך לצמצם את הסיכונים.</p>' +
    '</td></tr>' +

    // דוברים
    '<tr><td align="center" style="background:#ffffff;border-inline:1px solid ' + LINE + ';padding:8px 18px 28px">' +
      '<div style="border-top:1px solid ' + LINE + ';padding-top:22px;font-size:16px;font-weight:700;color:' + NAVY + '">' + SPEAKERS.length + ' דוברים ודוברות</div>' +
      '<div style="margin:14px 0 0;font-size:14px;line-height:1.9;color:' + INK + '">' + names + '</div>' +
    '</td></tr>' +

    // הרשמה
    '<tr><td align="center" style="background:' + NAVY + ';border-radius:0 0 16px 16px;padding:30px 24px 32px">' +
      '<p style="margin:0 0 16px;font-size:15px;color:' + SOFT + '">ההשתתפות בהרשמה מראש</p>' +
      '<table role="presentation" cellpadding="0" cellspacing="0" border="0" align="center" style="margin:0 auto"><tr>' +
        '<td align="center" style="background:#ffffff;border-radius:999px">' +
          '<a href="' + SITE + '" style="display:inline-block;padding:14px 34px;font-size:17px;font-weight:700;color:' + NAVY + ';text-decoration:none;font-family:' + FONT + '">להרשמה לכנס &larr;</a>' +
        '</td>' +
      '</tr></table>' +
      '<p style="margin:16px 0 0;font-size:14px"><a href="' + cal + '" style="color:#ffffff;text-decoration:underline">הוספה ליומן</a>' +
      '<span style="color:' + SOFT + ';padding:0 10px">|</span>' +
      '<a href="' + SITE + '" style="color:#ffffff;text-decoration:underline" dir="ltr">aijournalism.co.il</a></p>' +
    '</td></tr>' +

    // תחתית
    '<tr><td align="center" style="padding:20px 16px 6px;font-size:13px;line-height:1.6;color:' + MUTED + '">' +
      'החוג לתקשורת ע"ש דן · הפקולטה לניהול ע"ש קולר<br>אוניברסיטת תל אביב' +
    '</td></tr>' +

  '</table></td></tr></table></body></html>';
}
