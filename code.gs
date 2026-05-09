/**
 * ศูนย์บริการประชาชนออนไลน์ รพ.สต.บ้านคำใหญ่
 * ฐานข้อมูล: Google Sheet
 * Sheet ID: 1AjIaB1dofobVMUZMQ4QTaEn0URU5ebr8jEkYaTaLHmA
 */

const SHEET_ID = '1AjIaB1dofobVMUZMQ4QTaEn0URU5ebr8jEkYaTaLHmA';

const SHEETS = {
  COMPLAINTS: 'Complaints',
  SETTINGS: 'Settings',
  USERS: 'Users',
  NEWS: 'News'
};

const COMPLAINT_HEADERS = [
  'case_id',
  'timestamp',
  'full_name',
  'phone',
  'village',
  'topic_type',
  'detail',
  'status',
  'recorder',
  'updated_at',
  'remark'
];

const USER_HEADERS = [
  'username',
  'password',
  'role',
  'status',
  'display_name',
  'created_at'
];

const NEWS_HEADERS = [
  'news_id',
  'created_at',
  'updated_at',
  'title',
  'category',
  'summary',
  'content',
  'image_url',
  'status',
  'author',
  'published_at'
];

function doGet(e) {
  setupSheets();

  return HtmlService
    .createTemplateFromFile('index')
    .evaluate()
    .setTitle('ศูนย์บริการประชาชน รพ.สต.บ้านคำใหญ่')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getSpreadsheet_() {
  return SpreadsheetApp.openById(SHEET_ID);
}

function getOrCreateSheet_(sheetName, headers) {
  const ss = getSpreadsheet_();
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  if (headers && headers.length) {
    const lastCol = sheet.getLastColumn();
    const firstRowValues = lastCol > 0
      ? sheet.getRange(1, 1, 1, Math.max(lastCol, headers.length)).getValues()[0]
      : [];

    const needsHeader = firstRowValues.filter(String).length === 0 ||
      headers.some((h, i) => firstRowValues[i] !== h);

    if (needsHeader) {
      sheet.clear();
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length)
        .setFontWeight('bold')
        .setBackground('#16a34a')
        .setFontColor('#ffffff');
      sheet.setFrozenRows(1);
      sheet.autoResizeColumns(1, headers.length);
      sheet.getRange('A:K').setNumberFormat('@');
    }
  }

  return sheet;
}

function setupSheets() {
  const complaintSheet = getOrCreateSheet_(SHEETS.COMPLAINTS, COMPLAINT_HEADERS);


  const newsSheet = getOrCreateSheet_(SHEETS.NEWS, NEWS_HEADERS);
  newsSheet.getRange('A:K').setNumberFormat('@');

  const usersSheet = getOrCreateSheet_(SHEETS.USERS, USER_HEADERS);
  if (usersSheet.getLastRow() < 2) {
    usersSheet.getRange(2, 1, 1, USER_HEADERS.length).setValues([[
      'admin',
      'pcu05374ky2512',
      'admin',
      'active',
      'ผู้ดูแลระบบ',
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss')
    ]]);
    usersSheet.getRange('A:F').setNumberFormat('@');
  }

  const settingsSheet = getOrCreateSheet_(SHEETS.SETTINGS, ['key', 'value', 'description']);
  if (settingsSheet.getLastRow() < 2) {
    settingsSheet.getRange(2, 1, 4, 3).setValues([
      ['organization', 'โรงพยาบาลส่งเสริมสุขภาพตำบลบ้านคำใหญ่', 'ชื่อหน่วยงาน'],
      ['phone', '043-840-180', 'เบอร์ติดต่อ'],
      ['line_oa', '@khamyaihealth', 'LINE OA'],
      ['email', 'khamyai.health@gmail.com', 'อีเมล']
    ]);
  }

  return {
    success: true,
    message: 'สร้าง/ตรวจสอบชีตเรียบร้อย',
    sheets: Object.values(SHEETS)
  };
}

function saveComplaint(payload) {
  try {
    setupSheets();

    if (!payload) {
      throw new Error('ไม่พบข้อมูลที่ส่งมา');
    }

    const fullName = String(payload.fullName || '').trim();
    const phone = String(payload.phone || '').trim();
    const detail = String(payload.detail || '').trim();

    if (!fullName || !phone || !detail) {
      throw new Error('กรุณากรอกชื่อ เบอร์โทร และรายละเอียด');
    }

    const now = new Date();
    const caseId = createCaseId_(now);
    const sheet = getOrCreateSheet_(SHEETS.COMPLAINTS, COMPLAINT_HEADERS);

    const row = [
      caseId,
      Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
      fullName,
      "'" + phone,
      String(payload.village || ''),
      String(payload.topicType || ''),
      detail,
      'รับเรื่องแล้ว',
      String(payload.recorder || ''),
      Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
      ''
    ];

    sheet.appendRow(row);
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 1, 1, COMPLAINT_HEADERS.length).setNumberFormat('@');

    return {
      success: true,
      caseId: caseId,
      message: 'บันทึกข้อมูลสำเร็จ'
    };
  } catch (err) {
    return {
      success: false,
      message: err.message || String(err)
    };
  }
}

function createCaseId_(dateObj) {
  const tz = Session.getScriptTimeZone();
  const datePart = Utilities.formatDate(dateObj, tz, 'yyyyMMdd');
  const randomPart = Math.floor(Math.random() * 9000) + 1000;
  return 'KY-' + datePart + '-' + randomPart;
}

function getDashboardData() {
  try {
    setupSheets();

    const sheet = getOrCreateSheet_(SHEETS.COMPLAINTS, COMPLAINT_HEADERS);
    const values = sheet.getDataRange().getValues();

    if (values.length <= 1) {
      return {
        success: true,
        summary: { today: 0, month: 0, year: 0, all: 0 },
        status: { received: 0, processing: 0, completed: 0 },
        topicTypes: {},
        villages: {}
      };
    }

    const headers = values[0];
    const rows = values.slice(1).filter(r => r.join('').trim() !== '');

    const idx = {};
    headers.forEach((h, i) => idx[h] = i);

    const now = new Date();
    const tz = Session.getScriptTimeZone();
    const todayKey = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
    const monthKey = Utilities.formatDate(now, tz, 'yyyy-MM');
    const yearKey = Utilities.formatDate(now, tz, 'yyyy');

    const summary = { today: 0, month: 0, year: 0, all: rows.length };
    const status = { received: 0, processing: 0, completed: 0 };
    const topicTypes = {};
    const villages = {};

    rows.forEach(row => {
      const timestampText = String(row[idx.timestamp] || '');
      const statusText = String(row[idx.status] || 'รับเรื่องแล้ว');
      const topicText = String(row[idx.topic_type] || 'ไม่ระบุ');
      const villageText = String(row[idx.village] || 'ไม่ระบุ');

      if (timestampText.indexOf(todayKey) === 0) summary.today++;
      if (timestampText.indexOf(monthKey) === 0) summary.month++;
      if (timestampText.indexOf(yearKey) === 0) summary.year++;

      if (statusText === 'เสร็จสิ้น') status.completed++;
      else if (statusText === 'กำลังดำเนินการ') status.processing++;
      else status.received++;

      topicTypes[topicText] = (topicTypes[topicText] || 0) + 1;
      villages[villageText] = (villages[villageText] || 0) + 1;
    });

    return {
      success: true,
      summary,
      status,
      topicTypes,
      villages
    };
  } catch (err) {
    return {
      success: false,
      message: err.message || String(err),
      summary: { today: 0, month: 0, year: 0, all: 0 },
      status: { received: 0, processing: 0, completed: 0 }
    };
  }
}

/**
 * ใช้ทดสอบจาก Apps Script Editor
 */
function testSetup() {
  return setupSheets();
}

function testSaveComplaint() {
  return saveComplaint({
    fullName: 'ทดสอบ ระบบ',
    phone: '043-840-180',
    village: 'หมู่ 7 บ้านคำใหญ่',
    topicType: 'ขอรับบริการสุขภาพ',
    detail: 'ทดสอบการบันทึกข้อมูลจาก Apps Script',
    recorder: 'admin'
  });
}


function adminLogin(username, password) {
  try {
    setupSheets();

    username = String(username || '').trim();
    password = String(password || '');

    if (!username || !password) {
      throw new Error('กรุณากรอก username และ password');
    }

    const sheet = getOrCreateSheet_(SHEETS.USERS, USER_HEADERS);
    const values = sheet.getDataRange().getValues();
    const headers = values[0];
    const idx = {};
    headers.forEach((h, i) => idx[h] = i);

    const userRow = values.slice(1).find(r =>
      String(r[idx.username]).trim() === username &&
      String(r[idx.password]) === password &&
      String(r[idx.role]).toLowerCase() === 'admin' &&
      String(r[idx.status]).toLowerCase() === 'active'
    );

    if (!userRow) {
      return {
        success: false,
        message: 'ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง หรือบัญชีไม่ได้เปิดใช้งาน'
      };
    }

    const token = Utilities.getUuid();
    CacheService.getScriptCache().put('ADMIN_TOKEN_' + token, username, 21600); // 6 ชั่วโมง

    return {
      success: true,
      token: token,
      displayName: userRow[idx.display_name] || username
    };
  } catch (err) {
    return {
      success: false,
      message: err.message || String(err)
    };
  }
}

function verifyAdminToken_(token) {
  token = String(token || '').trim();
  if (!token) return false;
  const username = CacheService.getScriptCache().get('ADMIN_TOKEN_' + token);
  return !!username;
}

function getAdminComplaintDashboard(token) {
  try {
    if (!verifyAdminToken_(token)) {
      throw new Error('Session หมดอายุ หรือยังไม่ได้เข้าสู่ระบบ กรุณา Login ใหม่');
    }

    setupSheets();

    const sheet = getOrCreateSheet_(SHEETS.COMPLAINTS, COMPLAINT_HEADERS);
    const values = sheet.getDataRange().getValues();

    if (values.length <= 1) {
      return {
        success: true,
        total: 0,
        rows: [],
        status: { received: 0, processing: 0, completed: 0 },
        topicTypes: {},
        villages: {},
        monthly: {}
      };
    }

    const headers = values[0];
    const idx = {};
    headers.forEach((h, i) => idx[h] = i);

    const rows = values.slice(1)
      .filter(r => r.join('').trim() !== '')
      .map(r => {
        return {
          caseId: String(r[idx.case_id] || ''),
          timestamp: String(r[idx.timestamp] || ''),
          fullName: String(r[idx.full_name] || ''),
          phone: String(r[idx.phone] || '').replace(/^'/, ''),
          village: String(r[idx.village] || 'ไม่ระบุ'),
          topicType: String(r[idx.topic_type] || 'ไม่ระบุ'),
          detail: String(r[idx.detail] || ''),
          status: String(r[idx.status] || 'รับเรื่องแล้ว'),
          recorder: String(r[idx.recorder] || ''),
          updatedAt: String(r[idx.updated_at] || ''),
          remark: String(r[idx.remark] || '')
        };
      })
      .reverse();

    const status = { received: 0, processing: 0, completed: 0 };
    const topicTypes = {};
    const villages = {};
    const monthly = {};

    rows.forEach(r => {
      if (r.status === 'เสร็จสิ้น') status.completed++;
      else if (r.status === 'กำลังดำเนินการ') status.processing++;
      else status.received++;

      topicTypes[r.topicType] = (topicTypes[r.topicType] || 0) + 1;
      villages[r.village] = (villages[r.village] || 0) + 1;

      const monthKey = r.timestamp ? r.timestamp.substring(0, 7) : 'ไม่ระบุ';
      monthly[monthKey] = (monthly[monthKey] || 0) + 1;
    });

    const sortedMonthly = {};
    Object.keys(monthly).sort().forEach(k => sortedMonthly[k] = monthly[k]);

    return {
      success: true,
      total: rows.length,
      rows: rows,
      status: status,
      topicTypes: topicTypes,
      villages: villages,
      monthly: sortedMonthly
    };
  } catch (err) {
    return {
      success: false,
      message: err.message || String(err)
    };
  }
}

function updateComplaintStatus(token, caseId, status, remark) {
  try {
    if (!verifyAdminToken_(token)) {
      throw new Error('Session หมดอายุ หรือยังไม่ได้เข้าสู่ระบบ กรุณา Login ใหม่');
    }

    const allowed = ['รับเรื่องแล้ว', 'กำลังดำเนินการ', 'เสร็จสิ้น'];
    if (allowed.indexOf(status) === -1) {
      throw new Error('สถานะไม่ถูกต้อง');
    }

    const sheet = getOrCreateSheet_(SHEETS.COMPLAINTS, COMPLAINT_HEADERS);
    const values = sheet.getDataRange().getValues();
    const headers = values[0];
    const idx = {};
    headers.forEach((h, i) => idx[h] = i);

    for (let i = 1; i < values.length; i++) {
      if (String(values[i][idx.case_id]) === String(caseId)) {
        const rowNo = i + 1;
        sheet.getRange(rowNo, idx.status + 1).setValue(status);
        sheet.getRange(rowNo, idx.updated_at + 1).setValue(
          Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss')
        );
        sheet.getRange(rowNo, idx.remark + 1).setValue(String(remark || ''));
        return { success: true, message: 'อัปเดตสถานะสำเร็จ' };
      }
    }

    throw new Error('ไม่พบเลขรับเรื่องนี้');
  } catch (err) {
    return {
      success: false,
      message: err.message || String(err)
    };
  }
}


function getPublishedNews(limit) {
  try {
    setupSheets();
    limit = Number(limit || 6);

    const sheet = getOrCreateSheet_(SHEETS.NEWS, NEWS_HEADERS);
    const rows = getNewsRows_()
      .filter(n => n.status === 'เผยแพร่')
      .sort((a, b) => String(b.publishedAt || b.createdAt).localeCompare(String(a.publishedAt || a.createdAt)))
      .slice(0, limit);

    return {
      success: true,
      rows: rows
    };
  } catch (err) {
    return {
      success: false,
      message: err.message || String(err),
      rows: []
    };
  }
}

function getNewsById(newsId) {
  try {
    setupSheets();
    const news = getNewsRows_().find(n => n.newsId === String(newsId || ''));
    if (!news) throw new Error('ไม่พบข่าวนี้');
    if (news.status !== 'เผยแพร่') throw new Error('ข่าวนี้ยังไม่เผยแพร่');

    return {
      success: true,
      news: news
    };
  } catch (err) {
    return {
      success: false,
      message: err.message || String(err)
    };
  }
}

function getAdminNews(token) {
  try {
    if (!verifyAdminToken_(token)) {
      throw new Error('Session หมดอายุ หรือยังไม่ได้เข้าสู่ระบบ กรุณา Login ใหม่');
    }

    setupSheets();

    const rows = getNewsRows_()
      .sort((a, b) => String(b.updatedAt || b.createdAt).localeCompare(String(a.updatedAt || a.createdAt)));

    return {
      success: true,
      rows: rows
    };
  } catch (err) {
    return {
      success: false,
      message: err.message || String(err),
      rows: []
    };
  }
}

function saveNews(token, payload) {
  try {
    if (!verifyAdminToken_(token)) {
      throw new Error('Session หมดอายุ หรือยังไม่ได้เข้าสู่ระบบ กรุณา Login ใหม่');
    }

    payload = payload || {};
    const title = String(payload.title || '').trim();
    const summary = String(payload.summary || '').trim();

    if (!title || !summary) {
      throw new Error('กรุณากรอกหัวข้อข่าวและสรุปข่าว');
    }

    const sheet = getOrCreateSheet_(SHEETS.NEWS, NEWS_HEADERS);
    const values = sheet.getDataRange().getValues();
    const headers = values[0];
    const idx = {};
    headers.forEach((h, i) => idx[h] = i);

    const now = new Date();
    const nowText = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    const status = String(payload.status || 'เผยแพร่');
    const newsId = String(payload.newsId || '').trim() || createNewsId_(now);

    let foundRow = -1;
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][idx.news_id]) === newsId) {
        foundRow = i + 1;
        break;
      }
    }

    const publishedAt = status === 'เผยแพร่' ? nowText : '';

    if (foundRow > -1) {
      const oldPublishedAt = String(sheet.getRange(foundRow, idx.published_at + 1).getValue() || '');
      sheet.getRange(foundRow, idx.updated_at + 1).setValue(nowText);
      sheet.getRange(foundRow, idx.title + 1).setValue(title);
      sheet.getRange(foundRow, idx.category + 1).setValue(String(payload.category || 'ข่าวประชาสัมพันธ์'));
      sheet.getRange(foundRow, idx.summary + 1).setValue(summary);
      sheet.getRange(foundRow, idx.content + 1).setValue(String(payload.content || ''));
      sheet.getRange(foundRow, idx.image_url + 1).setValue(String(payload.imageUrl || ''));
      sheet.getRange(foundRow, idx.status + 1).setValue(status);
      sheet.getRange(foundRow, idx.published_at + 1).setValue(status === 'เผยแพร่' ? (oldPublishedAt || publishedAt) : '');
    } else {
      sheet.appendRow([
        newsId,
        nowText,
        nowText,
        title,
        String(payload.category || 'ข่าวประชาสัมพันธ์'),
        summary,
        String(payload.content || ''),
        String(payload.imageUrl || ''),
        status,
        'admin',
        publishedAt
      ]);
      sheet.getRange(sheet.getLastRow(), 1, 1, NEWS_HEADERS.length).setNumberFormat('@');
    }

    return {
      success: true,
      newsId: newsId,
      message: 'บันทึกข่าวสำเร็จ'
    };
  } catch (err) {
    return {
      success: false,
      message: err.message || String(err)
    };
  }
}

function updateNewsStatus(token, newsId, status) {
  try {
    if (!verifyAdminToken_(token)) {
      throw new Error('Session หมดอายุ หรือยังไม่ได้เข้าสู่ระบบ กรุณา Login ใหม่');
    }

    const allowed = ['เผยแพร่', 'ฉบับร่าง'];
    if (allowed.indexOf(status) === -1) throw new Error('สถานะข่าวไม่ถูกต้อง');

    const sheet = getOrCreateSheet_(SHEETS.NEWS, NEWS_HEADERS);
    const values = sheet.getDataRange().getValues();
    const headers = values[0];
    const idx = {};
    headers.forEach((h, i) => idx[h] = i);

    const nowText = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

    for (let i = 1; i < values.length; i++) {
      if (String(values[i][idx.news_id]) === String(newsId)) {
        const rowNo = i + 1;
        sheet.getRange(rowNo, idx.status + 1).setValue(status);
        sheet.getRange(rowNo, idx.updated_at + 1).setValue(nowText);
        sheet.getRange(rowNo, idx.published_at + 1).setValue(status === 'เผยแพร่' ? nowText : '');
        return { success: true, message: 'อัปเดตสถานะข่าวสำเร็จ' };
      }
    }

    throw new Error('ไม่พบข่าวนี้');
  } catch (err) {
    return {
      success: false,
      message: err.message || String(err)
    };
  }
}

function deleteNews(token, newsId) {
  try {
    if (!verifyAdminToken_(token)) {
      throw new Error('Session หมดอายุ หรือยังไม่ได้เข้าสู่ระบบ กรุณา Login ใหม่');
    }

    const sheet = getOrCreateSheet_(SHEETS.NEWS, NEWS_HEADERS);
    const values = sheet.getDataRange().getValues();
    const headers = values[0];
    const idx = {};
    headers.forEach((h, i) => idx[h] = i);

    for (let i = 1; i < values.length; i++) {
      if (String(values[i][idx.news_id]) === String(newsId)) {
        sheet.deleteRow(i + 1);
        return { success: true, message: 'ลบข่าวสำเร็จ' };
      }
    }

    throw new Error('ไม่พบข่าวนี้');
  } catch (err) {
    return {
      success: false,
      message: err.message || String(err)
    };
  }
}

function getNewsRows_() {
  const sheet = getOrCreateSheet_(SHEETS.NEWS, NEWS_HEADERS);
  const values = sheet.getDataRange().getValues();

  if (values.length <= 1) return [];

  const headers = values[0];
  const idx = {};
  headers.forEach((h, i) => idx[h] = i);

  return values.slice(1)
    .filter(r => r.join('').trim() !== '')
    .map(r => ({
      newsId: String(r[idx.news_id] || ''),
      createdAt: String(r[idx.created_at] || ''),
      updatedAt: String(r[idx.updated_at] || ''),
      title: String(r[idx.title] || ''),
      category: String(r[idx.category] || ''),
      summary: String(r[idx.summary] || ''),
      content: String(r[idx.content] || ''),
      imageUrl: String(r[idx.image_url] || ''),
      status: String(r[idx.status] || 'ฉบับร่าง'),
      author: String(r[idx.author] || ''),
      publishedAt: String(r[idx.published_at] || '')
    }));
}

function createNewsId_(dateObj) {
  const tz = Session.getScriptTimeZone();
  const datePart = Utilities.formatDate(dateObj, tz, 'yyyyMMdd');
  const randomPart = Math.floor(Math.random() * 9000) + 1000;
  return 'NEWS-' + datePart + '-' + randomPart;
}

function testSaveNews() {
  return saveNews(
    CacheService.getScriptCache().put('ADMIN_TOKEN_TEST', 'admin', 60) || 'TEST',
    {
      title: 'ทดสอบข่าวประชาสัมพันธ์',
      category: 'ข่าวประชาสัมพันธ์',
      summary: 'ทดสอบการบันทึกข่าวประชาสัมพันธ์',
      content: 'รายละเอียดข่าวประชาสัมพันธ์สำหรับทดสอบระบบ',
      imageUrl: '',
      status: 'เผยแพร่'
    }
  );
}
