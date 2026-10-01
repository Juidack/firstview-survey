const SHEET_NAME = '回答データ';
const IMAGE_FOLDER_NAME = '卒業研究_ファーストビュー_画像';

const PAGES = [
  ['企業1', '01_kokuyo.png'],
  ['企業2', '02_asics.png'],
  ['企業3', '03_jrtokai.png'],
  ['企業4', '04_fastretailing.png'],
  ['企業5', '05_suntory.png'],
  ['用語・概念1', '06_taiseikai.png'],
  ['用語・概念2', '07_sky.png'],
  ['用語・概念3', '08_desknet.png'],
  ['用語・概念4', '09_nisa.png'],
  ['用語・概念5', '10_env.png'],
  ['方法・対処1', '11_suido.png'],
  ['方法・対処2', '12_osoujihonpo.png'],
  ['方法・対処3', '13_kagome.png'],
  ['方法・対処4', '14_orangepage.png'],
  ['方法・対処5', '15_taisho.png']
];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Webページ・ファーストビュー調査')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/*
 * GitHub PagesからのPOSTを受け取る。
 * e.parameter は application/x-www-form-urlencoded の値を取得する。
 */
function doPost(e) {
  try {
    const p = e && e.parameter ? e.parameter : {};

    const data = {
      grade: String(p.grade || ''),
      answers: JSON.parse(String(p.answers || '[]')),
      responseId: String(p.responseId || '')
    };

    const result = saveResponse(data);

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({
        success: false,
        error: String(err && err.message ? err.message : err)
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getImageUrls() {
  const folder = getImageFolder_();
  const result = {};
  const missing = [];

  PAGES.forEach(([, filename]) => {
    const files = folder.getFilesByName(filename);

    if (!files.hasNext()) {
      missing.push(filename);
      return;
    }

    const file = files.next();
    const id = file.getId();

    result[filename] =
      'https://drive.google.com/uc?export=view&id=' +
      encodeURIComponent(id);
  });

  return {
    success: missing.length === 0,
    urls: result,
    missing: missing
  };
}

function checkImages() {
  const info = getImageUrls();

  return {
    folderUrl: getImageFolder_().getUrl(),
    missing: info.missing
  };
}

function getImageFolder_() {
  const folders = DriveApp.getFoldersByName(IMAGE_FOLDER_NAME);

  if (folders.hasNext()) {
    return folders.next();
  }

  return DriveApp.createFolder(IMAGE_FOLDER_NAME);
}

function setupImagesFolder() {
  const folder = getImageFolder_();

  return {
    name: folder.getName(),
    url: folder.getUrl(),
    message: 'このフォルダに15枚の画像を入れてください。'
  };
}

function saveResponse(data) {
  if (!data || typeof data !== 'object') {
    throw new Error('回答データがありません。');
  }

  const grade = String(data.grade || '');
  const answers = Array.isArray(data.answers) ? data.answers : [];
  const responseId = String(data.responseId || '');

  if (!/^[1-6]$/.test(grade)) {
    throw new Error('学年が正しくありません。');
  }

  if (
    answers.length !== 15 ||
    answers.some(v => !/^[1-5]$/.test(String(v)))
  ) {
    throw new Error('15ページすべてに回答してください。');
  }

  if (!responseId) {
    throw new Error('回答IDがありません。');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow([
        '回答ID',
        '回答日時',
        '学年',
        ...Array.from({ length: 15 }, (_, i) => 'ページ' + (i + 1))
      ]);
    }

    const lastRow = sheet.getLastRow();

    if (lastRow >= 2) {
      const ids = sheet
        .getRange(2, 1, lastRow - 1, 1)
        .getValues()
        .flat()
        .map(String);

      if (ids.includes(responseId)) {
        return { success: true, duplicate: true };
      }
    }

    sheet.appendRow([
      responseId,
      new Date(),
      Number(grade),
      ...answers.map(Number)
    ]);

    return { success: true, duplicate: false };

  } finally {
    lock.releaseLock();
  }
}