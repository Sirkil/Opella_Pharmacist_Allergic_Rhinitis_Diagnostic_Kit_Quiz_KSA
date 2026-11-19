const SPREADSHEET_ID = '1xiumalIElJlSdx4zfLLH0pMVqMby6G3_zV7HfY75wN4';
const SHEET_NAME = 'Sheet1';
const SUBMISSION_ID_COLUMN = 10;
const MAX_SUBMISSIONS = 2500;
const LOCK_AT = Date.parse('2027-11-19T00:00:00+02:00'); // Midnight Cairo time.

const UUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const YES = 'نعم';
const NO = 'لا';
const SKIPPED = 'لم يُطرح';
const THICK = 'سميكة أو مائلة إلى الصفرة';
const WATERY = 'خفيفة وشفافة ومائية';

// Each path is ordered by the five question columns, D through H in Sheet1.
const QUIZ_PATHS = [
  { answers: [YES, YES, SKIPPED, SKIPPED, SKIPPED], result: 'إنفلونزا محتملة' },
  { answers: [YES, NO, SKIPPED, SKIPPED, SKIPPED], result: 'نزلة برد محتملة' },
  { answers: [NO, SKIPPED, THICK, SKIPPED, SKIPPED], result: 'نزلة برد محتملة' },
  { answers: [NO, SKIPPED, WATERY, NO, SKIPPED], result: 'أسباب أخرى لالتهاب الأنف · يلزم تقييم إضافي' },
  { answers: [NO, SKIPPED, WATERY, YES, YES], result: 'التهاب أنف تحسسي متفاقم بتلوث الهواء' },
  { answers: [NO, SKIPPED, WATERY, YES, NO], result: 'التهاب أنف تحسسي · محفزات محتملة أخرى' }
];

function doGet(e) {
  const params = (e && e.parameter) || {};
  const callback = params.callback || '';
  if (callback && !/^quizStatus_[A-Za-z0-9_]{1,64}$/.test(callback)) {
    return response({ ok: false, error: 'Invalid callback' });
  }

  try {
    const sheet = getSubmissionSheet();
    const lastRow = sheet.getLastRow();
    return response({
      ok: true,
      locked: isLocked(sheet, lastRow),
      accepted: hasSubmission(sheet, lastRow, params.submissionId)
    }, callback);
  } catch (error) {
    console.error(error);
    return response({ ok: false, error: 'Unable to check quiz status' }, callback);
  }
}

function doPost(e) {
  let data;
  try {
    data = JSON.parse(e.postData.contents);
  } catch (error) {
    return response({ ok: false, error: 'Invalid JSON' });
  }

  const path = validateSubmission(data);
  if (!path) return response({ ok: false, error: 'Invalid submission' });

  try {
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const sheet = getSubmissionSheet();
      const lastRow = sheet.getLastRow();
      const id = data.submissionId.toLowerCase();

      // A retry for a saved result can succeed even if the quiz has since closed.
      if (hasSubmission(sheet, lastRow, id)) {
        return response({ ok: true, accepted: true, duplicate: true });
      }
      if (isLocked(sheet, lastRow)) {
        return response({ ok: false, locked: true, error: 'Quiz locked' });
      }

      const row = lastRow + 1;
      const maxRows = sheet.getMaxRows();
      if (row > maxRows) {
        sheet.insertRowsAfter(maxRows, Math.max(row, MAX_SUBMISSIONS + 1) - maxRows);
      }

      const values = [new Date(), data.age, data.gender, ...path.answers, path.result, id];
      sheet.getRange(row, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
      sheet.getRange(row, 1, 1, values.length).setValues([values]);
      SpreadsheetApp.flush();
      return response({ ok: true, accepted: true });
    } finally {
      lock.releaseLock();
    }
  } catch (error) {
    console.error(error);
    return response({ ok: false, error: 'Unable to save submission' });
  }
}

function validateSubmission(data) {
  if (!data || data.schemaVersion !== 2 ||
      typeof data.submissionId !== 'string' || !UUID.test(data.submissionId) ||
      !Number.isInteger(data.age) || data.age < 1 || data.age > 120 ||
      !['ذكر', 'أنثى'].includes(data.gender) ||
      !Array.isArray(data.answersV2) || data.answersV2.length !== 5) {
    return null;
  }

  return QUIZ_PATHS.find(path => path.result === data.result &&
    path.answers.every((answer, index) => answer === data.answersV2[index]));
}

function getSubmissionSheet() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('Sheet1 was not found');
  return sheet;
}

function hasSubmission(sheet, lastRow, id) {
  if (lastRow < 2 || typeof id !== 'string' || !UUID.test(id)) return false;
  return Boolean(sheet.getRange(2, SUBMISSION_ID_COLUMN, lastRow - 1, 1)
    .createTextFinder(id).matchEntireCell(true).findNext());
}

function isLocked(sheet, lastRow) {
  if (Date.now() >= LOCK_AT) return true;
  if (lastRow <= MAX_SUBMISSIONS) return false;

  // Count recorded timestamps so blank rows and notes in other columns do not use the quota.
  const timestamps = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  return timestamps.filter(([timestamp]) => timestamp !== '').length >= MAX_SUBMISSIONS || Date.now() >= LOCK_AT;
}

function response(value, callback = '') {
  const json = JSON.stringify(value);
  if (callback) {
    return ContentService.createTextOutput(`${callback}(${json});`)
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}
