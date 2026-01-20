/**
 * 源泉徴収 系メニュー
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('源泉徴収')
    .addItem('年次集計を作成', 'gensenCreateAnnualSummary')
    .addItem('計算台を開く', 'gensenOpenCalcSheet')
    .addItem('確定データに反映', 'gensenUpsertConfirmedData')
    .addItem('帳票を生成', 'gensenGenerateReports')
    .addToUi();
}

/**
 * === 設定（必要最低限） ===
 * シート名・列名は既存の構成に合わせて調整してください。
 */
const GENSEN_CONFIG = {
  settingsSheetName: '設定',
  settingsYearCell: 'B2',

  monthlySheetName: '月次給与データ',
  monthlyHeaderRow: 1,
  monthlyColumns: {
    year: '年度',
    employeeId: '従業員ID',
    name: '氏名',
    gross: '年間総支給額',
    socialInsurance: '社会保険料合計',
    withholdingTax: '源泉徴収税額合計',
  },

  annualSummarySheetPrefix: '年次集計_',
  calcSheetPrefix: '源泉徴収_計算台_',
  confirmedSheetPrefix: '源泉徴収_確定データ_',
  reportTemplateSheetName: '源泉徴収票_帳票テンプレ',

  confirmedFields: [
    { header: '従業員ID', rangeName: '源泉徴収_従業員ID' },
    { header: '氏名', rangeName: '源泉徴収_氏名' },
    { header: '年間総支給額', rangeName: '源泉徴収_年間総支給額' },
    { header: '社会保険料合計', rangeName: '源泉徴収_社会保険料合計' },
    { header: '源泉徴収税額合計', rangeName: '源泉徴収_源泉徴収税額合計' },
    { header: '所得税額', rangeName: '源泉徴収_所得税額' },
    { header: '過不足', rangeName: '源泉徴収_過不足' },
  ],

  reportFields: [
    { header: '従業員ID', rangeName: '帳票_従業員ID' },
    { header: '氏名', rangeName: '帳票_氏名' },
    { header: '年間総支給額', rangeName: '帳票_年間総支給額' },
    { header: '社会保険料合計', rangeName: '帳票_社会保険料合計' },
    { header: '源泉徴収税額合計', rangeName: '帳票_源泉徴収税額合計' },
    { header: '所得税額', rangeName: '帳票_所得税額' },
    { header: '過不足', rangeName: '帳票_過不足' },
  ],

  reportFolderName: '源泉徴収',
};

function gensenCreateAnnualSummary() {
  const ss = SpreadsheetApp.getActive();
  const year = gensenGetTargetYear_();
  const monthlySheet = ss.getSheetByName(GENSEN_CONFIG.monthlySheetName);
  if (!monthlySheet) {
    throw new Error('月次給与データのシートが見つかりません。設定を確認してください。');
  }

  const dataRange = monthlySheet.getDataRange();
  const values = dataRange.getValues();
  if (values.length <= GENSEN_CONFIG.monthlyHeaderRow) {
    throw new Error('月次給与データにヘッダー以外の行がありません。');
  }

  const headers = values[GENSEN_CONFIG.monthlyHeaderRow - 1];
  const headerIndex = gensenBuildHeaderIndex_(headers);
  const col = GENSEN_CONFIG.monthlyColumns;
  const required = [col.year, col.employeeId, col.name, col.gross, col.socialInsurance, col.withholdingTax];
  required.forEach((name) => {
    if (!(name in headerIndex)) {
      throw new Error('月次給与データのヘッダーに必要な列がありません: ' + name);
    }
  });

  const summary = {};
  values.slice(GENSEN_CONFIG.monthlyHeaderRow).forEach((row) => {
    const rowYear = String(row[headerIndex[col.year]]).trim();
    if (rowYear !== String(year)) {
      return;
    }
    const employeeId = String(row[headerIndex[col.employeeId]]).trim();
    if (!employeeId) {
      return;
    }
    if (!summary[employeeId]) {
      summary[employeeId] = {
        employeeId,
        name: row[headerIndex[col.name]],
        gross: 0,
        socialInsurance: 0,
        withholdingTax: 0,
      };
    }
    summary[employeeId].gross += Number(row[headerIndex[col.gross]]) || 0;
    summary[employeeId].socialInsurance += Number(row[headerIndex[col.socialInsurance]]) || 0;
    summary[employeeId].withholdingTax += Number(row[headerIndex[col.withholdingTax]]) || 0;
  });

  const sheetName = GENSEN_CONFIG.annualSummarySheetPrefix + year;
  const summarySheet = gensenGetOrCreateSheet_(ss, sheetName);
  summarySheet.clearContents();

  const output = [
    ['従業員ID', '氏名', '年間総支給額', '社会保険料合計', '源泉徴収税額合計'],
  ];
  Object.keys(summary)
    .sort()
    .forEach((employeeId) => {
      const item = summary[employeeId];
      output.push([
        item.employeeId,
        item.name,
        item.gross,
        item.socialInsurance,
        item.withholdingTax,
      ]);
    });

  summarySheet.getRange(1, 1, output.length, output[0].length).setValues(output);
}

function gensenOpenCalcSheet() {
  const ss = SpreadsheetApp.getActive();
  const year = gensenGetTargetYear_();
  const sheetName = GENSEN_CONFIG.calcSheetPrefix + year;
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error('計算台シートが見つかりません: ' + sheetName);
  }
  ss.setActiveSheet(sheet);
}

function gensenUpsertConfirmedData() {
  const ss = SpreadsheetApp.getActive();
  const year = gensenGetTargetYear_();
  const calcSheetName = GENSEN_CONFIG.calcSheetPrefix + year;
  const calcSheet = ss.getSheetByName(calcSheetName);
  if (!calcSheet) {
    throw new Error('計算台シートが見つかりません: ' + calcSheetName);
  }

  const confirmedSheetName = GENSEN_CONFIG.confirmedSheetPrefix + year;
  const confirmedSheet = gensenGetOrCreateSheet_(ss, confirmedSheetName);

  const confirmedFields = GENSEN_CONFIG.confirmedFields.map((field) => {
    const range = ss.getRangeByName(field.rangeName);
    if (!range) {
      throw new Error('Named Range が見つかりません: ' + field.rangeName);
    }
    return {
      header: field.header,
      value: range.getValue(),
    };
  });

  const headerRow = confirmedFields.map((field) => field.header);
  const lastRow = Math.max(confirmedSheet.getLastRow(), 1);
  const existingHeaders = confirmedSheet
    .getRange(1, 1, 1, confirmedSheet.getLastColumn() || headerRow.length)
    .getValues()[0];
  const headersMatch = headerRow.every((header, index) => header === existingHeaders[index]);
  if (!headersMatch) {
    confirmedSheet.clearContents();
    confirmedSheet.getRange(1, 1, 1, headerRow.length).setValues([headerRow]);
  }

  const employeeId = confirmedFields.find((field) => field.header === '従業員ID');
  if (!employeeId || !employeeId.value) {
    throw new Error('従業員IDが取得できません。計算台のNamed Rangeを確認してください。');
  }

  const dataRange = confirmedSheet.getRange(2, 1, Math.max(lastRow - 1, 0), headerRow.length);
  const dataValues = dataRange.getValues();
  const existingRowIndex = dataValues.findIndex((row) => String(row[0]) === String(employeeId.value));
  const rowValues = confirmedFields.map((field) => field.value);

  if (existingRowIndex >= 0) {
    confirmedSheet
      .getRange(existingRowIndex + 2, 1, 1, rowValues.length)
      .setValues([rowValues]);
  } else {
    confirmedSheet.appendRow(rowValues);
  }
}

function gensenGenerateReports() {
  const ss = SpreadsheetApp.getActive();
  const year = gensenGetTargetYear_();
  const confirmedSheetName = GENSEN_CONFIG.confirmedSheetPrefix + year;
  const confirmedSheet = ss.getSheetByName(confirmedSheetName);
  if (!confirmedSheet) {
    throw new Error('確定データシートが見つかりません: ' + confirmedSheetName);
  }

  const templateSheet = ss.getSheetByName(GENSEN_CONFIG.reportTemplateSheetName);
  if (!templateSheet) {
    throw new Error('帳票テンプレートシートが見つかりません。');
  }

  const values = confirmedSheet.getDataRange().getValues();
  if (values.length < 2) {
    throw new Error('確定データがありません。');
  }

  const headers = values[0];
  const folder = gensenGetOrCreateFolder_(GENSEN_CONFIG.reportFolderName, year);
  const reportFields = GENSEN_CONFIG.reportFields.map((field) => {
    const range = ss.getRangeByName(field.rangeName);
    if (!range) {
      throw new Error('帳票テンプレートのNamed Range が見つかりません: ' + field.rangeName);
    }
    if (range.getSheet().getName() !== templateSheet.getName()) {
      throw new Error('帳票テンプレート以外のNamed Rangeが指定されています: ' + field.rangeName);
    }
    return { header: field.header, range };
  });

  values.slice(1).forEach((row) => {
    const record = {};
    headers.forEach((header, index) => {
      record[header] = row[index];
    });
    const employeeName = record['氏名'] || 'unknown';

    reportFields.forEach((field) => {
      field.range.setValue(record[field.header]);
    });

    const pdfBlob = gensenExportSheetToPdf_(ss, templateSheet, employeeName + '.pdf');
    folder.createFile(pdfBlob);
  });
}

function gensenGetTargetYear_() {
  const ss = SpreadsheetApp.getActive();
  const settingsSheet = ss.getSheetByName(GENSEN_CONFIG.settingsSheetName);
  if (!settingsSheet) {
    throw new Error('設定シートが見つかりません。');
  }
  const year = settingsSheet.getRange(GENSEN_CONFIG.settingsYearCell).getValue();
  if (!year) {
    throw new Error('設定!B2 に対象年度を入力してください。');
  }
  return String(year).trim();
}

function gensenBuildHeaderIndex_(headers) {
  const index = {};
  headers.forEach((header, i) => {
    index[String(header).trim()] = i;
  });
  return index;
}

function gensenGetOrCreateSheet_(ss, sheetName) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  return sheet;
}

function gensenGetOrCreateFolder_(rootName, year) {
  const rootFolder = gensenFindOrCreateFolder_(DriveApp.getRootFolder(), rootName);
  return gensenFindOrCreateFolder_(rootFolder, String(year));
}

function gensenFindOrCreateFolder_(parent, name) {
  const iterator = parent.getFoldersByName(name);
  if (iterator.hasNext()) {
    return iterator.next();
  }
  return parent.createFolder(name);
}

function gensenExportSheetToPdf_(ss, sheet, filename) {
  const url =
    'https://docs.google.com/spreadsheets/d/' +
    ss.getId() +
    '/export' +
    '?format=pdf' +
    '&gid=' +
    sheet.getSheetId() +
    '&size=A4' +
    '&portrait=true' +
    '&fitw=true' +
    '&sheetnames=false&printtitle=false&pagenumbers=false&gridlines=false&fzr=false';

  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(url, {
    headers: {
      Authorization: 'Bearer ' + token,
    },
  });
  const blob = response.getBlob().setName(filename);
  return blob;
}
