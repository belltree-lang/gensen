/**
 * 源泉徴収 系メニュー
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('源泉徴収')
    .addItem('年次集計を作成', 'gensenCreateAnnualSummary')
    .addItem('年次集計を計算台に反映', 'gensenReflectAnnualSummaryToCalcSheet')
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

  annualSummarySheetPrefix: '年次集計_',
  calcSheetPrefix: '源泉徴収_計算台_',
  confirmedSheetPrefix: '源泉徴収_確定データ_',
  reportTemplateSheetName: '源泉徴収票_帳票テンプレ',
  summaryHeaders: ['従業員番号', '総支給額', '社会保険', '雇用保険', '源泉所得税'],
  optionalSummaryHeaders: ['扶養人数'],

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
  reflectTargets: {
    employeeIdRangeName: '源泉徴収_従業員ID',
    grossRangeName: '源泉徴収_年間総支給額',
    socialInsuranceRangeName: '源泉徴収_社会保険料合計',
    withholdingTaxRangeName: '源泉徴収_源泉徴収税額合計',
    dependentsRangeName: '源泉徴収_扶養人数',
  },
};

function gensenCreateAnnualSummary() {
  const ss = SpreadsheetApp.getActive();
  const year = gensenGetTargetYear_();
  const summary = {};
  const targetYear = String(year);
  const columnMap = {
    month: 2, // B: 月分（YYYY/MM）
    employeeId: 3, // C: 従業員番号
    employeeName: 4, // D: 従業員名
    gross: 15, // O: 総支給額
    socialInsurance: 16, // P: 社会保険
    employmentInsurance: 17, // Q: 雇用保険
    withholdingTax: 19, // S: 源泉所得税
    dependents: 20, // T: 扶養人数（参考）
  };
  const excludedNames = new Set([
    GENSEN_CONFIG.settingsSheetName,
    GENSEN_CONFIG.reportTemplateSheetName,
  ]);

  ss.getSheets().forEach((sheet) => {
    const sheetName = sheet.getName();
    if (
      excludedNames.has(sheetName) ||
      sheetName.startsWith(GENSEN_CONFIG.annualSummarySheetPrefix) ||
      sheetName.startsWith(GENSEN_CONFIG.calcSheetPrefix) ||
      sheetName.startsWith(GENSEN_CONFIG.confirmedSheetPrefix)
    ) {
      return;
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return;
    }

    const numRows = lastRow - 1;
    const monthValues = sheet.getRange(2, columnMap.month, numRows, 1).getValues();
    const employeeIdValues = sheet.getRange(2, columnMap.employeeId, numRows, 1).getValues();
    const employeeNameValues = sheet.getRange(2, columnMap.employeeName, numRows, 1).getValues();
    const grossValues = sheet.getRange(2, columnMap.gross, numRows, 1).getValues();
    const socialInsuranceValues = sheet.getRange(
      2,
      columnMap.socialInsurance,
      numRows,
      1
    ).getValues();
    const employmentInsuranceValues = sheet.getRange(
      2,
      columnMap.employmentInsurance,
      numRows,
      1
    ).getValues();
    const withholdingTaxValues = sheet.getRange(
      2,
      columnMap.withholdingTax,
      numRows,
      1
    ).getValues();
    const dependentsValues = sheet.getRange(2, columnMap.dependents, numRows, 1).getValues();

    monthValues.forEach((row, index) => {
      const rowNumber = index + 2;
      const monthValue = row[0];
      const rowYear = gensenExtractYearFromMonth_(monthValue);
      if (!rowYear) {
        gensenLogAnnualSummaryIssue_(
          sheetName,
          rowNumber,
          '月分がYYYY/MM形式でパースできません。',
          monthValue
        );
        return;
      }
      if (rowYear !== targetYear) {
        return;
      }
      const employeeId = String(employeeIdValues[index][0]).trim();
      if (!employeeId) {
        gensenLogAnnualSummaryIssue_(
          sheetName,
          rowNumber,
          '従業員番号が空です。',
          employeeIdValues[index][0]
        );
        return;
      }
      const employeeName = String(employeeNameValues[index][0] || '').trim();
      const gross = gensenParseAnnualSummaryAmount_(
        grossValues[index][0],
        sheetName,
        rowNumber,
        '総支給額'
      );
      const socialInsurance = gensenParseAnnualSummaryAmount_(
        socialInsuranceValues[index][0],
        sheetName,
        rowNumber,
        '社会保険'
      );
      const employmentInsurance = gensenParseAnnualSummaryAmount_(
        employmentInsuranceValues[index][0],
        sheetName,
        rowNumber,
        '雇用保険'
      );
      const withholdingTax = gensenParseAnnualSummaryAmount_(
        withholdingTaxValues[index][0],
        sheetName,
        rowNumber,
        '源泉所得税'
      );
      const dependents = Number(dependentsValues[index][0]) || 0;
      if (!summary[employeeId]) {
        summary[employeeId] = {
          employeeId,
          employeeName,
          dependents,
          gross: 0,
          socialInsurance: 0,
          employmentInsurance: 0,
          withholdingTax: 0,
        };
      } else {
        if (employeeName) {
          summary[employeeId].employeeName = employeeName;
        }
        summary[employeeId].dependents = dependents;
      }
      summary[employeeId].gross += gross;
      summary[employeeId].socialInsurance += socialInsurance;
      summary[employeeId].employmentInsurance += employmentInsurance;
      summary[employeeId].withholdingTax += withholdingTax;
    });
  });

  const sheetName = GENSEN_CONFIG.annualSummarySheetPrefix + year;
  const summarySheet = gensenGetOrCreateSheet_(ss, sheetName);
  const existingHeaders = gensenGetSheetHeaders_(summarySheet);
  const includeDependents = existingHeaders.includes('扶養人数');
  summarySheet.clearContents();

  const summaryHeaders = includeDependents
    ? GENSEN_CONFIG.summaryHeaders.concat(GENSEN_CONFIG.optionalSummaryHeaders)
    : GENSEN_CONFIG.summaryHeaders;
  const output = [summaryHeaders];
  Object.keys(summary)
    .sort()
    .forEach((employeeId) => {
      const item = summary[employeeId];
      const row = [
        item.employeeId,
        item.gross,
        item.socialInsurance,
        item.employmentInsurance,
        item.withholdingTax,
      ];
      if (includeDependents) {
        row.push(item.dependents);
      }
      output.push(row);
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

function gensenReflectAnnualSummaryToCalcSheet() {
  const ss = SpreadsheetApp.getActive();
  const year = gensenGetTargetYear_();
  const summarySheetName = GENSEN_CONFIG.annualSummarySheetPrefix + year;
  const summarySheet = ss.getSheetByName(summarySheetName);
  if (!summarySheet) {
    throw new Error('年次集計シートが見つかりません: ' + summarySheetName);
  }

  const calcSheetName = GENSEN_CONFIG.calcSheetPrefix + year;
  const calcSheet = ss.getSheetByName(calcSheetName);
  if (!calcSheet) {
    throw new Error('計算台シートが見つかりません: ' + calcSheetName);
  }

  const employeeRange = gensenGetNamedRangeOnSheet_(
    ss,
    GENSEN_CONFIG.reflectTargets.employeeIdRangeName,
    calcSheet
  );
  const employeeId = String(employeeRange.getValue() || '').trim();
  if (!employeeId) {
    throw new Error('計算台の従業員IDが取得できません。');
  }

  const values = summarySheet.getDataRange().getValues();
  if (values.length < 2) {
    throw new Error('年次集計にデータがありません。');
  }

  const headers = values[0];
  const headerIndex = gensenBuildHeaderIndex_(headers);
  const requiredHeaders = ['従業員番号', '総支給額', '社会保険', '雇用保険', '源泉所得税'];
  requiredHeaders.forEach((header) => {
    if (headerIndex[header] === undefined) {
      throw new Error('年次集計に必要な列がありません: ' + header);
    }
  });

  const row = values.slice(1).find((dataRow) => String(dataRow[headerIndex['従業員番号']]).trim() === employeeId);
  if (!row) {
    throw new Error('年次集計に従業員番号が見つかりません: ' + employeeId);
  }

  const gross = Number(row[headerIndex['総支給額']]) || 0;
  const socialInsurance = Number(row[headerIndex['社会保険']]) || 0;
  const employmentInsurance = Number(row[headerIndex['雇用保険']]) || 0;
  const withholdingTax = Number(row[headerIndex['源泉所得税']]) || 0;
  const dependents =
    headerIndex['扶養人数'] !== undefined ? Number(row[headerIndex['扶養人数']]) || 0 : null;

  gensenGetNamedRangeOnSheet_(
    ss,
    GENSEN_CONFIG.reflectTargets.grossRangeName,
    calcSheet
  ).setValue(gross);
  gensenGetNamedRangeOnSheet_(
    ss,
    GENSEN_CONFIG.reflectTargets.socialInsuranceRangeName,
    calcSheet
  ).setValue(socialInsurance + employmentInsurance);
  gensenGetNamedRangeOnSheet_(
    ss,
    GENSEN_CONFIG.reflectTargets.withholdingTaxRangeName,
    calcSheet
  ).setValue(withholdingTax);

  if (dependents !== null) {
    const dependentsRange = ss.getRangeByName(GENSEN_CONFIG.reflectTargets.dependentsRangeName);
    if (dependentsRange && dependentsRange.getSheet().getName() === calcSheet.getName()) {
      dependentsRange.setValue(dependents);
    }
  }
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

function gensenGetSheetHeaders_(sheet) {
  if (sheet.getLastRow() < 1 || sheet.getLastColumn() < 1) {
    return [];
  }
  return sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getValues()[0]
    .map((header) => String(header).trim());
}

function gensenGetNamedRangeOnSheet_(ss, rangeName, sheet) {
  const range = ss.getRangeByName(rangeName);
  if (!range) {
    throw new Error('Named Range が見つかりません: ' + rangeName);
  }
  if (range.getSheet().getName() !== sheet.getName()) {
    throw new Error('Named Range の参照先が計算台ではありません: ' + rangeName);
  }
  return range;
}

function gensenExtractYearFromMonth_(monthValue) {
  if (!monthValue) {
    return '';
  }
  if (Object.prototype.toString.call(monthValue) === '[object Date]' && !isNaN(monthValue.getTime())) {
    return String(monthValue.getFullYear());
  }
  const normalized = String(monthValue).trim();
  const match = normalized.match(/^(\d{4})\s*\/\s*\d{1,2}$/);
  if (match) {
    return match[1];
  }
  const numericMatch = normalized.match(/^(\d{4})(\d{2})$/);
  if (numericMatch) {
    return numericMatch[1];
  }
  return '';
}

function gensenParseAnnualSummaryAmount_(value, sheetName, rowNumber, label) {
  const normalized = String(value).trim();
  if (normalized === '') {
    gensenLogAnnualSummaryIssue_(
      sheetName,
      rowNumber,
      label + 'が数値ではありません。',
      value
    );
    return 0;
  }
  const numberValue = Number(value);
  if (Number.isNaN(numberValue)) {
    gensenLogAnnualSummaryIssue_(
      sheetName,
      rowNumber,
      label + 'が数値ではありません。',
      value
    );
    return 0;
  }
  return numberValue;
}

function gensenLogAnnualSummaryIssue_(sheetName, rowNumber, message, value) {
  Logger.log(
    '[年次集計] %s シート=%s 行=%s 値=%s',
    message,
    sheetName,
    rowNumber,
    value
  );
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
