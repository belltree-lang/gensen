/***** メニュー追加 *****/
function buildPayrollMenu_() {
  Logger.log('[buildPayrollMenu_] start');
  SpreadsheetApp.getUi()
    .createMenu('給与明細発行')
    .addItem('給与明細（発行月・行を指定）', 'showMonthAndRowsDialog')
    .addToUi();
  Logger.log('[buildPayrollMenu_] end');
}

/***** 月＋行入力のダイアログ *****/
function showMonthAndRowsDialog() {
  Logger.log('[showMonthAndRowsDialog] start');
  var timezone = Session.getScriptTimeZone();
  var currentYear = parseInt(
    Utilities.formatDate(new Date(), timezone, 'yyyy'),
    10
  );
  var yearOptions = '';
  var year;
  for (year = currentYear - 2; year <= currentYear + 2; year++) {
    yearOptions += '<option value="' + year + '"';
    if (year === currentYear) {
      yearOptions += ' selected';
    }
    yearOptions += '>' + year + '年</option>';
  }
  var monthOptions = '';
  var month;
  for (month = 1; month <= 12; month++) {
    monthOptions += '<option value="' + month + '">' + month + '月</option>';
  }
  var html = `
<html>
  <head>
    <base target="_top">
  </head>
  <body>
    <label>発行年</label><br>
    <select id="year" style="width:100%;padding:6px">
      ${yearOptions}
    </select>
    <br><br>
    <label>発行月</label><br>
    <select id="month" style="width:100%;padding:6px">
      ${monthOptions}
    </select>
    <br><br>
    <label>行番号／範囲（複数可）</label><br>
    <h2>給与明細PDF出力</h2>
    <input id="rows" placeholder="例：5-8,12,20-22" style="width:100%;padding:6px">
    <button id="okBtn" type="button">OK</button>
    <br><small>カンマ区切りで複数指定、ハイフンで範囲指定できます。</small>
    <br><br><button id="cancelBtn" type="button">キャンセル</button>
    <script>
      function submitForm() {
        // Step 3: ボタン取得
        var btn = document.getElementById("okBtn");
        // Step 4: ボタン無効化
        btn.disabled = true;
        btn.innerText = "処理中...";
        // Step 5: 入力値取得
        var yearValue = document.getElementById("year").value;
        var monthValue = document.getElementById("month").value;
        var rowsValue = document.getElementById("rows").value;
        // Step 6: バリデーション
        if (!yearValue) { alert("発行年を選択してください"); btn.disabled = false; btn.innerText = "OK"; return; }
        if (!monthValue) { alert("発行月を選択してください"); btn.disabled = false; btn.innerText = "OK"; return; }
        if (!rowsValue) { alert("行番号を入力してください"); btn.disabled = false; btn.innerText = "OK"; return; }
        // Step 7: サーバー呼び出し（ここでダイアログがフリーズ）
        google.script.run
          .withSuccessHandler(function(msg) {
            if (msg) { alert(msg); }
            google.script.host.close();
          })
          .processMonthAndRows(yearValue, monthValue, rowsValue);
      }
      function bindEvents() {
        var okBtn = document.getElementById("okBtn");
        var cancelBtn = document.getElementById("cancelBtn");
        if (okBtn) {
          okBtn.addEventListener("click", submitForm);
        }
        if (cancelBtn) {
          cancelBtn.addEventListener("click", function() {
            google.script.host.close();
          });
        }
      }
      if (document.readyState === "loading") {
        document.addEventListener("DOMContentLoaded", bindEvents);
      } else {
        bindEvents();
      }
    </script>
  </body>
</html>`;
  var htmlOutput = HtmlService.createHtmlOutput(html)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setWidth(380)
    .setHeight(360);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '給与明細PDF出力');
  Logger.log('[showMonthAndRowsDialog] end');
}

/***** ダイアログから処理を呼び出し *****/
function processMonthAndRows(year, month, rowsInput) {
  Logger.log(`[processMonthAndRows] start year=${year} month=${month} rowsInput=${rowsInput}`);
  const sheetName = `給与明細${month}月支払分`;
  const rowNumbers = parseRowInput(rowsInput);
  if (!rowNumbers.length) return '正しい行番号が入力されていません。';
  const uiYear = parseInt(year, 10);
  const uiMonth = parseInt(month, 10);
  exportRows(sheetName, rowNumbers, uiYear, uiMonth);
  Logger.log(`[processMonthAndRows] end rowNumbers=${rowNumbers.join(',')}`);
  return '指定した行のPDF出力が完了しました！';
}

/***** "5-8,12,20-22" を配列に展開 *****/
function parseRowInput(input) {
  const result = [];
  (input||'').split(',').forEach(part => {
    part = part.trim();
    if (!part) return;
    if (part.includes('-')) {
      const [s,e] = part.split('-').map(v=>parseInt(v,10));
      if (!isNaN(s) && !isNaN(e)) for (let i=s;i<=e;i++) result.push(i);
    } else {
      const n = parseInt(part,10);
      if (!isNaN(n)) result.push(n);
    }
  });
  return Array.from(new Set(result)).sort((a,b)=>a-b);
}

/***** 429対応：リトライ付きの fetch *****/
function fetchWithRetry(url, token, retries = 3) {
  for (let i=0; i<retries; i++) {
    try {
      return UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + token } });
    } catch (e) {
      if (e.message.includes('429') && i < retries-1) {
        Utilities.sleep(5000); // 5秒待ってリトライ
        continue;
      }
      throw e;
    }
  }
}

const PAYROLL_TEMPLATE_CONFIG = {
  issueDateCell: 'B2',
  targetYmCell: 'B3'
};

function getTemplateRange_(ss, templateSheet, rangeName, fallbackA1) {
  const namedRange = ss.getRangeByName(rangeName);
  if (namedRange) {
    return namedRange;
  }
  if (!fallbackA1) {
    return null;
  }
  return templateSheet.getRange(fallbackA1);
}

/***** 実際のPDF出力処理（ここは既存の exportRows の内容を流用） *****/
function exportRows(sheetName, rowNumbers, uiYear, uiMonth) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) { SpreadsheetApp.getUi().alert(`シート「${sheetName}」が見つかりません。`); return; }

  const template       = ss.getSheetByName('給与明細テンプレート');
  const masterTemplate = ss.getSheetByName('給与明細テンプレート元式');
  const folder         = DriveApp.getFolderById('1Jw_QcZ1ph_mi92Y5I2efvpg15VivyV1X');
  const spreadsheetId  = ss.getId();

  const dataRange      = masterTemplate.getDataRange();
  const numRows        = dataRange.getNumRows();
  const numCols        = dataRange.getNumColumns();
  const masterFormulas = dataRange.getFormulas();
  const masterValues   = dataRange.getValues();
  const token          = ScriptApp.getOAuthToken();

  rowNumbers.forEach(rowNum => {
    const name = sheet.getRange(rowNum, 4).getValue();
    if (!name) return;

    if (!uiYear || !uiMonth || uiMonth < 1 || uiMonth > 12 || isNaN(uiYear) || isNaN(uiMonth)) {
      SpreadsheetApp.getUi().alert(`【${rowNum}行目／${name}さん：年月が不正です】発行年・発行月を確認してください。`);
      return;
    }

    // ==== 数式＆ラベル反映 ====
    for (let r = 0; r < numRows; r++) {
      for (let c = 0; c < numCols; c++) {
        const formula = masterFormulas[r][c];
        if (formula && /'給与明細.*?支払分'!/g.test(formula)) {
          const newFormula = formula.replace(/='?給与明細.*?支払分'?!([A-Z]+)\d+/g, `='${sheetName}'!$1${rowNum}`);
          template.getRange(r + 1, c + 1).setFormula(newFormula);
        } else {
          template.getRange(r + 1, c + 1).setValue(masterValues[r][c]);
        }
      }
    }
    SpreadsheetApp.flush();

    // ==== 対象年月・発行日 ====
    const issueDate = new Date(uiYear, uiMonth, 0);
    const targetYm = `${uiYear}年${uiMonth}月分`;
    const issueDateRange = getTemplateRange_(ss, template, 'PAYROLL_ISSUE_DATE', PAYROLL_TEMPLATE_CONFIG.issueDateCell);
    const targetYmRange = getTemplateRange_(ss, template, 'PAYROLL_TARGET_YM', PAYROLL_TEMPLATE_CONFIG.targetYmCell);
    if (!issueDateRange || !targetYmRange) {
      SpreadsheetApp.getUi().alert('給与明細テンプレートの発行日/対象年月セルが見つかりません。');
      return;
    }
    issueDateRange.setValue(issueDate);
    targetYmRange.setValue(targetYm);

    Logger.log(`[exportRows] row=${rowNum} name=${name} uiYear=${uiYear} uiMonth=${uiMonth} issueDate=${issueDate} targetYm=${targetYm}`);

    // ==== ファイル名生成 ====
    const reiwaYear = uiYear - 2018;
    const safeName = String(name).replace(/[\/\\:*?"<>|]/g, '');
    const fileName = `${safeName}_給与支払明細_令和${reiwaYear}年${uiMonth}月分`;

    const folderName = `${safeName}殿`;
    const folders = folder.getFoldersByName(folderName);
    const subFolder = folders.hasNext() ? folders.next() : folder.createFolder(folderName);

    const exportUrl =
      `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?exportFormat=pdf&format=pdf&gid=${template.getSheetId()}&size=A4&portrait=true&fitw=true&sheetnames=false&printtitle=false&pagenumbers=false&gridlines=false&fzr=false`;

    try {
      const res = fetchWithRetry(exportUrl, token);
      subFolder.createFile(res.getBlob()).setName(fileName + '.pdf');
    } catch (err) {
      SpreadsheetApp.getUi().alert(`【${name}さん分でエラー】\n${err}`);
    }

    Utilities.sleep(2000); // 負荷対策（2秒に延長）
  });
}
