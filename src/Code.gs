// === 固定ラベル ===
const FIXED_HISTORY_LABEL = "Ⓒ1か月以内ニーズ有 × 選ばなそう";


//================ onEdit ================
// onEdit が反応するシート（それ以外の編集はすぐ終わる）
const ONEDIT_SHEETS_ = new Set([
  "掲載状況", "掲載開始顧客", "新規ホット", "飛び込み先リスト", "メール管理",
  "営業先リスト/新規", "営業先リスト/現S", "会話履歴/新規", "会話履歴/現S"
]);

// 営業先リスト E列入力のたびに S列へ順番に出すトーク
const TALK_PHRASES_ = [
  "「最近“新人どう育てるか”って話題多くて…御社ではどうされてます？」",
  "「なんとなく現場の空気変わってきた気しません？御社って何か感じてます？」",
  "「みなさん結構“判断迷ってること”あるみたいで…何か今、引っかかってることあります？」",
  "「今後どうしていくか、結構みんな探ってる感じで…御社はどんな展望あります？」",
  "「他社さん、今かなり動いてるみたいなんですけど…御社は落ち着いてます？」",
  "「最近よく聞くのが“他社の動き”でして…ざっくりどんな感じかって耳にされたりしてます？」",
  "「この前話してたことなんですけど、これってもう古いのかなって…御社だとどうです？」",
  "「最近“社内でよく聞く言葉”ってあります？地味にキーワードで空気変わりますよね」",
  "「そういえば最近“業界の噂”とか何か耳にされました？結構情報錯綜してて…」",
  "「最近“あと一歩で決まりそうだった案件”とかありました？その温度感、すごく大事で…」"
];

function onEdit(e) {
  try {
    const sheet = e.range.getSheet();
    const sheetName = sheet.getName();
    if (!ONEDIT_SHEETS_.has(sheetName)) return;

    const ss = e.source;
    const tz = Session.getScriptTimeZone();
    const now = new Date();
    const ymd = Utilities.formatDate(now, tz, "yyyy/MM/dd");
    const row = e.range.getRow();
    const col = e.range.getColumn();
    const value = e.range.getValue();
    const filled = !(value === "" || value === 0 || value === null);
    const checked = value === true || e.value === "TRUE";
    const memo = () => `${ymd} ${Utilities.formatDate(now, tz, "HH時")} - ${value}`;
    const clearCell = () => sheet.getRange(row, col).clearContent();
    const countUp = (logName, c) => {
      const log = ss.getSheetByName(logName);
      if (log) updateCallLog(log, ymd, c);
    };

    switch (sheetName) {

      case "掲載状況": {
        //=== J(10)入力 → 会話履歴/現S 追記＋営業先/現S I更新＋架電/現S B+1＋セルクリア
        if (col === 10 && filled) {
          const company = sheet.getRange(row, 2).getValue();
          const hist = ss.getSheetByName("会話履歴/現S");
          if (hist && company) appendHistory_(hist, company, memo(), FIXED_HISTORY_LABEL);
          const sales = ss.getSheetByName("営業先リスト/現S");
          const r = sales && company ? findRowInColA_(sales, company) : 0;
          if (r) sales.getRange(r, 9).setValue(ymd); // I列
          countUp("架電記録/現S", 2); // B列+1
          clearCell();
          return;
        }
        //=== G/H/I チェック → 同じ社名の行へ同期（押した列だけ）＋ H=TRUE で行灰色（条件付書式）
        if (row > 1 && col >= 7 && col <= 9) {
          const keyB = sheet.getRange(row, 2).getDisplayValue(); // B：社名
          if (!keyB) return;
          const n = sheet.getLastRow() - 1;
          if (n <= 0) return;
          const names = sheet.getRange(2, 2, n, 1).getDisplayValues();
          const rng = sheet.getRange(2, col, n, 1);
          const vals = rng.getValues();
          names.forEach((b, i) => { if (b[0] === keyB) vals[i][0] = checked; });
          rng.setValues(vals);

          if (col === 8) {
            ensureGreyRuleForKeisai_(sheet);
            if (checked) {
              setGenSLastValidDate_(ss, [sheet.getRange(row, 1).getDisplayValue(), keyB], ymd);
              countUp("架電記録/現S", 3); // C列+1
            }
          }
        }
        return;
      }

      case "掲載開始顧客": {
        //=== K(11)チェック → NGリスト追記＋行灰色
        if (col === 11) {
          if (!checked || row === 1) return;
          const name = sheet.getRange(row, 1).getDisplayValue();
          if (!name) return;
          const ng = ss.getSheetByName("NGリスト") || ss.insertSheet("NGリスト");
          if (ng.getLastRow() === 0) ng.getRange(1, 1).setValue("NGリスト");
          ng.appendRow([name]);
          paintRow_(sheet, row, "#D3D3D3");
          return;
        }
        if (row === 1) return;

        //=== J(10)入力 → Lへ蓄積＋会話履歴/新規 追記＋架電/新規 B+1＋Jクリア＋行薄紫
        if (col === 10 && filled) {
          const company = sheet.getRange(row, 1).getDisplayValue();
          if (!company) return;
          const hist = ss.getSheetByName("会話履歴/新規");
          const log = ss.getSheetByName("架電記録/新規");
          if (!hist || !log) return;
          const text = memo();
          const lCell = sheet.getRange(row, 12);
          const prevL = String(lCell.getValue() || "").trim();
          lCell.setValue(prevL ? prevL + "\n" + text : text).setWrap(true);
          appendHistory_(hist, company, text, FIXED_HISTORY_LABEL);
          updateCallLog(log, ymd, 2); // B+1
          clearCell();
          paintRow_(sheet, row, "#E6E6FA");
        }
        //=== F(6)チェック → TRUE:緑 / FALSE:薄紫
        else if (col === 6) setRowColorByFlag_(sheet, row, sheet.getLastColumn(), checked);
        //=== G(7)チェック → 架電記録/新規 C列+1
        else if (col === 7 && checked) countUp("架電記録/新規", 3);
        //=== H(8)チェック → 架電記録/新規 D列+1（チェックを外したときは数えない）
        else if (col === 8 && checked) countUp("架電記録/新規", 4);
        //=== I(9)入力 → 営業先リスト/新規へ転記＋会話履歴へ追記＋相互リンク＋行グレー化
        else if (col === 9 && value !== "") upsertShinkiAndHistoryFromKeisai_(ss, sheet, row, ymd, tz);
        return;
      }

      case "新規ホット": {
        //=== H(8) / I(9)チェック → 同じ社名の行を色塗り＋営業先リスト/新規の日付更新
        if (row > 1 && (col === 8 || col === 9)) {
          if (!checked) return;
          const keyA = String(sheet.getRange(row, 1).getDisplayValue() || "").trim();
          if (!keyA) return;
          const n = sheet.getLastRow() - 1;
          if (n <= 0) return;
          const names = sheet.getRange(2, 1, n, 1).getDisplayValues();
          const rng = sheet.getRange(2, 1, n, sheet.getLastColumn());
          const bgs = rng.getBackgrounds();
          names.forEach((a, i) => {
            if (String(a[0]).trim() !== keyA) return;
            if (col === 9) bgs[i].fill("#D3D3D3");
            else if (!bgs[i].some(c => String(c || "").toUpperCase() === "#D3D3D3")) bgs[i].fill("#E6E6FA");
          });
          rng.setBackgrounds(bgs);

          const sales = ss.getSheetByName("営業先リスト/新規");
          const r = sales ? findRowInColA_(sales, keyA) : 0;
          if (r) sales.getRange(r, 9, 1, col === 9 ? 2 : 1).setValue(ymd); // Hチェック→I列 / Iチェック→I列とJ列
          // Hチェック＝今日電話した → 記録して営業先リスト/新規の行を薄紫（G列に値があれば黄緑）に
          if (col === 8) {
            markCalledToday_(keyA, ymd);
            if (r) updateRowColor(sales, r, sales.getLastColumn());
          }
          return;
        }
        //=== G(7)チェック → TRUE:緑 / FALSE:薄紫
        if (col === 7 && row > 1) {
          setRowColorByFlag_(sheet, row, sheet.getLastColumn(), checked);
          return;
        }
        //=== J(10)入力 → 会話履歴/新規へ追記（B固定なし）＋架電記録/新規 B+1＋営業先リスト/新規 I更新＋セルクリア＋行薄紫
        if (col === 10 && filled) {
          const company = String(sheet.getRange(row, 1).getDisplayValue() || "").trim();
          if (!company) return;
          const hist = ss.getSheetByName("会話履歴/新規");
          if (hist) appendHistory_(hist, company, memo(), "");
          countUp("架電記録/新規", 2); // B+1
          markCalledToday_(company, ymd); // 営業先リスト/新規で薄紫にする
          const sales = ss.getSheetByName("営業先リスト/新規");
          const r = sales ? findRowInColA_(sales, company) : 0;
          if (r) sales.getRange(r, 9).setValue(ymd); // I列
          clearCell();
          paintRow_(sheet, row, "#E6E6FA");
        }
        return;
      }

      case "飛び込み先リスト": {
        //=== K(11)入力でMM/ddセット
        if (col === 11 && filled) sheet.getRange(row, 11).setValue(Utilities.formatDate(now, tz, "MM/dd"));
        return;
      }

      case "メール管理": {
        //=== C入力 → 日付＋行色変更＋該当の架電記録 B/C+1
        if (col !== 3 || row === 1 || !filled) return;
        const cell = sheet.getRange(row, 3);
        cell.setValue(Utilities.formatDate(now, tz, "MM/dd"));
        const bgHex = (cell.getBackground() || "").toUpperCase();
        if (bgHex === "#DC143C") paintRow_(sheet, row, "#FFC0CB");
        else if (bgHex === "#5B9BD5") paintRow_(sheet, row, "#ADD8E6");
        const logName = { "#FFC0CB": "架電記録/新規", "#DC143C": "架電記録/新規", "#ADD8E6": "架電記録/現S", "#5B9BD5": "架電記録/現S" }[bgHex];
        if (logName) {
          countUp(logName, 2); // B+1
          countUp(logName, 3); // C+1
        }
        return;
      }

      case "営業先リスト/新規":
      case "営業先リスト/現S": {
        if (row === 1 || (!filled && col !== 7)) return;
        const isNew = sheetName === "営業先リスト/新規";
        const logSheet = ss.getSheetByName(isNew ? "架電記録/新規" : "架電記録/現S");
        const histName = isNew ? "会話履歴/新規" : "会話履歴/現S";
        // 新規：E/F/H 列の入力＝今日電話した → 記録して行を薄紫（G列に値があれば黄緑）に
        const markCalled = () => {
          if (!isNew) return;
          markCalledToday_(sheet.getRange(row, 1).getDisplayValue(), ymd);
          updateRowColor(sheet, row, sheet.getLastColumn());
        };

        switch (col) {
          case 7: { // G列：予定の有無で行色更新
            if (!isNew) {
              updateRowColor(sheet, row, sheet.getLastColumn());
              return;
            }
            // 新規：G列に値あり＝黄緑、空で今日電話済み＝薄紫、空で未電話＝黄緑だったときだけ白に戻す
            const name = String(sheet.getRange(row, 1).getDisplayValue() || "").trim();
            if (filled || getCalledToday_(ymd).has(name)) updateRowColor(sheet, row, sheet.getLastColumn());
            else {
              const rng = sheet.getRange(row, 1, 1, sheet.getLastColumn());
              if (String(rng.getBackground() || "").toUpperCase() === "#ADFF2F") rng.setBackground(null);
            }
            return;
          }

          case 6: // F列：J に日付＋架電 C+1＋クリア（新規は L もクリア）
            if (isNew) sheet.getRange(row, 12).clearContent();
            sheet.getRange(row, 10).setValue(ymd); // J
            if (logSheet) updateCallLog(logSheet, ymd, 3); // C+1
            sheet.getRange(row, 6).clearContent();
            markCalled();
            return;

          case 5: { // E列：会話追記＋I更新＋B+1＋B1回転＋（新規）時間帯カウンタ＋Sローテ
            const hist = ss.getSheetByName(histName);
            if (!hist) return;
            const company = sheet.getRange(row, 1).getValue();
            const newHistRow = appendHistory_(hist, company, memo(), FIXED_HISTORY_LABEL);
            if (newHistRow) createHyperlinks(sheet, hist, row, newHistRow, company);

            sheet.getRange(row, 9).setValue(ymd); // I：最終接触日
            if (isNew) markCalledToday_(company, ymd);
            updateRowColor(sheet, row, sheet.getLastColumn());
            sheet.getRange(row, 5).clearContent();
            if (logSheet) updateCallLog(logSheet, ymd, 2); // B+1

            const countCell = sheet.getRange("B1");
            countCell.setValue((Number(countCell.getValue() || 0) % 10) + 1);

            if (isNew) {
              const statusCell = sheet.getRange(row, 14);
              const status0 = getOrInitStatus_(statusCell.getValue());
              const hour = now.getHours();
              const label = hour < 12 ? "午前中" : hour < 15 ? "13-15時" : hour < 17 ? "15-17時" : "";
              statusCell.setValue(label ? bumpCounterText_(status0, label) : status0);
            }

            const sCell = sheet.getRange(row, 19);
            const idx = TALK_PHRASES_.indexOf(sCell.getValue());
            sCell.setValue(TALK_PHRASES_[idx < 0 ? 0 : (idx + 1) % TALK_PHRASES_.length]);
            return;
          }

          case 8: { // H列：架電 D+1＋アポイント回数+1＋クリア
            if (logSheet) updateCallLog(logSheet, ymd, 4);
            const statusCell = sheet.getRange(row, 14);
            statusCell.setValue(bumpCounterText_(getOrInitStatus_(statusCell.getValue()), "アポイント"));
            sheet.getRange(row, 8).clearContent();
            markCalled();
            return;
          }

          case 13: { // M列：今Qヨミへ行コピー＆最新会話をEへ
            const yomi = ss.getSheetByName("今Qヨミ");
            if (yomi) {
              const targetRow = yomi.getLastRow() + 1;
              sheet.getRange(row, 1, 1, sheet.getLastColumn()).copyTo(yomi.getRange(targetRow, 1), { contentsOnly: false });
              const company = sheet.getRange(row, 1).getValue();
              yomi.getRange(targetRow, 5).setValue(getLatestConversation(ss.getSheetByName(histName), company));
            }
            sheet.getRange(row, 13).clearContent();
            return;
          }
        }
        return;
      }

      case "会話履歴/新規":
      case "会話履歴/現S": {
        //=== B空なら黄、入力でクリア
        if (col === 2) e.range.setBackground(value !== "" ? null : "#FFFF00");
        return;
      }
    }
  } catch (error) {
    console.error("onEdit エラー:", error);
  }
}

// A列で社名が完全一致する最初の行番号（なければ0）
function findRowInColA_(sheet, text) {
  const hit = sheet.getRange("A:A").createTextFinder(String(text))
    .matchCase(true).matchEntireCell(true).findNext();
  return hit ? hit.getRow() : 0;
}

// 会話履歴のC列へ1行追記。label があればB列にも書く。新しい行を作ったときだけその行番号を返す
function appendHistory_(hist, company, text, label) {
  const r = findRowInColA_(hist, company);
  if (r) {
    const c = hist.getRange(r, 3);
    const cur = c.getValue();
    c.setValue((cur ? cur + "\n" : "") + text);
    if (label) hist.getRange(r, 2).setValue(label);
    return 0;
  }
  hist.appendRow([company, label || "", text]);
  return hist.getLastRow();
}

// 行全体（最終列まで）の背景色を変える
function paintRow_(sheet, row, color) {
  sheet.getRange(row, 1, 1, sheet.getLastColumn()).setBackground(color);
}


//================ 共通関数 ================
// 営業先リスト/現S のA列（大文字小文字・前後空白は無視）で名前を順に探し、最初に見つかった行のJ列へ日付
function setGenSLastValidDate_(ss, names, ymd) {
  const sh = ss.getSheetByName("営業先リスト/現S");
  if (!sh || sh.getLastRow() < 1) return;
  const norm = s => String(s || "").trim().toLowerCase();
  const colA = sh.getRange(1, 1, sh.getLastRow(), 1).getDisplayValues().map(r => norm(r[0]));
  for (const name of names) {
    const key = norm(name);
    const i = key ? colA.indexOf(key) : -1;
    if (i >= 0) {
      sh.getRange(i + 1, 10).setValue(ymd);
      return;
    }
  }
}

function updateRowColor(sheet, row, lastCol) {
  const gVal = sheet.getRange(row, 7).getValue();
  const color = (!gVal || String(gVal).trim() === "") ? "#E6E6FA" : "#ADFF2F";
  const rng = sheet.getRange(row, 1, 1, lastCol);
  const cur = (rng.getBackground() || "").toUpperCase();
  const tgt = color.toUpperCase();
  if (cur !== tgt) rng.setBackground(color);
}

function ensureGreyRuleForKeisai_(sheet) {
  const sp = PropertiesService.getScriptProperties();
  const flagKey = "greyRuleSet:" + sheet.getSheetId();
  if (sp.getProperty(flagKey) === "1") return;

  const rangeA2Z = sheet.getRange(2, 1, Math.max(sheet.getMaxRows() - 1, 1), sheet.getMaxColumns());
  const rules = sheet.getConditionalFormatRules() || [];
  const formula = '=$H2=TRUE';

  const bg = SpreadsheetApp.newConditionalFormatRule()
    .setRanges([rangeA2Z])
    .whenFormulaSatisfied(formula)
    .setBackground("#D3D3D3")
    .build();

  const exists = rules.some(r => {
    try {
      const bc = r.getBooleanCondition && r.getBooleanCondition();
      return !!bc && String(bc.getCriteriaType()) === 'CUSTOM_FORMULA' && String(bc.getCriteriaValues()[0]) === formula;
    } catch (_) {
      return false;
    }
  });

  if (!exists) {
    rules.unshift(bg);
    sheet.setConditionalFormatRules(rules);
  }
  sp.setProperty(flagKey, "1");
}

function setRowColorByFlag_(sheet, row, lastCol, flag) {
  const color = (flag === true || flag === "TRUE" || flag === 1) ? "#ADFF2F" : "#E6E6FA";
  const rng = sheet.getRange(row, 1, 1, lastCol);
  const cur = (rng.getBackground() || "").toUpperCase();
  const tgt = color.toUpperCase();
  if (cur !== tgt) rng.setBackground(color);
}

//================ ステータス（N列）管理：初期化 & インクリメント =================
function getOrInitStatus_(val) {
  const def = "午前中/0回 13-15時/0回 15-17時/0回 アポイント/0回";
  const s = String(val || "").trim();
  if (!s) return def;

  const need = ["午前中", "13-15時", "15-17時", "アポイント"];
  let out = s;

  need.forEach(label => {
    if (!new RegExp(label + "\\/[0-9]+回").test(out)) {
      out = (out ? (out + " ") : "") + `${label}/0回`;
    }
  });

  return out;
}

function bumpCounterText_(statusText, label) {
  const parts = String(statusText || "").trim().split(/\s+/).filter(Boolean);
  let found = false;

  for (let i = 0; i < parts.length; i++) {
    const m = parts[i].match(new RegExp(`^(${label})\\/(\\d+)回$`));
    if (m) {
      const n = parseInt(m[2], 10) + 1;
      parts[i] = `${label}/${n}回`;
      found = true;
      break;
    }
  }

  if (!found) parts.push(`${label}/1回`);
  return parts.join(" ");
}

//================ 掲載開始顧客 → 営業先/新規 へ転記＋履歴追記＋相互リンク =================
function upsertShinkiAndHistoryFromKeisai_(ss, keisaiSheet, row, ymd, tz) {
  const src = keisaiSheet.getRange(row, 1, 1, 9).getDisplayValues()[0];

  const company = String(src[0] || "").trim();
  if (!company) return;

  const phone = src[1] || "";
  const address = src[2] || "";
  const keyman = String(src[8] || "").trim();
  const media = [src[3], src[4]]
    .map(v => String(v || "").trim())
    .filter(Boolean)
    .join("\n");

  const newList = ss.getSheetByName("営業先リスト/新規");
  const hist = ss.getSheetByName("会話履歴/新規");
  if (!newList || !hist) return;

  const targetRow = findRowInColA_(newList, company) || newList.getLastRow() + 1;

  // 掲載開始顧客 A/B/C/I → 営業先リスト/新規 A/B/C/D
  newList.getRange(targetRow, 1, 1, 4).setValues([[company, phone, address, keyman]]);

  // 転記実行日を I（最終TEL）・J（最終有効）へ日付型で記録
  const transferTz = ss.getSpreadsheetTimeZone();
  const transferDate = Utilities.parseDate(Utilities.formatDate(new Date(), transferTz, "yyyy/MM/dd"), transferTz, "yyyy/MM/dd");
  newList.getRange(targetRow, 9, 1, 2)
    .setValues([[transferDate, transferDate]])
    .setNumberFormat("MM/dd");

  // 掲載開始顧客 D/E（媒体）→ 営業先リスト/新規 M（掲載状況）
  newList.getRange(targetRow, 13).setValue(media);

  const nCell = newList.getRange(targetRow, 14);
  nCell.setValue(getOrInitStatus_(nCell.getValue()));
  newList.getRange(targetRow, 22).setValue(ymd);

  const text = `${ymd} ${Utilities.formatDate(new Date(), tz, "HH時")} - ${keyman || "掲載開始顧客から転記"}`;
  appendHistory_(hist, company, text, FIXED_HISTORY_LABEL);

  // 同名が複数あるときは一番下の行とリンクする
  const hits = hist.getRange("A:A").createTextFinder(company).matchCase(true).matchEntireCell(true).findAll();
  if (hits && hits.length > 0) {
    const newHistRow = Math.max.apply(null, hits.map(r => r.getRow()));
    createHyperlinks(newList, hist, targetRow, newHistRow, company);
  }

  paintRow_(keisaiSheet, row, "#D3D3D3");
}


//================ 初期化/ユーティリティ ================
// ============================
// ナビゲーション用関数
// ============================
function showNowQYomi() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("今Qヨミ");
  if (sheet && sheet.isSheetHidden()) sheet.showSheet();
  if (sheet) ss.setActiveSheet(sheet);
}

function showNGWord() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("NGワード");
  if (sheet && sheet.isSheetHidden()) sheet.showSheet();
  if (sheet) ss.setActiveSheet(sheet);
}

function showTaskManagement() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("タスク管理");
  if (sheet && sheet.isSheetHidden()) sheet.showSheet();
  if (sheet) ss.setActiveSheet(sheet);
}

// ============================
// 分割後の統合実行関数
// ============================
function runAllProcessesCombined() {
  runAllProcesses_Phase1();
  runAllProcesses_Phase2();
  runAllProcesses_Phase3();
}

// ============================
// 高速版：NGワード＆重複チェック（全体一括で背景反映）
// ============================
function highlightRows() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ngSheet = ss.getSheetByName("NGワード");
  var sheetNames = ["営業先リスト/新規", "営業先リスト/現S"];

  if (!ngSheet) {
    Logger.log("NGワードシートが見つかりません。");
    return;
  }

  var ngLastRow = ngSheet.getLastRow();
  if (ngLastRow < 1) {
    Logger.log("NGワードがありません。");
    return;
  }

  var ngWords = ngSheet.getRange(1, 1, ngLastRow, 1)
    .getValues()
    .map(function (r) { return String(r[0]).trim(); })
    .filter(function (w) { return w !== ""; });

  var sheets = [];
  var telMap = Object.create(null);
  var compMap = Object.create(null);

  for (var s = 0; s < sheetNames.length; s++) {
    var sh = ss.getSheetByName(sheetNames[s]);
    if (!sh) {
      sheets.push(null);
      continue;
    }

    var lastRow = sh.getLastRow();
    var lastCol = sh.getLastColumn();
    if (lastRow < 2 || lastCol < 1) {
      sheets.push({ sh: sh, data: [], bgs: [], lastRow: lastRow, lastCol: lastCol });
      continue;
    }

    var rng = sh.getRange(1, 1, lastRow, lastCol);
    var values = rng.getValues();
    var numCols = values[0].length;

    var bgs = new Array(lastRow);
    for (var i = 0; i < lastRow; i++) {
      var rowBg = new Array(numCols);
      for (var j = 0; j < numCols; j++) rowBg[j] = null;
      bgs[i] = rowBg;
    }

    for (var r = 1; r < lastRow; r++) {
      var companyName = String(values[r][0] || "").trim();
      var telephone = String(values[r][1] || "").trim();

      if (companyName && ngWords.some(function (ng) { return companyName.indexOf(ng) !== -1; })) {
        for (var c = 0; c < numCols; c++) bgs[r][c] = "#FFC0CB";
      }

      if (telephone) {
        (telMap[telephone] || (telMap[telephone] = [])).push({
          sidx: s,
          rIdx: r,
          numCols: numCols
        });
      }

      if (companyName) {
        (compMap[companyName] || (compMap[companyName] = [])).push({
          sidx: s,
          rIdx: r,
          numCols: numCols
        });
      }
    }

    sheets.push({ sh: sh, data: values, bgs: bgs, lastRow: lastRow, lastCol: lastCol });
  }

  function paintDup(map) {
    for (var key in map) {
      var arr = map[key];
      if (arr && arr.length > 1) {
        for (var k = 0; k < arr.length; k++) {
          var info = arr[k];
          var pack = sheets[info.sidx];
          if (!pack) continue;
          for (var c = 0; c < info.numCols; c++) {
            pack.bgs[info.rIdx][c] = "#FF0000";
          }
        }
      }
    }
  }

  paintDup(telMap);
  paintDup(compMap);

  sheets.forEach(function (pack) {
    if (!pack || !pack.sh || !pack.bgs) return;
    var sh = pack.sh;
    var lastRow = pack.lastRow;
    var lastCol = pack.lastCol;
    if (lastRow < 2) return;

    var dataRange = sh.getRange(1, 1, lastRow, lastCol);
    dataRange.setBackground(null);
    dataRange.setBackgrounds(pack.bgs);
  });

  SpreadsheetApp.getUi().alert("NGワードと重複チェックが高速版で完了しました。");
}

// ============================
// 会話履歴：指定社名の最新の会話（C列）を取得
// ============================
function getLatestConversation(convSheet, companyName) {
  if (!convSheet) return "";

  var foundCells = convSheet.getRange("A:A")
    .createTextFinder(companyName)
    .matchCase(true)
    .matchEntireCell(true)
    .findAll();

  if (!foundCells || foundCells.length === 0) return "";

  var maxRow = Math.max.apply(null, foundCells.map(function (cell) {
    return cell.getRow();
  }));
  var fullConversation = convSheet.getRange(maxRow, 3).getValue();
  return fullConversation ? String(fullConversation).split("\n").pop() : "";
}

// ============================
// 相互リンク：営業先と他シートの間でリンク作成
// ============================
function createHyperlinks(sourceSheet, targetSheet, sourceRow, targetRow, companyName) {
  const name = escapeForFormula_(companyName);
  const link = (toSheet, toRow) =>
    `=IFERROR(HYPERLINK("#gid=${toSheet.getSheetId()}&range=A${toRow}", "${name}"), "${name}")`;
  sourceSheet.getRange(sourceRow, 1).setFormula(link(targetSheet, targetRow));
  targetSheet.getRange(targetRow, 1).setFormula(link(sourceSheet, sourceRow));
}

function deleteABCD_andLeftShift() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getActiveSheet();
  dropFirstNColsByLeftShift_(sh, 1, 4);
  SpreadsheetApp.getUi().alert("A〜D列を削除して左詰めしました。");
}

// 既存のトリガーを削除
function 危険_全トリガー削除() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) ScriptApp.deleteTrigger(trigger);
}

// リセット→本処理の順に安全実行（ログ付き）
function 実行_リセット後に全処理() {
  Logger.log("▶️ リセット → 全処理 実行 開始");

  try {
    Logger.log("🟡 リセットカラー処理 開始");
    updateAllSheetsAndResetColors_Safe();
    Logger.log("🟢 リセットカラー処理 完了");
  } catch (e) {
    Logger.log("❌ リセットカラー処理 エラー: " + e.message);
    throw e;
  }

  try {
    Logger.log("🟡 本処理 実行 開始");
    runAllProcessesCombined();
    Logger.log("🟢 本処理 実行 完了");
  } catch (e) {
    Logger.log("❌ 本処理 エラー: " + e.message);
    throw e;
  }

  Logger.log("✅ リセット → 全処理 実行 完了");
}

// ============================
// 列左詰め：先頭 n 列を落としてシフト
// ＋ A:B結合解除
// ＋ 詰めた後に B列とC列を入れ替え
// ＋ C:D列の入力規則を削除
// ＋ メイリオ太字10ptに統一
// ============================
function dropFirstNColsByLeftShift_(sh, headerRows, nDrop) {
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();

  if (lastRow <= headerRows) return;

  // A:B の結合解除
  try {
    sh.getRange(headerRows + 1, 1, lastRow - headerRows, 2).breakApart();
  } catch (e) {}

  if (lastCol <= nDrop) {
    sh.getRange(headerRows + 1, 1, lastRow - headerRows, lastCol).clearContent();
    return;
  }

  const numRows = lastRow - headerRows;

  // E列以降を取得
  const data = sh
    .getRange(headerRows + 1, nDrop + 1, numRows, lastCol - nDrop)
    .getValues();

  // 左詰め後の B/C 入れ替え
  for (let i = 0; i < data.length; i++) {
    if (data[i].length >= 3) {
      const temp = data[i][1]; // B
      data[i][1] = data[i][2]; // C → B
      data[i][2] = temp;       // B → C
    }
  }

  // 一旦全消し
  sh.getRange(headerRows + 1, 1, numRows, lastCol).clearContent();

  // 書き戻し
  const writeRange = sh.getRange(
    headerRows + 1,
    1,
    numRows,
    data[0].length
  );

  writeRange.setValues(data);

  // C:D列の入力規則を削除
  const newLastRow = sh.getLastRow();

  if (newLastRow > headerRows) {
    sh.getRange(headerRows + 1, 3, newLastRow - headerRows, 2)
      .clearDataValidations();
  }

  // メイリオ・太字・10pt
  writeRange
    .setFontFamily('Meiryo')
    .setFontWeight('bold')
    .setFontSize(10);
}

function processPhoneNumbers() {
  var sheetName = "営業先リスト/新規";
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    SpreadsheetApp.getUi().alert("シート「" + sheetName + "」が見つかりません。");
    return;
  }

  // 1行目はヘッダーとし、B列（2列目）の2行目以降を対象
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  var range = sheet.getRange(2, 2, lastRow - 1, 1);
  var values = range.getValues();

  // ループ内で使う関数呼び出しのオーバーヘッドを避けるため、事前にマップ関数を取得
  for (var i = 0; i < values.length; i++) {
    var original = values[i][0];
    if (original) {
      var originalStr = String(original);
      // セル内に「/」がある場合は分割して各番号を整形
      if (originalStr.indexOf('/') !== -1) {
        var parts = originalStr.split('/');
        for (var j = 0; j < parts.length; j++) {
          parts[j] = formatPhoneNumber(parts[j].trim());
        }
        values[i][0] = parts.join('/');
      } else {
        // 区切りがなくても、複数の番号が連結されている可能性があるので対応
        values[i][0] = formatPhoneNumber(originalStr);
      }
    }
  }

  range.setValues(values);
}


/***** C列（住所）整形―営業先リスト/新規  v3 *****/

function normalizeAddress_inNewProspectSheet() {
  const SHEET_NAME = '営業先リスト/新規';
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) {
    SpreadsheetApp.getUi().alert(`シート「${SHEET_NAME}」が見つかりません。`);
    return;
  }
  const numRows = sheet.getLastRow() - 1; // 1行目は見出し
  if (numRows <= 0) return;

  const range = sheet.getRange(2, 3, numRows, 1); // C列
  range.setValues(range.getValues().map(r => [cleanAddress(String(r[0]))]));
  Logger.log('住所を整形した行数: ' + numRows);
}

/***************
 * 実行ハブ & 共通ユーティリティ（高速化対応版）
 ***************/

function getSheetOrAlert(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    writeExecLog_(`シート「${sheetName}」が見つかりません`, 'ERROR');
    return null;
  }
  return sheet;
}

// ---- ログ系ユーティリティ ----
function logTime(label, time) {
  Logger.log("🔹 " + label + " - " + Utilities.formatDate(time, Session.getScriptTimeZone(), "HH:mm:ss"));
}
function logProcessTime(name, func) {
  const start = new Date();
  logTime(name + " 開始", start);
  try { func(); }
  catch (e) { Logger.log("⚠️ " + name + " 実行中にエラー発生: " + e.message); }
  const end = new Date();
  Logger.log("✅ " + name + " 終了 所要時間: " + ((end - start) / 1000) + "秒");
}

// ---- 軽量な差分書き込みユーティリティ ----
function setFormulasIfChanged_(range, newFormulas) {
  const old = range.getFormulas();
  if (!arraysEqual2D_(old, newFormulas)) range.setFormulas(newFormulas);
}
function arraysEqual2D_(a, b) {
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i++) {
    const r1 = a[i], r2 = b[i];
    if (!r1 || !r2 || r1.length !== r2.length) return false;
    for (let j = 0; j < r1.length; j++) {
      if (String(r1[j]) !== String(r2[j])) return false;
    }
  }
  return true;
}

function runAllProcesses_Phase2() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(3000)) {
    Logger.log("⏭️ Phase2: ロック取得失敗のためスキップ");
    return;
  }
  try {
    const start = new Date();
    logTime('Phase2 開始', start);
    logProcessTime("転記更新（会話履歴対応）", () => 転記実行_更新付き_会話履歴対応_safe());
    logProcessTime("社名分類（新規A→O）", () => classifyTargets());
    Logger.log("✅ Phase2 実行時間: " + ((new Date() - start) / 1000) + "秒");
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

function runAllProcesses_Phase3() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(3000)) return;

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const start = new Date();
    logTime('Phase3 開始', start);

    logProcessTime("Indeedリンク生成（新規）", () => generateIndeedSearchLinks_NewSheet(ss));
    logProcessTime("Googleマップリンク生成（C列）", () => setHyperlinksInColumnC(ss));
    logProcessTime("電話番号クリックリンク生成", () => createPhoneHyperlinks(ss));
    logProcessTime("相互リンク再生成", () => regenerateLinks(ss));

    const end = new Date();
    Logger.log("✅ Phase3 実行時間: " + ((end - start) / 1000) + "秒");
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

/***************
 * 架電記録の更新
 ***************/
function updateKadenRecordSheet(ss) {
  const tz = Session.getScriptTimeZone();
  const todayStr = Utilities.formatDate(new Date(), tz, "yyyy/MM/dd");
  const names = ["架電記録/新規", "架電記録/現S"];

  names.forEach(name => {
    const sheet = getSheetOrAlert(ss, name);
    if (!sheet) return;

    let rowIndex;
    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      const lastStr = sheet.getRange(lastRow, 1).getDisplayValue();
      if (lastStr !== todayStr) {
        sheet.appendRow([todayStr, 0, 0, 0]);
        rowIndex = sheet.getLastRow();
      } else {
        rowIndex = lastRow;
      }
    } else {
      sheet.appendRow([todayStr, 0, 0, 0]);
      rowIndex = sheet.getLastRow();
    }

    const dateObj = ymdToDate_(todayStr);
    const isHol = isJapanHoliday_(dateObj);
    const firstBizWeek = getFirstBusinessDayOfWeek_(dateObj);
    const firstBizMonth = getFirstBusinessDayOfMonth_(dateObj);
    const fmt = d => Utilities.formatDate(d, tz, "yyyy/MM/dd");

    const paintCols = Math.min(5, sheet.getLastColumn());
    const rng = sheet.getRange(rowIndex, 1, 1, paintCols);

    if (isHol) {
      rng.setBackground('#D3D3D3');
    } else if (fmt(dateObj) === fmt(firstBizMonth)) {
      rng.setBackground('#ADFF2F');
      sheet.getRange("I21").setValue(rowIndex);
    } else if (fmt(dateObj) === fmt(firstBizWeek)) {
      rng.setBackground('#FFF2CC');
      sheet.getRange("I22").setValue(rowIndex);
    } else {
      rng.setBackground(null);
    }
  });

  const shNew = getSheetOrAlert(ss, "架電記録/新規");
  if (!shNew) return;
  const lr = shNew.getLastRow();
  if (lr < 2) return;
  const val = shNew.getRange("J7").getValue();
  shNew.getRange(lr, 5).setValue((typeof val === "number") ? Math.ceil(val) : "");
}

/***************
 * Indeed 検索リンク生成（新規・高速化）
 ***************/
function generateIndeedSearchLinks_NewSheet(ss) {
  const sheet = getSheetOrAlert(ss, "営業先リスト/新規");
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const values = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  const out = new Array(values.length);

  for (let i = 0; i < values.length; i++) {
    const company = values[i][0];
    const address = values[i][2];
    if (!company || !address) { out[i] = [""]; continue; }
    const location = extractPrefectureAndCity(address);
    const url = "https://jp.indeed.com/jobs?q=" + encodeURIComponent(company) + "&l=" + encodeURIComponent(location);
    out[i] = ['=HYPERLINK("' + url + '","Indeed求人")'];
  }
  const rng = sheet.getRange(2, 13, out.length, 1);
  setFormulasIfChanged_(rng, out);
}

/***************
 * 電話番号のクリックリンク化
 ***************/
function createPhoneHyperlinks(ss) {
  const sheetTarget = getSheetOrAlert(ss, "切り返し");
  if (!sheetTarget) return;

  const baseUrl = "https://docs.google.com/spreadsheets/d/" + ss.getId() + "/edit";
  const targetId = sheetTarget.getSheetId();

  const tgtLR = sheetTarget.getLastRow();
  const map = {};
  if (tgtLR >= 1) {
    const tgtVals = sheetTarget.getRange(1, 1, tgtLR, 2).getValues();
    for (let i = 0; i < tgtVals.length; i++) {
      const status = (tgtVals[i][0] || "").toString().trim();
      const ind = (tgtVals[i][1] || "").toString().trim();
      const key = status + "|" + ind;
      if (key !== "|") map[key] = i + 1;
    }
  }

  const sheets = ["営業先リスト/新規", "営業先リスト/現S"];
  for (let s = 0; s < sheets.length; s++) {
    const sh = getSheetOrAlert(ss, sheets[s]);
    if (!sh) continue;

    const lr = sh.getLastRow();
    if (lr < 2) continue;

    const phones = sh.getRange(2, 2, lr - 1, 1).getValues();
    const statuses = sh.getRange(2, 12, lr - 1, 1).getValues();
    const industryCol = (sheets[s] === "営業先リスト/現S") ? 19 : 15;
    const industries = sh.getRange(2, industryCol, lr - 1, 1).getValues();

    const formulas = new Array(lr - 1);
    for (let r = 0; r < lr - 1; r++) {
      const phone = (phones[r][0] || "").toString().trim();
      if (!phone) { formulas[r] = [""]; continue; }

      const status = (statuses[r][0] || "").toString().trim();
      const industry = (industries[r][0] || "").toString().trim();
      const key = status + "|" + industry;
      const targetRow = map[key];

      if (targetRow) {
        const url = baseUrl + "#gid=" + targetId + "&range=B" + targetRow;
        formulas[r] = ['=HYPERLINK("' + url + '","' + phone + '")'];
      } else {
        formulas[r] = ['="' + phone + '"'];
      }
    }
    const rng = sh.getRange(2, 2, formulas.length, 1);
    setFormulasIfChanged_(rng, formulas);
  }
}

/***************
 * 住所 → 都道府県・市区町村 抽出（簡易）
 ***************/
function extractPrefectureAndCity(address) {
  const regex = /(北海道|東京都|大阪府|京都府|.{2,3}県)(.{1,6}市|.{1,6}区|.{1,6}町|.{1,6}村)/;
  const match = (address || "").toString().match(regex);
  return match ? (match[1] + ' ' + match[2]) : address;
}

/***************
 * 現S C列へ外部サイトリンク（P:ID → URL化）
 ***************/
function setHyperlinksInColumnC(ss) {
  const sheet = getSheetOrAlert(ss, "営業先リスト/現S");
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const idValues = sheet.getRange(2, 16, lastRow - 1, 1).getValues();
  const displayTexts = sheet.getRange(2, 3, lastRow - 1, 1).getValues();
  const out = new Array(idValues.length);

  for (let i = 0; i < idValues.length; i++) {
    const id = idValues[i][0];
    const text = displayTexts[i][0];
    if (id && text) {
      const url = "https://ats.rct.airwork.net/agency/a/clients/" + id;
      out[i] = ['=HYPERLINK("' + url + '","' + text + '")'];
    } else if (text) {
      out[i] = ['="' + text + '"'];
    } else {
      out[i] = [""];
    }
  }
  const rng = sheet.getRange(2, 3, out.length, 1);
  setFormulasIfChanged_(rng, out);
}

/***************
 * 相互リンク生成（O(n)化）基盤
 ***************/
function getCompanyData(sheetName, ss) {
  const sheet = getSheetOrAlert(ss, sheetName);
  const map = {};
  if (!sheet) return map;

  const lr = sheet.getLastRow();
  if (lr >= 2) {
    const data = sheet.getRange(2, 1, lr - 1, 1).getValues();
    const sheetId = sheet.getSheetId();
    for (let i = 0; i < data.length; i++) {
      const company = (data[i][0] || "").toString().trim();
      if (company) map[company] = { sheet, sheetName, row: i + 2, sheetId };
    }
  }
  return map;
}

/***************
 * 相互リンク更新（完全修正版）
 * ・getFormulasで比較（リンク消失防止）
 * ・一括処理で高速
 ***************/
function updateLinkFormulas(sheetNames, toMap, ss) {
  sheetNames.forEach(sheetName => {
    const sheet = getSheetOrAlert(ss, sheetName);
    if (!sheet) return;

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    const names = sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues();
    const formulas = new Array(names.length);

    for (let i = 0; i < names.length; i++) {
      const company = (names[i][0] || "").toString().trim();

      if (!company) {
        formulas[i] = [""];
        continue;
      }

      const to = toMap[company];

      if (to) {
        formulas[i] = [
          `=HYPERLINK("#gid=${to.sheetId}&range=A${to.row}", "${escapeForFormula_(company)}")`
        ];
      } else {
        formulas[i] = [`="${escapeForFormula_(company)}"`];
      }
    }

    const rng = sheet.getRange(2, 1, formulas.length, 1);

    // ★ここが最重要（getValues→getFormulasに変更）
    const old = rng.getFormulas();

    if (!arraysEqual2D_(old, formulas)) {
      rng.setFormulas(formulas);
    }
  });
}

/***************
 * 相互リンク一括再生成（完全修正版）
 ***************/
function regenerateLinks(ss) {

  // 新規
  const salesNewMap   = getCompanyData("営業先リスト/新規", ss);
  const historyNewMap = getCompanyData("会話履歴/新規", ss);

  updateLinkFormulas(["営業先リスト/新規"], historyNewMap, ss);
  updateLinkFormulas(["会話履歴/新規"], salesNewMap, ss);

  // 現S
  const salesGenSMap   = getCompanyData("営業先リスト/現S", ss);
  const historyGenSMap = getCompanyData("会話履歴/現S", ss);

  updateLinkFormulas(["営業先リスト/現S"], historyGenSMap, ss);
  updateLinkFormulas(["会話履歴/現S"], salesGenSMap, ss);

  // 掲載状況
  const sheetKisai = getSheetOrAlert(ss, "掲載状況");
  if (sheetKisai) {
    const lr = sheetKisai.getLastRow();
    if (lr >= 2) {
      const names = sheetKisai.getRange(2, 2, lr - 1, 1).getValues();
      const formulas = names.map(([name]) => {
        const n = (name || "").toString().trim();
        const entry = historyGenSMap[n];

        return entry
          ? [`=HYPERLINK("#gid=${entry.sheetId}&range=A${entry.row}", "${escapeForFormula_(n)}")`]
          : [`="${escapeForFormula_(n)}"`];
      });

      const rng = sheetKisai.getRange(2, 2, formulas.length, 1);

      const old = rng.getFormulas();
      if (!arraysEqual2D_(old, formulas)) {
        rng.setFormulas(formulas);
      }
    }
  }

  // 飛び込み
  const sheetTobikomi = getSheetOrAlert(ss, "飛び込み先リスト");
  const sheetHistoryN = getSheetOrAlert(ss, "会話履歴/新規");

  if (sheetTobikomi && sheetHistoryN) {
    const historySheetId = sheetHistoryN.getSheetId();

    const hMap = {};
    const lrH = sheetHistoryN.getLastRow();

    if (lrH >= 2) {
      const hVals = sheetHistoryN.getRange(2, 1, lrH - 1, 1).getValues();
      for (let j = 0; j < hVals.length; j++) {
        const n = (hVals[j][0] || "").toString().trim();
        if (n) hMap[n] = j + 2;
      }
    }

    const lrT = sheetTobikomi.getLastRow();

    if (lrT >= 2) {
      const tNames = sheetTobikomi.getRange(2, 1, lrT - 1, 1).getValues();

      const formulas = tNames.map(([nm]) => {
        const n = (nm || "").toString().trim();

        return hMap[n]
          ? [`=HYPERLINK("#gid=${historySheetId}&range=A${hMap[n]}", "${escapeForFormula_(n)}")`]
          : [`="${escapeForFormula_(n)}"`];
      });

      const rng = sheetTobikomi.getRange(2, 1, formulas.length, 1);

      const old = rng.getFormulas();
      if (!arraysEqual2D_(old, formulas)) {
        rng.setFormulas(formulas);
      }
    }
  }
}

function escapeForFormula_(s) {
  return String(s).replace(/"/g, '""');
}

function ymdToDate_(ymd) {
  const [y, m, d] = String(ymd).split('/').map(Number);
  return new Date(y, m - 1, d);
}

function isJapanHoliday_(date) {
  const tz = Session.getScriptTimeZone();
  const fmt = Utilities.formatDate(date, tz, "yyyy-MM-dd");
  const key = "holiday:" + fmt;
  const cache = CacheService.getScriptCache();
  const hit = cache.get(key);
  if (hit !== null) return hit === "1";

  const CAL_ID = 'ja.japanese#holiday@group.v.calendar.google.com';
  const cal = CalendarApp.getCalendarById(CAL_ID);
  let isHol = false;
  if (cal) {
    const ev = cal.getEventsForDay(date);
    isHol = !!(ev && ev.length > 0);
  }
  const dow = date.getDay();
  if (dow === 0 || dow === 6) isHol = true;
  cache.put(key, isHol ? "1" : "0", 60 * 60 * 24);
  return isHol;
}

function getFirstBusinessDayOfWeek_(date) {
  const d = new Date(date.getTime());
  const deltaToMon = (d.getDay() + 6) % 7;
  d.setDate(d.getDate() - deltaToMon);
  for (let i = 0; i < 5; i++) {
    const cand = new Date(d.getFullYear(), d.getMonth(), d.getDate() + i);
    if (!isJapanHoliday_(cand)) return cand;
  }
  return d;
}

function getFirstBusinessDayOfMonth_(date) {
  const y = date.getFullYear();
  const m = date.getMonth();
  for (let d = 1; d <= 31; d++) {
    const cand = new Date(y, m, d);
    if (cand.getMonth() !== m) break;
    if (!isJapanHoliday_(cand)) return cand;
  }
  return new Date(y, m, 1);
}

// 背景は触らず「文字装飾だけ」リセット（ラベンダー保護のため）
// 掲載開始顧客：優先度＝灰(L有) → ADE判定 → 薄緑(A有)
// ポイント：最初に薄緑で全Aあり行を塗り、ADEと灰で上書きする
function colorKeisaiRowsUnified_(opt){
  const {
    sheetName = '掲載開始顧客',
    headerRows = 1,
    // ADEの色
    blueExistingColor = '#d6eaff',          // 既存S（営業先/現Sに存在）
    yellowStoreColor  = YELLOW_STORE_COLOR, // 「店」を含む
    pinkBaitoruOnly   = PINK_BAITORU_ONLY , // D空白 & E=バイトルのみ & 会社記号なし
    // 新要件の色
    grayIfLHasValue   = '#dddddd',          // L列に値があれば灰（最優先）
    lightGreenIfAHas  = '#ccffcc',          // A列に値があれば薄緑（最下位）
    applySort = true
  } = opt || {};

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(sheetName); if(!sh) return;
  const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  if(lastRow <= headerRows) return;

  // L列参照があるか一応確認（12列以上必要）
  if(lastCol < 12) {
    // L列が無いならA薄緑だけ適用して帰る
    const rngA = sh.getRange(headerRows+1,1,lastRow-headerRows,lastCol);
    rngA.setBackground(null);
    const valsA = rngA.getValues();
    const greenList = [];
    for (let i=0;i<valsA.length;i++){
      const r = headerRows+i+1;
      const a = (valsA[i][0]||'').toString().trim();
      if(a) greenList.push(`${r}:${r}`);
    }
    if(greenList.length) sh.getRangeList(greenList).setBackground(lightGreenIfAHas);
    return;
  }

  const rng  = sh.getRange(headerRows+1,1,lastRow-headerRows,lastCol);
  const vals = rng.getValues();

  // 背景を全クリア（ラベンダー保護は廃止）
  rng.setBackground(null);

  // 既存Sの社名集合（営業先リスト/現S A列）
  const existingSet = new Set();
  const cur = ss.getSheetByName('営業先リスト/現S');
  if(cur && cur.getLastRow()>1){
    cur.getRange(2,1,cur.getLastRow()-1,1).getDisplayValues()
      .forEach(r=>{ const v=(r[0]+'').trim(); if(v) existingSet.add(v); });
  }

  // ① 最下位：Aに値あり → 薄緑
  const greenList = [];
  for (let i=0;i<vals.length;i++){
    const r = headerRows+i+1;
    const a = (vals[i][0]||'').toString().trim();
    if(a) greenList.push(`${r}:${r}`);
  }
  if(greenList.length) sh.getRangeList(greenList).setBackground(lightGreenIfAHas);

  // ② 中間：ADEで上書き（※「その他→橙」は無し！）
  const bucket = {
    [blueExistingColor]: [],
    [yellowStoreColor]:  [],
    [pinkBaitoruOnly]:   []
  };
  for(let i=0;i<vals.length;i++){
    const r = headerRows+i+1;
    const a = (vals[i][0]||'').toString().trim();
    if(!a) continue; // A空は薄緑対象外なのでスキップ
    const d = (vals[i][3]||'').toString().trim();
    const e = (vals[i][4]||'').toString().trim();

    if(existingSet.has(a)){ bucket[blueExistingColor].push(`${r}:${r}`); continue; }
    if(/店/.test(a)){      bucket[yellowStoreColor].push(`${r}:${r}`);   continue; }
    const onlyBaitoru = /^\s*バイトル\s*$/.test(e);
    if(!d && onlyBaitoru && !hasLegalCompanyMarker_(a)){
      bucket[pinkBaitoruOnly].push(`${r}:${r}`); continue;
    }
    // 「その他」は薄緑のまま保持
  }
  if(bucket[blueExistingColor].length) sh.getRangeList(bucket[blueExistingColor]).setBackground(blueExistingColor);
  if(bucket[yellowStoreColor].length)  sh.getRangeList(bucket[yellowStoreColor]).setBackground(yellowStoreColor);
  if(bucket[pinkBaitoruOnly].length)   sh.getRangeList(bucket[pinkBaitoruOnly]).setBackground(pinkBaitoruOnly);

  // ③ 最上位：Lに値あり → 灰色で上書き
  const grayList = [];
  for(let i=0;i<vals.length;i++){
    const r = headerRows+i+1;
    const l = (vals[i][11]||'').toString().trim(); // L列 index 11
    if(l) grayList.push(`${r}:${r}`);
  }
  if(grayList.length) sh.getRangeList(grayList).setBackground(grayIfLHasValue);

  // 並べ替え（上詰め）：灰 → ADE（青→黄→ピンク） → 薄緑
  if(applySort){
    sortColoredRowsToTop_Light_(sh, headerRows, {
      priorityColors: [
        grayIfLHasValue,
        blueExistingColor, yellowStoreColor, pinkBaitoruOnly,
        lightGreenIfAHas
      ]
    });
  }
}





// ==============================
// エントリーポイント（安全一括実行）— 一本化版
// ==============================
function updateAllSheetsAndResetColors_Safe() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 0) 会話履歴←営業先リスト（L→B）逆転記（新規／現S）
  [
    { history: "会話履歴/新規", sales: "営業先リスト/新規" },
    { history: "会話履歴/現S",  sales: "営業先リスト/現S" }
  ].forEach(({ history, sales }) => transferStatusToHistory_(ss, history, sales));

  // 1) 背景色リセット＆G列再設定（新規／現S）
  ["営業先リスト/新規", "営業先リスト/現S"].forEach(name => resetAndPaintTodayRows_(ss, name));

  // 2) データ取得 B～D クリア（内容のみ）
  clearDataSheetColumns();

  // 3) ステータス別の背景色ハイライト（現S）
  highlightCellsByStatus_Param("営業先リスト/現S");

  // 4) 現S：J列が今月以外 かつ B列≠「×」の行を薄黄
  highlightRowsJNotThisMonth();

  // 5) ステータス集計・スコア計算
  countStatusSummary();
  calcSalesPriorityScore();

  // 5.5) 新規：並び替えの後に色（今日電話した＝薄紫／今日かけるべき＝オレンジ／それ以外＝白）
  paintShinkiByPriority_(ss);

  // 6) 掲載開始顧客：色塗りを一括適用（ラベンダー保護・外部for赤・優先整列まで）
  colorKeisaiRowsUnified_({
    sheetName: "掲載開始顧客",
    headerRows: 1,
    // 必要があれば色や優先度をここで上書き可能
    // yellowStoreColor: "#fff9c4",
    // pinkBaitoruOnly:  "#ffd1e6",
    // applySort: true
  });

  // 7) 営業日補完（fillBusinessDays）。失敗しても Toast で知らせて続行
  runFillBusinessDaysSafely_();
}

function runFillBusinessDaysSafely_() {
  try {
    fillBusinessDays();
  } catch (err) {
    Logger.log('fillBusinessDays 実行中に例外: ' + err);
    try {
      SpreadsheetApp.getActive().toast('fillBusinessDays 実行で例外: ' + err, 'Error', 5);
    } catch (e) {}
  }
}

// ==============================
// 共通ユーティリティ
// ==============================
function parseToDate(val) {
  if (val instanceof Date && !isNaN(val)) return val;
  if (typeof val === 'string') {
    const d = new Date(val);
    if (!isNaN(d)) return d;
  }
  return null;
}

function isSameDay(a, b) {
  return a && b &&
    a.getFullYear() === b.getFullYear() &&
    a.getMonth() === b.getMonth() &&
    a.getDate() === b.getDate();
}

// 今日/昨日の基準を1回だけ作る（パフォーマンス）
function getTodayYesterday_() {
  const today = new Date(); today.setHours(0,0,0,0);
  const yest  = new Date(today); yest.setDate(today.getDate() - 1);
  return { today, yest };
}
// ==============================
// 逆転記：営業先(A:L)→会話履歴(Aの一致でB=営業L)
// ==============================
function transferStatusToHistory_(ss, historyName, salesName) {
  const historySheet = ss.getSheetByName(historyName);
  const salesSheet   = ss.getSheetByName(salesName);
  if (!historySheet || !salesSheet) return;

  const salesLast = salesSheet.getLastRow();
  const histLast  = historySheet.getLastRow();
  if (salesLast < 2 || histLast < 2) return;

  // 営業先：A(社名)〜L(ステータス)
  const salesRange = salesSheet.getRange(2, 1, salesLast - 1, 12);
  const salesData  = salesRange.getValues();
  const statusMap  = Object.create(null);
  for (let i=0;i<salesData.length;i++) {
    const name = (salesData[i][0] || "").toString().trim();
    if (name) statusMap[name] = salesData[i][11]; // L列
  }

  // 会話履歴：A列社名→B列へ
  const names = historySheet.getRange(2,1,histLast-1,1).getValues();
  const out   = new Array(names.length);
  for (let i=0;i<names.length;i++) {
    const key = names[i][0] ? names[i][0].toString().trim() : "";
    out[i] = [statusMap[key] || ""];
  }
  historySheet.getRange(2,2,out.length,1).setValues(out);
}

// ==============================
// 背景色リセット＋G列再設定（今日のみ色・値を残す）
// ==============================
function resetAndPaintTodayRows_(ss, sheetName) {
  const sh = ss.getSheetByName(sheetName);
  if (!sh) return;

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2) return;

  // 2行目以降のみ対象（ヘッダーを汚さない）
  const rngData = sh.getRange(2, 1, lastRow - 1, lastCol);
  const data    = rngData.getValues();
  const bg      = rngData.getBackgrounds();

  const { today } = getTodayYesterday_();

  // まず全行デフォルト（空文字=デフォルト色）＆G列クリア
  for (let r=0;r<data.length;r++) {
    for (let c=0;c<lastCol;c++) bg[r][c] = "";
    // I列= index8, G列=index6
    const dt = data[r][8];
    let isToday = false;
    if (dt instanceof Date && !isNaN(dt)) {
      const d = new Date(dt); d.setHours(0,0,0,0);
      isToday = isSameDay(d, today);
    }
    if (!isToday) data[r][6] = ""; // 今日以外はGクリア
    // 今日なら薄紫、Gが数値なら黄緑
    if (isToday) {
      const g = data[r][6];
      const isNum = g !== "" && !isNaN(g);
      for (let c=0;c<lastCol;c++) bg[r][c] = isNum ? "#ADFF2F" : "#E6E6FA";
    }
  }

  // まとめて反映
  rngData.setBackgrounds(bg);
  // G列のみ一括書き戻し
  const gVals = data.map(row => [row[6]]);
  sh.getRange(2,7,gVals.length,1).setValues(gVals);
}

// ==============================
// データ取得 B2:D クリア（内容のみ）
// ==============================
function clearDataSheetColumns() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("データ取得");
  if (!sh) return;
  sh.getRange("B2:D").clearContent(); // 背景や罫線は保持
}
// ==============================
// ステータス別の期限超過ハイライト（行をオレンジ）
// 対象：営業先リスト/新規 or 営業先リスト/現S
// ==============================
function highlightCellsByStatus_Param(sheetName) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sh) return;

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2) return;

  const range = sh.getRange(2,1,lastRow-1,lastCol);
  const data  = range.getValues();
  const bg    = range.getBackgrounds();

  const { today, yest } = getTodayYesterday_();

  // 設定（L列=idx11、既定の基準日は J(=idx9) or N(=idx13)）
  const cfg = {
    "Ⓐ1か月以内ニーズ有 × 選んでくれる": { threshold: 7,   dateIndex: 9 },
    "Ⓓ2–3か月ニーズ有 × 選んでくれる": { threshold: 14,  dateIndex: 9 },
    "Ⓓ3か月超先でも選んでくれる":        { threshold: 30,  dateIndex: 9 },
    "Ⓑ1か月以内ニーズ有 × 選ぶか不明":  { threshold: 10,  dateIndex: 9 },
    "Ⓒ2–3か月ニーズ有 × 選ぶか不明":  { threshold: 20,  dateIndex: 9 },
    "Ⓒ3か月超先 × 選ぶか不明":        { threshold: 45,  dateIndex: 9 },
    "Ⓒ1か月以内ニーズ有 × 選ばなそう":{ threshold: 10,  dateIndex: 9 },
    "Ⓒ2–3か月ニーズ有 × 選ばなそう":{ threshold: 30,  dateIndex: 9 },
    "Ⓒ3か月超先 × 選ばなそう":       { threshold: 60,  dateIndex: 9 },
    "判断がつかない":                  { threshold: 3,   dateIndex: 9 },
    "一旦追わない":                    { threshold: 180, dateIndex: 9 }
  };

  for (let r=0;r<data.length;r++) {
    const row = data[r];
    const status = row[11]; // L
    if (!status || status === "一旦追わない") continue;

    const set = cfg[status];
    if (!set) continue;

    // 基準日：cfg.dateIndex か N列(=idx13) のどちらか
    const base = parseToDate(row[set.dateIndex]) || parseToDate(row[13]) || new Date("2000-01-01");
    base.setHours(0,0,0,0);

    // 直近コール日(I=idx8)が今日・昨日は対象外、また基準日が今日も対象外
    const call = parseToDate(row[8]);
    if ((call && (isSameDay(call,today) || isSameDay(call,yest))) || isSameDay(base,today)) continue;

    const diff = Math.floor((today - base) / 86400000);
    if (diff >= set.threshold) {
      for (let c=0;c<lastCol;c++) bg[r][c] = "#FFA500";
    }
  }
  range.setBackgrounds(bg);
}

// ==============================
// 現S：J列が今月以外 かつ B列≠「×」 かつ未着色 → 薄黄で行塗り
// ==============================
function highlightRowsJNotThisMonth() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("営業先リスト/現S");
  if (!sh) return;

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2) return;

  const range = sh.getRange(2,1,lastRow-1,lastCol);
  const data  = range.getValues();
  const bg    = range.getBackgrounds();

  const now = new Date();
  const y = now.getFullYear();
  const m = now.getMonth(); // 0-index

  for (let r=0;r<data.length;r++) {
    const dateInJ = data[r][9];  // J
    const markB   = data[r][1];  // B
    const isSameMonth = (dateInJ instanceof Date) &&
      dateInJ.getFullYear() === y && dateInJ.getMonth() === m;

    const alreadyColored = (bg[r][0] && String(bg[r][0]).toLowerCase() !== "#ffffff" && bg[r][0] !== "");

    if (!isSameMonth && markB !== "×" && !alreadyColored) {
      for (let c=0;c<lastCol;c++) bg[r][c] = "#FFFFE0";
    }
  }
  range.setBackgrounds(bg);
}

// 重複増殖ストッパー版（空白ゆれ完全吸収版）
function transfer営業先リスト完全版() {
  // === 設定 ===
  const sourceSpreadsheetUrl = 'https://docs.google.com/spreadsheets/d/1PU4DErlyWoqML31ZFK4CAvVS-BsFfGsbGV3Mg0C7f-o/edit';
  const targetSpreadsheetUrl = 'https://docs.google.com/spreadsheets/d/1UORXW2Pl-C6ESVTI63a3Z9BLFYA3bhWAsiINUxpaDCw/edit';
  const sourceSheetName = '顧客データ';
  const targetSheetName = '営業先リスト/現S';
  const statusSheetName = '掲載状況';
  const 担当者 = '菅野光';

  // 正規化
  // ・NFKCで全角半角差を吸収
  // ・不可視文字除去
  // ・すべての空白（半角/全角/改行/タブ）を除去
  const normalize = s => String(s ?? '')
    .normalize('NFKC')
    .replace(/[​﻿ ]/g, '')   // 不可視文字・NBSP
    .replace(/\s+/g, '')                    // 半角空白・改行・タブなど
    .replace(/　+/g, '')                    // 全角スペース
    .trim();

  const srcSh = SpreadsheetApp.openByUrl(sourceSpreadsheetUrl).getSheetByName(sourceSheetName);
  const tgtSh = SpreadsheetApp.openByUrl(targetSpreadsheetUrl).getSheetByName(targetSheetName);
  const stsSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(statusSheetName);

  const src = srcSh.getDataRange().getValues();
  const tgt = tgtSh.getDataRange().getValues();

  // --- ソースを正規化キーでマップ化（H列=管理番号）
  const srcMap = {};
  for (let i = 1; i < src.length; i++) {
    const r = src[i];
    if (r[5] !== 担当者) continue;   // F列: 担当者

    const rawKey = r[7];             // H列: 管理番号
    const key = normalize(rawKey);
    if (!key) continue;

    srcMap[key] = {
      管理番号原文: rawKey,
      C: r[9],
      D: r[13],
      O: r[21],
      P: r[4],
      Q: r[2],
      R: r[1]
    };
  }

  // --- ターゲット既存キー集合（A列全件を正規化）
  const tgtKeys = new Set();
  for (let i = 1; i < tgt.length; i++) {
    const k = normalize(tgt[i][0]);  // A列
    if (k) tgtKeys.add(k);
  }

  // --- まず更新：正規化キー一致行の C,D,O,P,Q,R を上書き（A列は触らない）
  //     まとめて書くため、一致しない行は今の中身（数式なら数式）をそのまま書き戻す
  if (tgt.length > 1) {
    const n = tgt.length - 1;
    const cd = tgtSh.getRange(2, 3, n, 2), or = tgtSh.getRange(2, 15, n, 4);
    const keep = rng => {
      const v = rng.getValues(), f = rng.getFormulas();
      return v.map((r, i) => r.map((x, j) => f[i][j] || x));
    };
    const cdVals = keep(cd), orVals = keep(or);
    let hit = 0;
    for (let i = 1; i < tgt.length; i++) {
      const e = srcMap[normalize(tgt[i][0])];
      if (!e) continue;
      cdVals[i - 1] = [e.C, e.D];
      orVals[i - 1] = [e.O, e.P, e.Q, e.R];
      hit++;
    }
    if (hit) {
      cd.setValues(cdVals);
      or.setValues(orVals);
    }
  }

  // --- 新規追加：既存セットにない正規化キーのみ
  const add = [];
  for (const key in srcMap) {
    if (!tgtKeys.has(key)) {
      const e = srcMap[key];
      add.push([
        e.管理番号原文, '', e.C, e.D, '', '', '', '', '', '', '', '', '', '', e.O, e.P, e.Q, e.R
      ]);
      tgtKeys.add(key); // 同一実行中の重複追加防止
    }
  }

  if (add.length > 0) {
    const startRow = tgtSh.getLastRow() + 1;
    tgtSh.getRange(startRow, 1, add.length, add[0].length).setValues(add);
  }

  // --- 掲載状況 → N列反映（P列参照）＆N列降順ソート
  if (stsSh) {
    const statusLastRow = stsSh.getLastRow();
    if (statusLastRow > 1) {
      const sts = stsSh.getRange(2, 1, statusLastRow - 1, 4).getValues();
      const statusMap = {};

      for (const r of sts) {
        const id = normalize(r[0]);
        const status = r[2];
        const date = r[3];
        if (status === '配信終了' && id && date) {
          statusMap[id] = date;
        }
      }

      const targetLastRow = tgtSh.getLastRow();
      if (targetLastRow > 1) {
        const pVals = tgtSh.getRange(2, 16, targetLastRow - 1, 1).getValues(); // P列
        const nVals = pVals.map(v => [statusMap[normalize(v[0])] || '']);
        tgtSh.getRange(2, 14, nVals.length, 1).setValues(nVals); // N列

        tgtSh.getRange(2, 1, targetLastRow - 1, tgtSh.getLastColumn())
          .sort({ column: 14, ascending: false });
      }
    }
  }

  Logger.log('✅ 完了：空白ゆれ完全吸収版で重複増殖を抑止');
}

function updateCallLog(sheet, dateString, columnIndex) {
  // A列（2行目以降・最終行まで）で今日の日付の行を探す
  const tz = Session.getScriptTimeZone();
  const lastRow = sheet.getLastRow();
  const dates = lastRow >= 2 ? sheet.getRange(2, 1, lastRow - 1, 1).getValues() : [];
  let lastDateRow = 0; // A列に日付がある最後の行番号

  for (let i = 0; i < dates.length; i++) {
    const v = dates[i][0];
    if (v === null || String(v).trim() === "") continue;
    // 日付型はフォーマット、文字列は先頭10文字（例："2025/02/26..."）で比較
    const s = v instanceof Date ? Utilities.formatDate(v, tz, "yyyy/MM/dd") : String(v).trim().substring(0, 10);
    lastDateRow = i + 2;
    if (s === dateString) {
      const cell = sheet.getRange(i + 2, columnIndex);
      cell.setValue((cell.getValue() || 0) + 1);
      return;
    }
  }

  // 該当日付がなければ、最後の日付行の下（なければヘッダー直下）に新規行を挿入
  const newRow = [dateString, 0, 0];
  newRow[columnIndex - 1] = 1;
  const insertRow = lastDateRow || 1;
  sheet.insertRowAfter(insertRow);
  sheet.getRange(insertRow + 1, 1, 1, newRow.length).setValues([newRow]);
}

function copyValuesToNextRows() {
  // スプレッドシートを取得
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 対象シートを取得
  const sheet = ss.getSheetByName("架電記録/新規");

  // I15 → I16
  const valueI15 = sheet.getRange("I15").getValue();
  sheet.getRange("I16").setValue(valueI15);

  // I17 → I18
  const valueI17 = sheet.getRange("I17").getValue();
  sheet.getRange("I18").setValue(valueI17);

  // I19 → I20
  const valueI19 = sheet.getRange("I19").getValue();
  sheet.getRange("I20").setValue(valueI19);
}


function countStatusSummary() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('営業先リスト/新規');
  if (!sheet || sheet.getLastRow() < 2) return;
  const values = sheet.getRange(2, 12, sheet.getLastRow() - 1).getValues().flat(); // L列

  const counts = { 'Ⓐ': 0, 'Ⓑ': 0, 'Ⓒ': 0, 'Ⓓ': 0, '判断がつかない': 0, '繋がってない': 0 };
  let followCount = 0;

  values.forEach(v => {
    if (!v) return;
    const value = String(v);
    if (value.includes('Ⓐ')) counts['Ⓐ']++;
    else if (value.includes('Ⓑ')) counts['Ⓑ']++;
    else if (value.includes('Ⓒ')) counts['Ⓒ']++;
    else if (value.includes('Ⓓ')) counts['Ⓓ']++;
    else if (value.includes('判断がつかない')) counts['判断がつかない']++;
    else if (value.includes('キーマンと繋がってない')) counts['繋がってない']++;

    if (value !== "一旦追わない") followCount++;
  });

  const total = counts['Ⓐ'] + counts['Ⓑ'] + counts['Ⓒ'] + counts['Ⓓ'];

  const summary =
    `Ⓐ${counts['Ⓐ']}/10社　Ⓓ${counts['Ⓓ']}/30社　Ⓑ${counts['Ⓑ']}/15社　Ⓒ${counts['Ⓒ']}/25社\n` +
    `判断がつかない ${counts['判断がつかない']}社　繋がってない ${counts['繋がってない']}社\n` +
    `合計 ${total}社　追いS数 ${followCount}社`;

  sheet.getRange('A1').setValue(summary).setWrap(true);
}


function calcSalesPriorityScore() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('営業先リスト/新規');
  if (!sheet) throw new Error('営業先リスト/新規が見つかりません');
  const board = ss.getSheetByName('現在地ボード');
  if (board && String(board.getRange('Z1').getValue()).trim() === 'LOCK') return;
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(3000)) return;
  try {
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;
    const width = Math.max(22, sheet.getLastColumn());
    if (sheet.getMaxColumns() < width) throw new Error('必要な列が不足しています');
    const data = sheet.getRange(2, 1, lastRow - 1, width).getValues();
    const tz = ss.getSpreadsheetTimeZone();
    const now = new Date();
    const historySheet = ss.getSheetByName('会話履歴/新規');
    const results = salesPriorityResults_(ss, data, now);
    // 内容・数式・書式を保ったバックアップ。旧方式の全行setValuesは使用しない。
    backupSheet_(ss, sheet, '_backup_営業先リスト新規');
    sheet.getRange(2, 17, results.length, 1).setValues(results.map(r => [r.score]));
    sheet.getRange(2, 17, results.length, 1).setNotes(results.map(r => ['更新 ' + Utilities.formatDate(now, tz, 'yyyy/MM/dd HH:mm') + '\n' + r.reason]));
    sheet.getRange(2, 1, lastRow - 1, sheet.getMaxColumns()).sort({ column: 17, ascending: false });
    // 行番号が変わるため、新規の相互リンクのみ既存の生成処理で更新。
    if (historySheet) {
      updateLinkFormulas(['営業先リスト/新規'], getCompanyData('会話履歴/新規', ss), ss);
      updateLinkFormulas(['会話履歴/新規'], getCompanyData('営業先リスト/新規', ss), ss);
    }
    SpreadsheetApp.flush();
    console.log('優先順位更新完了: ' + results.length + '行。Q列メモに理由を記録。');
  } finally {
    lock.releaseLock();
  }
}

// 営業先リスト/新規の各行の点数・理由・今日かけるべきか（calcSalesPriorityScore と色塗りで共通）
function salesPriorityResults_(ss, data, now) {
  const tz = ss.getSpreadsheetTimeZone();
  const today = salesPriorityDay_(now, tz);
  const hour = Number(Utilities.formatDate(now, tz, 'H')) + Number(Utilities.formatDate(now, tz, 'm')) / 60;
  const historySheet = ss.getSheetByName('会話履歴/新規');
  const histories = Object.create(null);
  if (historySheet && historySheet.getLastRow() > 1) {
    historySheet.getRange(2, 1, historySheet.getLastRow() - 1, 3).getValues().forEach(r => {
      const name = String(r[0] || '').trim();
      const h = salesPriorityLatestHistory_(r[2], tz);
      if (name && h && (!histories[name] || h.stamp >= histories[name].stamp)) histories[name] = h;
    });
  }
  return data.map(row => salesPriorityRow_(row, { today: today, hour: hour, tz: tz, history: histories[String(row[0] || '').trim()] }));
}

// 営業先リスト/新規の行全体を塗る：今日電話した＝薄紫（G列に数字があれば黄緑）／今日かけるべき＝オレンジ／それ以外＝白
function paintShinkiByPriority_(ss) {
  const sh = ss.getSheetByName('営業先リスト/新規');
  if (!sh || sh.getLastRow() < 2) return;
  const lastCol = sh.getLastColumn();
  const n = sh.getLastRow() - 1;
  const data = sh.getRange(2, 1, n, Math.max(22, lastCol)).getValues();
  const results = salesPriorityResults_(ss, data, new Date());
  const called = getCalledToday_(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd'));
  const bgs = data.map((row, i) => {
    const g = row[6];
    const color = called.has(String(row[0] || '').trim())
      ? (g !== '' && !isNaN(g) ? '#ADFF2F' : '#E6E6FA')
      : results[i].due ? '#FFA500' : null;
    return Array(lastCol).fill(color);
  });
  sh.getRange(2, 1, n, lastCol).setBackgrounds(bgs);
}

// 「今日電話した」社名の記録（スクリプトのプロパティに日付ごとに保存。前日以前の分は自動で消す）
function getCalledToday_(ymd) {
  const raw = PropertiesService.getScriptProperties().getProperty('CALLED_' + ymd);
  return new Set(raw ? JSON.parse(raw) : []);
}
// 記録に失敗しても入力時の他の処理は止めない
function markCalledToday_(name, ymd) {
  const key = String(name || '').trim();
  if (!key) return;
  try {
    const props = PropertiesService.getScriptProperties();
    const raw = props.getProperty('CALLED_' + ymd);
    if (!raw) {
      Object.keys(props.getProperties()).forEach(k => { if (k.indexOf('CALLED_') === 0) props.deleteProperty(k); });
    }
    const list = raw ? JSON.parse(raw) : [];
    if (list.indexOf(key) < 0) {
      list.push(key);
      props.setProperty('CALLED_' + ymd, JSON.stringify(list));
    }
  } catch (err) {
    console.error('今日電話した記録に失敗:', err);
  }
}

function backupSheet_(ss, sourceSheet, backupSheetName) {
  let backup = ss.getSheetByName(backupSheetName);
  if (!backup) backup = ss.insertSheet(backupSheetName);
  const range = sourceSheet.getDataRange();
  if (backup.getMaxRows() < range.getNumRows()) backup.insertRowsAfter(backup.getMaxRows(), range.getNumRows() - backup.getMaxRows());
  if (backup.getMaxColumns() < sourceSheet.getMaxColumns()) backup.insertColumnsAfter(backup.getMaxColumns(), sourceSheet.getMaxColumns() - backup.getMaxColumns());
  backup.clear();
  range.copyTo(backup.getRange(1, 1));
  if (!backup.isSheetHidden()) backup.hideSheet();
}

// 営業方針の初期値。成約率を学習した予測値ではなく、再連絡の目安。
function salesPriorityPolicy_(status) {
  const s = String(status || '').normalize('NFKC').replace(/[\s​]/g, '');
  if (s.includes('一旦追わない')) return { base: -10000, interval: 30, label: '一旦追わない', stop: true };
  const near = /1[かヶケ]月以内/.test(s);
  const mid = /2[–ー－−〜～-]3[かヶケ]月/.test(s);
  const far = /3[かヶケ]月超/.test(s);
  const yes = s.includes('選んでくれる');
  const unknown = s.includes('選ぶか不明');
  const no = s.includes('選ばなそう');
  if (near && yes) return { base: 90, interval: 3, label: '1か月以内・選んでくれる' };
  if (near && unknown) return { base: 75, interval: 7, label: '1か月以内・選ぶか不明' };
  if (mid && yes) return { base: 60, interval: 14, label: '2–3か月・選んでくれる' };
  if (near && no) return { base: 50, interval: 14, label: '1か月以内・選ばなそう' };
  if (mid && unknown) return { base: 45, interval: 14, label: '2–3か月・選ぶか不明' };
  if (far && yes) return { base: 40, interval: 30, label: '3か月超・選んでくれる' };
  if (far && unknown) return { base: 25, interval: 30, label: '3か月超・選ぶか不明' };
  if (no) return { base: 10, interval: 30, label: '長期・選ばなそう' };
  if (/キーマンと繋がって(い)?ない/.test(s)) return { base: 35, interval: 3, label: 'キーマン未接触' };
  if (!s || s.includes('判断がつかない')) return { base: 40, interval: 7, label: s ? '判断未確定' : 'ステータス未入力' };
  return { base: 35, interval: 7, label: 'ステータス要確認' };
}

function salesPriorityDay_(v, tz) {
  if (v === '' || v === null || v === undefined) return null;
  if (v instanceof Date) {
    if (isNaN(v.getTime())) return null;
    v = Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  } else if (typeof v === 'number') {
    return Number.isFinite(v) && v > 0 ? Math.floor(v) - 25569 : null;
  }
  const m = String(v).trim().match(/^(\d{4})[\/-](\d{1,2})[\/-](\d{1,2})$/);
  if (!m) return null;
  const d = new Date(Date.UTC(+m[1], +m[2] - 1, +m[3]));
  if (d.getUTCFullYear() !== +m[1] || d.getUTCMonth() !== +m[2] - 1 || d.getUTCDate() !== +m[3]) return null;
  return d.getTime() / 86400000;
}

function salesPriorityRow_(row, ctx) {
  if (!String(row[0] || '').trim()) return { score: -20000, reason: '社名なし', due: false };
  const p = salesPriorityPolicy_(row[11]);
  if (p.stop) return { score: p.base, reason: p.label + '：架電対象外', due: false };
  let score = p.base;
  const reasons = [p.label + ' ' + p.base + '点'];
  const add = (n, label) => { score += n; reasons.push(label + ' ' + (n >= 0 ? '+' : '') + n); };
  const i = salesPriorityDay_(row[8], ctx.tz);
  const j = salesPriorityDay_(row[9], ctx.tz);
  const history = ctx.history || null;
  const callback = salesPriorityCallback_(history, ctx);
  if (callback && String(row[1] || '').trim() && ![i, j].some(d => d !== null && d > history.day && d <= ctx.today)) {
    const due = callback.day < ctx.today || (callback.day === ctx.today && ctx.hour >= callback.hour);
    return { score: (due ? 1000 : -500) + p.base, reason: (due ? '再架電指示の日時到来：履歴を確認して架電' : '再架電指示の日時前：待機') + '\n最新履歴：' + callback.text, due: due };
  }
  const historyDay = history && history.day <= ctx.today ? history.day : null;
  const valid = [i, j, historyDay].filter(d => d !== null && d <= ctx.today);
  // I/Jの両方を見る。旧処理で片方だけが更新された行にも対応。
  const latest = valid.length ? Math.max.apply(null, valid) : null;
  if ([i, j].some(d => d !== null && d > ctx.today)) reasons.push('未来の接触日あり：要確認（加点なし）');
  if (latest === ctx.today) return { score: -1000 + p.base, reason: '本日接触記録あり：通常の再架電は後回し。' + reasons[0], due: false };
  if (!String(row[1] || '').trim()) return { score: -2000 + p.base, reason: '電話番号なし：連絡先の確認が必要', due: false };
  const gap = latest === null ? null : ctx.today - latest;
  if (gap === 1) add(-40, '前日に接触');
  // 有効接触が記録されていればJ、なければ直近I/Jを再連絡間隔の基準にする。
  const baseDay = j !== null && j <= ctx.today ? j : latest;
  // 今日かけるべき：接触日が未記録、または再連絡の目安日数が過ぎた（前日に接触した先は除く）
  const due = gap !== 1 && (baseDay === null || ctx.today - baseDay >= p.interval);
  if (baseDay !== null) {
    const elapsed = ctx.today - baseDay;
    if (elapsed < p.interval) add(-30, '再連絡目安' + p.interval + '日未満');
    else add(Math.min(25, Math.floor((elapsed - p.interval) / p.interval) * 5 + 5), '再連絡目安到来');
    reasons.push('最終接触から' + gap + '日');
  } else {
    reasons.push('接触日未記録：古い顧客と決めつけず加点なし');
  }
  const slot = String(row[10] || '').trim();
  const slots = { '午前中': [9, 12], '12時～15時': [12, 15], '15時～18時': [15, 18] };
  if (slots[slot]) add(ctx.hour >= slots[slot][0] && ctx.hour < slots[slot][1] ? 10 : -8, '記録済みの接触時間帯');
  const conn = String(row[15] || '');
  if (/低|ブロ|嘘/.test(conn)) add(-8, 'つながりにくい');
  else if (/中|いないことが多い/.test(conn)) add(-3, '不在が多い');
  else if (/高|普通/.test(conn)) add(3, 'つながりやすい');
  if (String(row[3] || '').trim()) add(3, 'キーマン把握');
  // N列は基本コードで有効接触時に増える。架電失敗数と見なして減点しない。
  const ap = String(row[13] || '').match(/アポイント\/(\d+)回/);
  if (ap) add(Math.min(6, Number(ap[1]) * 2), '過去アポイント実績（上限6点）');
  if (history && /電話|連絡|改め|折り?返/.test(history.text)) reasons.push('会話履歴の最新内容も確認：' + history.text);
  // 予算をつくる人（T列）：話せている＞不明・未回答＞話せていない。再架電の約束（1000点台）より下に収める
  const budget = salesPriorityBudget_(row[19]);
  add(budget.offset, budget.label);
  return { score: score, reason: reasons.join('\n'), due: due };
}

// 自由文からは「最新の記録にある、明確な再架電指示」だけを候補化する。
function salesPriorityLatestHistory_(text, tz) {
  let latest = null;
  String(text || '').split('\n').forEach(line => {
    const m = line.match(/^(\d{4}\/\d{1,2}\/\d{1,2})\s+(\d{1,2})時\s*-\s*(.*)$/);
    if (!m) return;
    const day = salesPriorityDay_(m[1], tz);
    if (day === null) return;
    const stamp = day * 24 + Number(m[2]);
    if (!latest || stamp >= latest.stamp) latest = { day: day, hour: +m[2], stamp: stamp, text: m[3] };
  });
  return latest;
}

function salesPriorityCallback_(record, ctx) {
  if (!record || record.day > ctx.today || ctx.today - record.day > 30) return null;
  const s = record.text.normalize('NFKC');
  // 折返し待ち、否定、不確実な見立てを自動の約束扱いにしない。
  if (/折り?返し|折り?返す|かも|たぶん|多分|不要|しない|しなく|電話.*(断|禁止)|いるなら/.test(s)) return null;
  if (!/電話して|再架電|かけ直|掛け直|改めて連絡|改めて電話|再度連絡|再度電話/.test(s)) return null;
  let day = null;
  let m = s.match(/(?:(\d{4})[\/年])?(\d{1,2})[\/月](\d{1,2})日?/);
  if (m) {
    let year = m[1] ? +m[1] : new Date(record.day * 86400000).getUTCFullYear();
    day = salesPriorityDay_(year + '/' + m[2] + '/' + m[3], ctx.tz);
    if (!m[1] && day !== null && day < record.day && +m[2] === 1 && new Date(record.day * 86400000).getUTCMonth() === 11) day = salesPriorityDay_((year + 1) + '/' + m[2] + '/' + m[3], ctx.tz);
  } else if (s.includes('明後日')) day = record.day + 2;
  else if (s.includes('明日')) day = record.day + 1;
  else if (s.includes('今日')) day = record.day;
  else if ((m = s.match(/来週([月火水木金土日])曜/))) {
    const dow = new Date(record.day * 86400000).getUTCDay();
    day = record.day - ((dow + 6) % 7) + 7 + '月火水木金土日'.indexOf(m[1]);
  }
  if (day === null || day < record.day || day - record.day > 31 || ctx.today - day > 7) return null;
  const tm = s.match(/(?:の|に|\s|^)(\d{1,2})時(?:半|([0-5]?\d)分)?/);
  let hour = tm ? Number(tm[1]) + (tm[0].includes('半') ? .5 : Number(tm[2] || 0) / 60) : 0;
  if (hour > 23) return null;
  if (tm && /午後/.test(s) && hour < 12) hour += 12;
  if (tm && hour < 8 && !/午前|午後/.test(s)) return null;
  // 複数候補・範囲指定は自動で決めない。
  if (/または|もしくは|か\d|時[～〜~–-]|日[～〜~–-]/.test(s)) return null;
  return { day: day, hour: hour, text: record.text };
}
function testSalesPriorityRules() {
  const ctx = { today: salesPriorityDay_('2026/09/14', 'Asia/Tokyo'), hour: 15, tz: 'Asia/Tokyo' };
  const row = (status, day) => { const r = Array(23).fill(''); r[0] = 'テスト'; r[1] = '000-0000-0000'; r[11] = status; r[8] = day || ''; r[9] = day || ''; return r; };
  let n = 0;
  const check = (v, message) => { if (!v) throw new Error(message); n++; };
  const a = row('Ⓐ1か月以内ニーズ有 × 選んでくれる', '2026/07/15');
  const far = row('Ⓒ3か月超先 × 選ぶか不明', '2026/07/15');
  check(salesPriorityRow_(a, ctx).score > salesPriorityRow_(far, ctx).score, '短期見込み優先');
  check(salesPriorityPolicy_('Ⓓ2–3か月ニーズ有 × 選んでくれる').base === 60, '実際のD表記');
  check(salesPriorityPolicy_('キーマンと繋がってない').base === 35, '未接触表記');
  check(salesPriorityRow_(row('一旦追わない'), ctx).score < -9000, '対象外');
  check(salesPriorityRow_(row(a[11], '2026/09/14'), ctx).score < 0, '本日接触');
  check(salesPriorityRow_(row('', '2026/07/15'), ctx).score < salesPriorityRow_(a, ctx).score, '空欄が最優先にならない');
  check(salesPriorityRow_(row(a[11], '2026/09/13'), ctx).score < salesPriorityRow_(a, ctx).score, '前日の連絡を抑制');
  const calls = a.slice(); calls[13] = '午前中/100回 13-15時/100回 15-17時/100回 アポイント/100回';
  check(salesPriorityRow_(calls, ctx).score === salesPriorityRow_(a, ctx).score + 6, '実績上限・成功回数を失敗扱いしない');
  const timeRow = a.slice(); timeRow[10] = '15時～18時';
  check(salesPriorityRow_(timeRow, ctx).score > salesPriorityRow_(timeRow, Object.assign({}, ctx, {hour: 11})).score, '既知の時間帯');
  check(salesPriorityDay_('2026/02/30', ctx.tz) === null, '不正日付');
  check(salesPriorityDay_(46279, ctx.tz) === ctx.today, '日付シリアル');
  const h = salesPriorityLatestHistory_('2026/09/03 13時 - 来週月曜日の13時に電話して\n2026/09/14 15時 - 今日は配達にでてしまった', ctx.tz);
  check(h.day === ctx.today && salesPriorityCallback_(h, ctx) === null, '後続履歴で古い約束を無効化');
  check(salesPriorityRow_(a, Object.assign({}, ctx, {history: h})).score < 0, '日付列未更新でも本日の履歴を反映');
  const future = salesPriorityLatestHistory_('2026/09/14 10時 - 9/18の15時に電話して', ctx.tz);
  check(salesPriorityRow_(a, Object.assign({}, ctx, {history: future})).score < 0, '約束の日時前は待機');
  const due = salesPriorityLatestHistory_('2026/09/14 10時 - 今日の13時に電話して', ctx.tz);
  check(salesPriorityRow_(a, Object.assign({}, ctx, {history: due})).score > 1000, '本日の明確な折り返し指示を優先');
  check(salesPriorityCallback_(salesPriorityLatestHistory_('2026/09/14 10時 - 明日ならいるかも。電話して', ctx.tz), ctx) === null, '曖昧な指示は推測しない');
  check(salesPriorityCallback_(salesPriorityLatestHistory_('2025/09/14 10時 - 明日電話して', ctx.tz), ctx) === null, '古い記録は再架電指示にしない');
  const tier = (v) => { const r = a.slice(); r[19] = v; return salesPriorityRow_(r, ctx).score; };
  check(tier('話せている（AI仮）') > tier('') && tier('') > tier('話せていない'), '予算をつくる人の段');
  check(tier('不明') === tier(''), '不明と未回答は同じ段');
  check(tier('話せていない（AI仮）｜社長') === tier('話せていない'), '仮・メモ付きでも同じ判定');
  check(tier('話せている') < 1000, '再架電の約束より下');
  console.log('PASS: ' + n + '件。シートへの書き込みなし。');
}

// T列「予算をつくる人」の段。値は「話せている」「話せていない」「不明」（末尾に（AI仮）や｜メモが付くことがある）
// 通常の点数はおよそ -70〜140 点の幅なので、段の差（300点）で順番が入れ替わらない
function salesPriorityBudget_(v) {
  const s = String(v || '').normalize('NFKC').replace(/\s/g, '');
  if (s.indexOf('話せていない') === 0) return { offset: 0, label: '予算をつくる人：話せていない' };
  if (s.indexOf('話せている') === 0) return { offset: 600, label: '予算をつくる人：話せている' };
  if (s.indexOf('不明') === 0) return { offset: 300, label: '予算をつくる人：不明' };
  return { offset: 300, label: '予算をつくる人：未回答' };
}

function 転記実行_更新付き_会話履歴対応_safe() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configs = [
    { src: "営業先リスト/新規", conv: "会話履歴/新規", addrCol: 17, color: "#FFC0CB", darkColor: "#FF8A9A" }, // 薄ピンク / やさしめ濃ピンク
    { src: "営業先リスト/現S",  conv: "会話履歴/現S",  addrCol: 14, color: "#ADD8E6", darkColor: "#5B9BD5" }  // 水色 / やさしめ濃ブルー
  ];
  const targetSheet = ss.getSheetByName("メール管理");
  if (!targetSheet) throw new Error("メール管理 シートが見つからぬぞい");

  const targetData = targetSheet.getDataRange().getValues();
  const keyMap = {};
  targetData.forEach((row, i) => { if (row[0]) keyMap[row[0]] = i + 1; });

  // 既存行のA〜H（8列）を手元で更新してから1回で書き込む
  const block = targetData.map(r => { const b = r.slice(0, 8); while (b.length < 8) b.push(""); return b; });
  const toAppend = [];
  const colorLog = [];
  const processedKeys = new Set();
  let updated = 0;

  configs.forEach(cfg => {
    const sheet = ss.getSheetByName(cfg.src);
    if (!sheet) { Logger.log(`⚠️ シートが無い：${cfg.src}`); return; }
    const data = sheet.getDataRange().getValues();

    const convMap = {};
    const cvs = ss.getSheetByName(cfg.conv);
    if (cvs) {
      cvs.getDataRange().getValues().forEach((r, i) => { if (i > 0 && r[0]) convMap[r[0]] = r[2]; });
    } else {
      Logger.log(`⚠️ 会話履歴シートが無い：${cfg.conv}`);
    }

    data.forEach((row, i) => {
      const key = row[0];
      if (i === 0 || !key || processedKeys.has(key)) return;
      if (row[11] === "一旦追わない") return;
      if (!row[cfg.addrCol]) return;

      const isGenS = cfg.src === "営業先リスト/現S";
      const newRow = [key, row[3], "", row[9], convMap[key] || "", row[11], row[isGenS ? 18 : 14], row[isGenS ? 14 : 17]];

      if (keyMap[key]) {
        const rowNum = keyMap[key];
        newRow[2] = block[rowNum - 1][2]; // C列は既存の値を残す
        block[rowNum - 1] = newRow;
        updated++;
        colorLog.push({ rowNum, color: cfg.color, darkColor: cfg.darkColor });
      } else {
        toAppend.push({ values: newRow, color: cfg.color, darkColor: cfg.darkColor });
      }
      processedKeys.add(key);
    });
  });

  Logger.log(`更新: ${updated}件、追加: ${toAppend.length}件`);

  if (updated) targetSheet.getRange(1, 1, block.length, 8).setValues(block);
  if (toAppend.length) {
    const start = targetSheet.getLastRow() + 1;
    targetSheet.getRange(start, 1, toAppend.length, 8).setValues(toAppend.map(r => r.values));
    toAppend.forEach((r, i) => colorLog.push({ rowNum: start + i, color: r.color, darkColor: r.darkColor }));
  }

  const today = new Date();
  const thresholdDate = new Date(today);
  thresholdDate.setMonth(today.getMonth() - 1);

  // 書き込み後のシートを読み直して並べ替え・色付け
  const allData = targetSheet.getDataRange().getValues();
  const header = allData[0];

  // 社名 → 色（同じ社名が複数あれば先に記録されたもの）
  const logByKey = new Map();
  colorLog.forEach(l => {
    const refRow = allData[l.rowNum - 1];
    if (refRow && !logByKey.has(refRow[0])) logByKey.set(refRow[0], l);
  });

  const enhancedRows = allData.slice(1).map(row => {
    const dateC = row[2] instanceof Date ? row[2] : null;
    const dateD = row[3] instanceof Date ? row[3] : null;
    const isOld = (dateC && dateC < thresholdDate) || (dateD && dateD < thresholdDate);
    const log = logByKey.get(row[0]);
    return {
      values: row,
      color: log ? (isOld ? log.darkColor : log.color) : "#FFFFFF",
      weight: log ? (isOld ? 2 : 1) : 0,
      sortKey: isOld ? Math.min(dateC?.getTime() || Infinity, dateD?.getTime() || Infinity) : today.getTime()
    };
  });

  enhancedRows.sort((a, b) => b.weight - a.weight || a.sortKey - b.sortKey);

  const width = header.length;
  const finalData = [header, ...enhancedRows.map(r => r.values)];
  const finalColors = [Array(width).fill("#FFFFFF"), ...enhancedRows.map(r => Array(width).fill(r.color))];

  targetSheet.getRange(1, 1, finalData.length, width).setValues(finalData);
  targetSheet.getRange(1, 1, finalColors.length, width).setBackgrounds(finalColors);
}

// ===== ログ強化設定（ログシート削除・ドライラン廃止版） =====
// ※「デバッグログ」シートに書き込む機能を完全撤去し、
//   LOG は Apps Script の Logger/console への出力のみ行う簡易版に変更。
//   SAFE_MODE（ドライラン）は廃止し、常に実書き込みを行う。

const LOG = {
  info: (p, m, c, d) => _log_('INFO', p, m, c, d),
  warn: (p, m, c, d) => _log_('WARN', p, m, c, d),
  error:(p, m, c, d) => _log_('ERROR',p, m, c, d),
};

function _log_(level, place, message, countOrKey, detail){
  try{
    const line = `[${level}] ${place} - ${message}` + (countOrKey != null ? ` (${countOrKey})` : '');
    Logger.log(line);
    try { console.log(line, detail || ''); } catch(e){}
  }catch(e){}
}

// ===== 既存定数（私用コピーのみ参照版） =====
const DEST_SHEET_NAME    = '掲載状況';
const MASTER_SHEET_NAME  = '営業先リスト/現S';
const EXCLUDE_PLANS      = ['ingプラス','D-mini','Indeed PLUS 運用費','サポートプラン','ダイレクトリー'];
const CACHE_TTL_SEC      = 21600;

// ★私用コピー（指定URL）
const USER_COPY_SPREADSHEET_ID = '18-wJd6U_ABXpd40udVlbts_JVbecOHjrK2uvExGzmRI';
const USER_COPY_SHEET_NAME     = 'phase③申込';

const addDays   = (d,n)=>{const r=new Date(d);r.setDate(r.getDate()+n);r.setHours(0,0,0,0);return r;};
const addMonths = (d,n)=>{const r=new Date(d);r.setMonth(r.getMonth()+n);r.setHours(0,0,0,0);return r;};
const addYears  = (d,n)=>{const r=new Date(d);r.setFullYear(r.getFullYear()+n);r.setHours(0,0,0,0);return r;};
const hyperlink =(id)=> id ? `=HYPERLINK("https://ats.rct.airwork.net/agency/a/clients/${id}","管理画面")` : '';

// === 営業先マスタ：会社名→管理番号 ===
function getCompanyIdMap() {
  const cache   = CacheService.getScriptCache();
  const cached  = cache.get('COMPANY_ID_MAP');
  if (cached) return JSON.parse(cached);

  const sh = SpreadsheetApp.getActive().getSheetByName(MASTER_SHEET_NAME);
  if (!sh) throw new Error('営業先マスタが無いぞい');

  const lastRow = sh.getLastRow();
  const data    = lastRow>=2 ? sh.getRange(2,1,lastRow-1,16).getValues() : [];
  const map     = {};
  data.forEach((r)=>{
    const name = String(r[0]).trim(); // A列：社名
    const id   = String(r[15]).trim();// P列：管理番号
    if(name&&id) map[name]=id;
  });
  LOG.info('getCompanyIdMap','マッピング構築', Object.keys(map).length);
  cache.put('COMPANY_ID_MAP', JSON.stringify(map), CACHE_TTL_SEC);
  return map;
}

// === 私用コピーからの取得（唯一のデータソース） ===
// ※ここ自体は try/catch しない。失敗時は例外をそのまま上に投げる。
function fetchCyclesFromUserCopy_(){
  const place = 'fetchCyclesFromUserCopy_';
  const srcSS = SpreadsheetApp.openById(USER_COPY_SPREADSHEET_ID);
  const sheet = srcSS.getSheetByName(USER_COPY_SHEET_NAME);
  if(!sheet) throw new Error('私用コピーの「phase③申込」シートが見つからん');

  const lastRow  = sheet.getLastRow();
  LOG.info(place,'lastRow', lastRow);
  if (lastRow < 2) return [];

  const headers  = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0]
                    .map(h=>String(h).replace(/\s+/g,'').trim());
  const idx      = {
    company: headers.indexOf('顧客名'),
    plan   : headers.findIndex(h=>h==='媒体（中項目）'||h==='媒体(中項目)'),
    start  : headers.indexOf('掲載日（運用開始・納品日）'),
    end    : headers.indexOf('運用終了日'),
    status : headers.indexOf('受注')
  };
  LOG.info(place, 'headerIndex', '', idx);
  ['company','plan','start'].forEach(k=>{ if(idx[k]===-1) throw new Error(`ヘッダー「${k}」列が見つからん（私用コピー）`); });

  const range = sheet.getRange(2,1,lastRow-1,headers.length);
  const data  = range.getValues();
  const today = new Date(); today.setHours(0,0,0,0);
  const idMap = getCompanyIdMap();
  const res   = [];

  let skipped = { not受注:0, excludedPlan:0, noStart:0, futureStart:0 };
  for (let i=0;i<data.length;i++){
    const r = data[i];
    const plan = String(r[idx.plan]||'').trim();
    if (EXCLUDE_PLANS.includes(plan)) { skipped.excludedPlan++; continue; }

    if (idx.status>=0 && String(r[idx.status]).trim()!=='受注') { skipped.not受注++; continue; }

    const start = r[idx.start];
    if(!(start instanceof Date)) { skipped.noStart++; continue; }
    if(start>today) { skipped.futureStart++; continue; }

    const end   = r[idx.end] instanceof Date ? r[idx.end] : null;
    const comp  = String(r[idx.company]||'').trim();
    res.push({company:comp,id:idMap[comp]||'',start,end});
  }

  LOG.info(place,'取得件数（有効）', res.length, {skipped});
  if (res.length) LOG.info(place,'先頭3件サンプル','', res.slice(0,3));
  return res.sort((a,b)=>b.start-a.start);
}
// === 掲載サイクル → 行配列 ===
function buildRowsFromCycles(cycles){
  const place='buildRowsFromCycles';
  LOG.info(place,'入力件数', cycles.length);

  const today = new Date(); today.setHours(0,0,0,0);
  const latest={}
  cycles.forEach(c=>{
    if(!latest[c.company] || c.start>latest[c.company].start) latest[c.company]=c;
  });

  const rows=[];
  const msDay = 86400000;
  const companies = Object.keys(latest);
  LOG.info(place,'企業数（最新サイクルのみ）', companies.length);

  for(const comp of companies){
    const {id,start:s,end:e}=latest[comp];

    const push=(st,dt)=>rows.push([id,comp,st,dt,hyperlink(id)]);
    // ※第4要素は Date オブジェクトのまま格納し、書き込み時に日付フォーマットを持たせる

    // 掲載開始
    push('掲載開始',s);

    if(e && e>s){
      // 中間
      const mid=new Date((s.getTime()+e.getTime())/2); mid.setHours(0,0,0,0);
      if(mid<=today) push('中間',mid);

      // 終了1週間前
      const minus7=addDays(e,-7);
      if(minus7<=today) push('掲載終了1週間前',minus7);

      // 終了日
      if(e<=today)      push('掲載終了',e);
    }

    const ended=e instanceof Date && e<=today;

    if(!ended){
      // 掲載中：毎週
      const weeks=Math.floor((today-s)/(7*msDay));
      for(let n=1;n<=weeks;n++){
        const dt=addDays(s,n*7);
        const nearEnd = e && Math.abs(dt-e)<=4*msDay;
        const near7   = e && Math.abs(dt-addDays(e,-7))<=4*msDay;
        if(!(nearEnd||near7)) push(`掲載${n}週間`,dt);
      }
    }else{
      // 掲載終了後フォロー
      const w1 = addDays(e, 7);
      if (w1 <= today) push('掲載終了から1週間', w1);
      const w2 = addDays(e, 14);
      if (w2 <= today) push('掲載終了から2週間', w2);

      // 終了後1〜10か月（1〜6か月は毎月、7以降は偶数月）
      for(let m=1;m<=10;m++){
        if (m <= 6 || m % 2 === 0) {
          const dm = addMonths(e,m);
          if(dm>today) break;
          push(`掲載終了${m}か月`, dm);
        }
      }

      // 1年
      const y=addYears(e,1);
      if(y<=today) push('掲載終了1年',y);

      // 14〜24か月を偶数月
      for(let m=14; m<=24; m+=2){
        const dm = addMonths(e,m);
        if(dm>today) break;
        push(`掲載終了${m}か月`, dm);
      }
    }
  }

  LOG.info(place,'生成行数（raw）', rows.length);
  if (rows.length) LOG.info(place,'先頭5行サンプル','', rows.slice(0,5));

  // 日付列（index 3: Date）で降順ソート
  return rows.sort((a,b)=>b[3]-a[3]);
}

// === A, C, D列が一致する行を1件だけ残す ===
function filterByUniqueACD(rows){
  const place='filterByUniqueACD';
  const seen = new Set();
  const filtered = [];
  for (const row of rows) {
    const key = makeKey_ACD(row[0], row[1], row[2], row[3]) + '|' + String(row[1] || '').trim();
    if (!seen.has(key)) {
      seen.add(key);
      filtered.push(row);
    }
  }
  LOG.info(place,'重複除去', `${rows.length} → ${filtered.length}`);
  return filtered;
}

// === B列を除いたチェック復元用キー（A,C,D） ===
var publicationKeyTimeZone_;
var publicationDateKeys_ = new Map();
function makeKey_ACD(id, name, status, dateVal) {
  if (!publicationKeyTimeZone_) publicationKeyTimeZone_ = Session.getScriptTimeZone();
  const d = dateVal instanceof Date ? dateVal : new Date(dateVal);
  const stamp = d.getTime();
  if (!Number.isFinite(stamp)) throw new Error('掲載状況: 不正な日付');
  if (!publicationDateKeys_.has(stamp)) publicationDateKeys_.set(stamp, Utilities.formatDate(d, publicationKeyTimeZone_, 'yyyy/MM/dd'));
  return [String(id || '').trim(), String(name || '').trim(), status, publicationDateKeys_.get(stamp)].join('|');
}


// === チェックボックス安定化（実データ行だけに絞る軽量版） ===
// ※日常運用では呼ばない。列構成を変えたときなどに手動メンテ関数から呼ぶ想定。
function ensureCheckboxColumns_(sheet, fromRow, maxRows){
  const startRow = Math.max(2, fromRow || 2);

  // シートの「本当の最終行」を見る
  const lastRow = sheet.getLastRow();
  if (lastRow < startRow) {
    LOG.warn('ensureCheckboxColumns_','データ行がないので何もしない');
    return;
  }

  // maxRows が指定されていれば上限として使う（指定なしなら最終行まで）
  const endRow = maxRows
    ? Math.min(lastRow, startRow + maxRows - 1)
    : lastRow;

  if (endRow < startRow) {
    LOG.warn('ensureCheckboxColumns_','範囲が逆転しているので何もしない');
    return;
  }

  const rows = endRow - startRow + 1;
  const cols = 3; // G,H,I

  const range = sheet.getRange(startRow, 7, rows, cols);
  const rule  = SpreadsheetApp.newDataValidation()
    .requireCheckbox()
    .setAllowInvalid(false)
    .build();

  range.setDataValidation(rule);
  LOG.info('ensureCheckboxColumns_','設定範囲', `${startRow}-${endRow} (3列)`);
}

// === 再構築本体（履歴レス・私用コピーのみ・常に本実行）===
// ※ rows.length === 0 の場合は既存シートを維持して終了（チェックもそのまま）
function rebuildPublicationStatus(){
  return withSheetLock_(() => {
    const place = 'rebuildPublicationStatus';
    const started = Date.now();
    const ss = SpreadsheetApp.getActive();
    const dest = ss.getSheetByName(DEST_SHEET_NAME);
    if (!dest) throw new Error('掲載状況 シートが見つかりません');
    const cycles = fetchCyclesFromUserCopy_();
    const rows = filterByUniqueACD(buildRowsFromCycles(cycles));
    if (!rows.length) return {written:0, keptChecks:0};
    const oldLast = dest.getLastRow();
    const old = oldLast > 1 ? dest.getRange(2,1,oldLast-1,10).getValues() : [];
    const signature = JSON.stringify(old);
    const byKey = new Map();
    const byName = new Map();
    const byId = new Map();
    const checked = v => v === true || String(v).toUpperCase() === 'TRUE';
    old.forEach(r => {
      if (!r[3]) return;
      const key = makeKey_ACD(r[0],r[1],r[2],r[3]);
      const prior = byKey.get(key);
      const saved = r.slice();
      if (prior) {
        [6,7,8].forEach(c => saved[c] = checked(prior[c]) || checked(r[c]));
        if (!saved[5]) saved[5] = prior[5];
        if (!saved[9]) saved[9] = prior[9];
      }
      byKey.set(key,saved);
      if (r[0]) {
        const ik = makeKey_ACD(r[0],'',r[2],r[3]);
        if (!byId.has(ik)) byId.set(ik,saved);
        else {
          const other = byId.get(ik);
          byId.set(ik,other && String(other[1]).trim() === String(saved[1]).trim() ? saved : null);
        }
      }
      const nk = makeKey_ACD('',r[1],r[2],r[3]);
      if (!byName.has(nk)) byName.set(nk,saved);
      else {
        const other = byName.get(nk);
        if (other && String(other[0] || '') === String(saved[0] || '')) byName.set(nk,saved);
        else byName.set(nk,null);
      }
    });
    const body = rows.map(r => {
      let prev = byKey.get(makeKey_ACD(r[0],r[1],r[2],r[3]));
      if (!prev && r[0]) prev = byId.get(makeKey_ACD(r[0],'',r[2],r[3]));
      if (!prev) {
        const named = byName.get(makeKey_ACD('',r[1],r[2],r[3]));
        if (named && (!r[0] || !named[0] || String(r[0]) === String(named[0]))) prev = named;
      }
      prev = prev || [];
      const id = r[0] || prev[0] || '';
      return [id,r[1],r[2],r[3],hyperlink(id),prev[5] || '',
        checked(prev[6]),checked(prev[7]),checked(prev[8]),prev[9] || ''];
    });
    // 計算・復元をすべて終えてから書き込む。本文の事前消去は禁止。
    const head = ['管理番号','会社名・部門名','ステータス','日付','管理画面','電話番号','電話','有効','メール','会話内容'];
    const output = [head].concat(body);
    while (output.length < oldLast) output.push(Array(10).fill(''));
    if (Date.now() - started > 240000) throw new Error('掲載状況: 時間不足のため書き込み前に停止（既存データ維持）');
    const latest = dest.getLastRow() > 1 ? dest.getRange(2,1,dest.getLastRow()-1,10).getValues() : [];
    if (JSON.stringify(latest) !== signature) throw new Error('掲載状況: 更新中の編集を検出したため停止（再実行してください）');
    // 交互のバックアップを用意し、前回分も保持する。
    const props = PropertiesService.getScriptProperties();
    const slot = props.getProperty('PUBLICATION_BACKUP_SLOT') === 'A' ? 'B' : 'A';
    const backupName = '_掲載状況_更新前' + slot;
    let backup = ss.getSheetByName(backupName);
    if (!backup) backup = ss.insertSheet(backupName);
    const snapshot = [head].concat(old);
    const backupHeight = Math.max(snapshot.length, backup.getLastRow());
    while (snapshot.length < backupHeight) snapshot.push(Array(10).fill(''));
    if (backup.getMaxRows() < backupHeight) backup.insertRowsAfter(backup.getMaxRows(), backupHeight-backup.getMaxRows());
    backup.getRange(1,1,backupHeight,10).setValues(snapshot);
    backup.getRange('L1').setValue(new Date());
    backup.hideSheet();
    SpreadsheetApp.flush();
    // 退避完了後に次回用スロットを記録。失敗時にも退避データは残る。
    props.setProperty('PUBLICATION_BACKUP_SLOT',slot);
    if (Date.now() - started > 270000) throw new Error('掲載状況: 退避後に時間不足で停止（既存データ維持）');
    if (dest.getMaxRows() < output.length) dest.insertRowsAfter(dest.getMaxRows(), output.length-dest.getMaxRows());
    dest.getRange(1,1,output.length,10).setValues(output);
    SpreadsheetApp.flush();
    const actual = dest.getRange(2,7,body.length,3).getValues();
    if (JSON.stringify(actual) !== JSON.stringify(body.map(r=>r.slice(6,9)))) {
      throw new Error('掲載状況: チェック検証不一致。復元元=' + backupName);
    }
    LOG.info(place,'本文・チェック一括更新完了',body.length);
    dest.getRange(2,4,body.length,1).setNumberFormat('yyyy/mm/dd');
    ensureCheckboxColumns_(dest,2,body.length);
    dest.getRange(2,10,body.length,1).clearDataValidations();
    dest.getRange(2,1,body.length,10).setBackgrounds(body.map(r=>Array(10).fill(r[7] ? '#D3D3D3' : null)));
    return {written:body.length,keptChecks:body.reduce((n,r)=>n+r.slice(6,9).filter(Boolean).length,0),backup:backupName};
  });
}
// === 電話番号・会社名補完（管理番号 or 社名で）===
function fillF列_from営業先マスタ(){
  const place='fillF列_from営業先マスタ';
  const ss  = SpreadsheetApp.getActive();
  const st  = ss.getSheetByName(DEST_SHEET_NAME);
  const mst = ss.getSheetByName(MASTER_SHEET_NAME);
  if (!st || !mst) throw new Error('必要なシートが見つからんのう…');

  const stLastRow = st.getLastRow();
  if (stLastRow < 2){ LOG.warn(place,'対象行なし'); return; }
  const stData = st.getRange(2, 1, stLastRow - 1, 6).getValues();

  const ids      = stData.map(r => String(r[0]).trim());
  const curNames = stData.map(r => String(r[1]).trim());
  const curTels  = stData.map(r => String(r[5]).trim());

  const mstLastRow = mst.getLastRow();
  if (mstLastRow < 2) { LOG.warn(place,'マスタ行なし'); return; }
  const mstData = mst.getRange(2, 1, mstLastRow - 1, 16).getValues();

  const idMap = {};    // 管理番号 → { cmp, tel }
  const nameMap = {};  // 社名 → { cmp, tel }

  mstData.forEach(r => {
    const name = String(r[0]).trim();    // A
    const tel  = String(r[1]).trim();    // B
    const id   = String(r[15]).trim();   // P
    if (id)   idMap[id] = { cmp: name, tel };
    if (name) nameMap[name] = { cmp: name, tel };
  });

  const bRepl = [];
  const fRepl = [];

  for (let i = 0; i < ids.length; i++) {
    const id       = ids[i];
    const curName  = curNames[i];
    const curTel   = curTels[i];
    const byId     = idMap[id];
    const byName   = nameMap[curName];
    const target   = byId || byName || {};
    bRepl.push([curName || target.cmp || '']);  // B列
    fRepl.push([target.tel || curTel]);   // F列
  }

  st.getRange(2, 2, bRepl.length, 1).setValues(bRepl);
  st.getRange(2, 6, fRepl.length, 1).setValues(fRepl);
  LOG.info(place,'補完書き込み完了', bRepl.length);
}

// === 実行関数 ===
function 実行_日次更新(){
  const res = rebuildPublicationStatus();
  fillF列_from営業先マスタ();
  return res;
}
function 実行_全件再構築(){ return 実行_日次更新(); }

// === 手動メンテ用：チェックボックス定義を張り直すだけ ===
// ※列構成を変えたときや初期導入時に一度だけ実行すればOK。
function 一度だけ_チェックボックス再設定(){
  const ss   = SpreadsheetApp.getActive();
  const dest = ss.getSheetByName(DEST_SHEET_NAME);
  if (!dest) throw new Error('掲載状況 シートが見つからんぞい');
  ensureCheckboxColumns_(dest, 2, 10000);
}

/* =========================================
   設定
========================================= */

const PHASE1_TIME_LIMIT_MS = 4 * 60 * 1000; // 4分で自主停止
const PHASE1_LOG_SHEET = '実行ログ';
const PHASE1_HEADER_ROWS = 1;

const ENABLE_EXTERNAL_HIGHLIGHT = true;     // 外部照合を使う
     // まずは false 推奨
      // 誤爆＆爆重回避のため false 推奨
          // prefix許可時の最小長

const EXTERNAL_HIGHLIGHT_MAX_SHEETS = 3;    // 月シート走査数の上限
const EXTERNAL_HIGHLIGHT_MAX_ROWS_PER_SHEET = 3000; // 1シート最大走査行数

const HL_SOURCE_SPREADSHEET_ID = '1uFDME4QOUdJifxh80PmUWVapqc9AXu2lauKhgayU0eM';
const HL_SOURCE_COL_INDEX = 3;
const HL_COLOR = '#ffcccc';
const HL_SHEET_NAME_REGEX = /^\d{6}月$/;

const YELLOW_STORE_COLOR   = '#fff9c4';
const PINK_BAITORU_ONLY    = '#ffd1e6';
const GREEN_DEFAULT_FILL   = '#e5ffe5';
const LAVENDER_FIXED_COLOR = '#e6e6fa';
const DUP_PHONE_COLOR      = '#ffe6cc';

const _CORE_NAME_CACHE = new Map();


/* =========================================
   実行ログ
========================================= */

function writeExecLog_(message, level) {
  level = level || 'INFO';
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(PHASE1_LOG_SHEET);
  if (!sh) {
    sh = ss.insertSheet(PHASE1_LOG_SHEET);
    sh.getRange(1, 1, 1, 4).setValues([['日時', 'レベル', '関数', '内容']]);
    sh.setFrozenRows(1);
  }
  sh.appendRow([new Date(), level, 'runAllProcesses_Phase1', message]);
}

function logStepTime_(label, fn) {
  const s = Date.now();
  writeExecLog_('開始: ' + label);
  const result = fn();
  const sec = ((Date.now() - s) / 1000).toFixed(3);
  writeExecLog_('終了: ' + label + ' / ' + sec + '秒');
  return result;
}

function checkTimeLimit_(startMs, limitMs, label) {
  if (Date.now() - startMs > limitMs) {
    throw new Error('TIME_GUARD: ' + label + ' の途中で自主停止');
  }
}


/* =========================================
   ロック
========================================= */
function withSheetLock_(fn) {
  const lock = LockService.getScriptLock();
  let wait = 200;
  const maxAttempts = 4;

  for (let i = 1; i <= maxAttempts; i++) {
    if (lock.tryLock(1000)) {
      try {
        return fn();
      } finally {
        try { lock.releaseLock(); } catch (e) {}
      }
    }
    Utilities.sleep(wait);
    wait = Math.min(wait * 2, 1500);
  }

  writeExecLog_('ロック取得失敗のため今回実行をスキップ', 'WARN');
  return;
}


/* =========================================
   Phase1 本体
========================================= */

function runAllProcesses_Phase1() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(3000)) {
    writeExecLog_('スキップ: DocumentLock取得失敗', 'WARN');
    return;
  }

  const startMs = Date.now();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet(); // これを追加

    writeExecLog_('Phase1 開始');

    logStepTime_('掲載開始顧客クリーニング', () => {
      cleanKeisaiKaishiKokyaku_Improved({
        timeLimitStartMs: startMs,
        timeLimitMs: PHASE1_TIME_LIMIT_MS
      });
    });

    checkTimeLimit_(startMs, PHASE1_TIME_LIMIT_MS, 'Phase1');

    logStepTime_('架電記録更新', () => {
      updateKadenRecordSheet(ss); // ここを修正
    });

    const totalSec = ((Date.now() - startMs) / 1000).toFixed(3);
    writeExecLog_('Phase1 完了 / ' + totalSec + '秒');
  } catch (e) {
    writeExecLog_('Phase1 異常終了: ' + (e && e.stack ? e.stack : e), 'ERROR');
    throw e;
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}


/* =========================================
   cleanKeisaiKaishiKokyaku 改善版
========================================= */

function cleanKeisaiKaishiKokyaku() {
  cleanKeisaiKaishiKokyaku_Improved({
    timeLimitStartMs: Date.now(),
    timeLimitMs: PHASE1_TIME_LIMIT_MS
  });
}

function cleanKeisaiKaishiKokyaku_Improved(opt) {
  opt = opt || {};
  const startMs = opt.timeLimitStartMs || Date.now();
  const limitMs = opt.timeLimitMs || PHASE1_TIME_LIMIT_MS;

  const SS = SpreadsheetApp.getActive();
  const SH_NAME = '掲載開始顧客';
  const NG_NAME = 'NGリスト';
  const HEADER_ROWS = 1;
  const PHONE_COL = 2;
  const ADDRESS_COL = 3;

  const sh = SS.getSheetByName(SH_NAME);
  const ngSh = SS.getSheetByName(NG_NAME);
  if (!sh) throw new Error('対象シートが見当たらん: 掲載開始顧客');

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();

  if (lastRow <= HEADER_ROWS) {
    ensureHeaderRow_(sh, HEADER_ROWS);
    addCheckboxesFGH_(sh, HEADER_ROWS);
    applyBoldLeftAlign_(sh, HEADER_ROWS);
    if (ENABLE_EXTERNAL_HIGHLIGHT) {
      safeHighlightMatchedCustomers_(sh, startMs, limitMs);
    }
    try { sh.getDataRange().setFontWeight('bold'); } catch (e) {}
    return;
  }

  checkTimeLimit_(startMs, limitMs, 'cleanKeisaiKaishiKokyaku 初期');

  resetDataRowStyle_(sh, HEADER_ROWS);

  const dataRows = lastRow - HEADER_ROWS;
  const rng = sh.getRange(HEADER_ROWS + 1, 1, dataRows, lastCol);
  const values = rng.getValues();

  checkTimeLimit_(startMs, limitMs, 'values取得後');

  const aeCols = Math.min(5, lastCol);
  const rngAE = sh.getRange(HEADER_ROWS + 1, 1, dataRows, aeCols);
  const formulasAE = rngAE.getFormulas();
  const richAE = rngAE.getRichTextValues();
  const displaysAE = rngAE.getDisplayValues();

  for (let i = 0; i < dataRows; i++) {
    if (i % 200 === 0) checkTimeLimit_(startMs, limitMs, 'A〜Eリンク剥がし');

    for (let j = 0; j < aeCols; j++) {
      const f = formulasAE[i][j] || '';
      if (/^=HYPERLINK\(/i.test(f)) {
        values[i][j] = displaysAE[i][j];
        continue;
      }
      const rt = richAE[i][j];
      if (rt) {
        try {
          const runs = rt.getRuns ? (rt.getRuns() || []) : [];
          if (rt.getLinkUrl() || runs.some(r => r.getLinkUrl())) {
            values[i][j] = rt.getText();
          }
        } catch (e) {}
      }
    }
  }

  const ngSet = new Set(
    ngSh && ngSh.getLastRow() > 0
      ? ngSh.getRange(1, 1, ngSh.getLastRow(), 1).getDisplayValues()
          .map(r => (r[0] + '').trim()).filter(Boolean)
      : []
  );

  const newSet = new Set();
  const existingSet = new Set();

  [
    { name: '営業先リスト/新規', set: newSet },
    { name: '営業先リスト/現S', set: existingSet }
  ].forEach(({ name, set }) => {
    const s = SS.getSheetByName(name);
    if (s && s.getLastRow() > 1) {
      s.getRange(2, 1, s.getLastRow() - 1, 1).getDisplayValues().forEach(r => {
        const v = (r[0] + '').trim();
        if (v) set.add(v);
      });
    }
  });

  const phoneCountMap = buildPhoneCountMap_(values, PHONE_COL);
  const seenFull = new Set();
  const dupAddr = new Map();

  const out = [];
  const rowsBlue = [];
  const rowsYellow = [];
  const rowsPink = [];
  const rowsLavender = [];
  const rowsOrange = [];
  const rowsDupPhone = [];

let abortedByTimeGuard = false;

for (let i = 0; i < values.length; i++) {
  if (i % 200 === 0 && isTimeLimitReached_(startMs, limitMs)) {
    writeExecLog_('TIME_GUARD: メインループを中断 i=' + i, 'WARN');
    abortedByTimeGuard = true;
    break;
  }

  const row = values[i];

    if (!row[1] && row[0] && String(row[0]).includes(' ')) {
      const parts = String(row[0]).split(/\s+/);
      if (parts.length >= 2) {
        row[0] = parts[0];
        row[1] = parts.slice(1).join(' ');
      }
    }

    const aRaw = (row[0] || '').toString().trim();
    const aCore = normalizeCompanyCore_(aRaw);
    const dVal = (row[3] || '').toString().trim();

    if (/^[@＠]type$/i.test(dVal)) continue;
    if (ngSet.has(aRaw)) continue;

    const addrClean = cleanAddress(row[ADDRESS_COL - 1]);
    row[ADDRESS_COL - 1] = addrClean;

    const phoneFormatted = formatPhoneNumber(String(row[PHONE_COL - 1] || ''));
    row[PHONE_COL - 1] = phoneFormatted;
    const phoneKey = normalizePhoneForKey_(phoneFormatted || values[i][PHONE_COL - 1]);

    let isDelete = false;
    let appliedColor = null;

    {
      const phoneRaw = (row[1] || '').toString();
      const addrRaw = (row[2] || '').toString();

      if (!isDelete && phoneRaw && addrRaw) {
        const digits = phoneRaw.replace(/[^0-9]/g, '');
        let prefix = null;

        if (/^03/.test(digits)) prefix = '03';
        else if (/^042/.test(digits)) prefix = '042';
        else if (/^048/.test(digits)) prefix = '048';

        if (!prefix) {
          isDelete = true;
        } else {
          const AREA_RULES = {
            '03':  ['東京', '東京都'],
            '042': ['東京', '東京都', '神奈川', '神奈川県'],
            '048': ['埼玉', '埼玉県'],
          };
          const allowedWords = AREA_RULES[prefix] || [];
          const matched = allowedWords.some(w => addrRaw.includes(w));
          if (!matched) isDelete = true;
        }
      }
    }

    const colD = (row[3] || '').toString().trim();
    const colE = (row[4] || '').toString().trim();
    const fullKey = JSON.stringify([
      (row[0] || '').toString().trim(),
      phoneKey,
      addrClean,
      colD,
      colE
    ]);

    if (fullKey && seenFull.has(fullKey)) continue;

    const phoneText = (row[PHONE_COL - 1] || '').toString().trim();
    if (addrClean) {
      let bucket = dupAddr.get(addrClean);
      if (!bucket) {
        bucket = { phone: new Set(), core: new Set() };
        dupAddr.set(addrClean, bucket);
      }
      if ((phoneText && bucket.phone.has(phoneText)) || (aCore && bucket.core.has(aCore))) {
        isDelete = true;
      } else {
        if (phoneText) bucket.phone.add(phoneText);
        if (aCore) bucket.core.add(aCore);
      }
    }

    if (!isDelete) {
      const d2 = colD;
      const e2 = colE;
      if (d2 === 'ワークポート' && !e2) isDelete = true;
    }

    if (!isDelete && newSet.has(aRaw)) {
      isDelete = true;
    } else if (!isDelete && existingSet.has(aRaw)) {
      appliedColor = '#d6eaff';
    }

    if (!isDelete && !phoneText) isDelete = true;
    if (!isDelete && !colD && !colE) isDelete = true;

    if (isDelete) continue;

    if (!appliedColor && phoneKey && (phoneCountMap.get(phoneKey) || 0) >= 2) {
      appliedColor = DUP_PHONE_COLOR;
    }

    const colL = (row[11] || '').toString().trim();
    if (!appliedColor && colL) {
      appliedColor = LAVENDER_FIXED_COLOR;
    }

    const companyDisp = (row[0] || '').toString().trim();
    if (!appliedColor && /(店|校)/.test(companyDisp)) {
      appliedColor = YELLOW_STORE_COLOR;
    }

    if (!appliedColor) {
      const onlyBaitoru = /^\s*バイトル\s*$/.test(colE);
      const noLegal = !hasLegalCompanyMarker_(companyDisp);
      const hasMizu = containsMizuShobaiKeyword_(companyDisp);
      if ((!colD && onlyBaitoru && noLegal) || hasMizu) {
        appliedColor = PINK_BAITORU_ONLY;
      }
    }

    if (!appliedColor && companyDisp) {
      appliedColor = GREEN_DEFAULT_FILL;
    }

    out.push(row);
    const newRowIndex = HEADER_ROWS + out.length;

    if (appliedColor === '#d6eaff') rowsBlue.push(newRowIndex);
    else if (appliedColor === YELLOW_STORE_COLOR) rowsYellow.push(newRowIndex);
    else if (appliedColor === PINK_BAITORU_ONLY) rowsPink.push(newRowIndex);
    else if (appliedColor === LAVENDER_FIXED_COLOR) rowsLavender.push(newRowIndex);
    else if (appliedColor === GREEN_DEFAULT_FILL) rowsOrange.push(newRowIndex);
    else if (appliedColor === DUP_PHONE_COLOR) rowsDupPhone.push(newRowIndex);

    seenFull.add(fullKey);
  }

if (abortedByTimeGuard) {
  writeExecLog_('途中終了のため今回は書き戻ししない', 'WARN');
  return;
}

withSheetLock_(() => {

    sh.getRange(HEADER_ROWS + 1, 1, dataRows, lastCol).clearContent();

    if (out.length > 0) {
      const newR = sh.getRange(HEADER_ROWS + 1, 1, out.length, lastCol);
      newR.setValues(out).setFontColor('#000000');

      if (rowsBlue.length)     sh.getRangeList(rowsBlue.map(r => `${r}:${r}`)).setBackground('#d6eaff');
      if (rowsYellow.length)   sh.getRangeList(rowsYellow.map(r => `${r}:${r}`)).setBackground(YELLOW_STORE_COLOR);
      if (rowsPink.length)     sh.getRangeList(rowsPink.map(r => `${r}:${r}`)).setBackground(PINK_BAITORU_ONLY);
      if (rowsLavender.length) sh.getRangeList(rowsLavender.map(r => `${r}:${r}`)).setBackground(LAVENDER_FIXED_COLOR);
      if (rowsOrange.length)   sh.getRangeList(rowsOrange.map(r => `${r}:${r}`)).setBackground(GREEN_DEFAULT_FILL);
      if (rowsDupPhone.length) sh.getRangeList(rowsDupPhone.map(r => `${r}:${r}`)).setBackground(DUP_PHONE_COLOR);

      try { sh.getRange(HEADER_ROWS + 1, PHONE_COL, out.length, 1).breakApart(); } catch (e) {}
      sh.getRange(HEADER_ROWS + 1, PHONE_COL, out.length, 1).setNumberFormat('@');
    }
  });

  checkTimeLimit_(startMs, limitMs, '補助処理前');

  addCheckboxesFGH_(sh, HEADER_ROWS);
  ensureHeaderRow_(sh, HEADER_ROWS);
  setJobMediaLinks_Improved();
  applyBoldLeftAlign_(sh, HEADER_ROWS);

  if (ENABLE_EXTERNAL_HIGHLIGHT) {
    safeHighlightMatchedCustomers_(sh, startMs, limitMs);
  }

  // 色順に並び替え
  sortColoredRowsToTop_Light_(sh, HEADER_ROWS);

  try { sh.getDataRange().setFontWeight('bold'); } catch (e) {}
}


/* =========================================
   setJobMediaLinks 改善版
========================================= */

function setJobMediaLinks_Improved() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('掲載開始顧客');
  if (!sh) throw new Error('シート「掲載開始顧客」が見つからんぞい');

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  const values = sh.getRange(2, 1, lastRow - 1, 5).getValues();
  const curDE = sh.getRange(2, 4, lastRow - 1, 2).getRichTextValues();

  const newD = [];
  const newE = [];
  let dirty = false;

  for (let i = 0; i < values.length; i++) {
    const [company,, address, colD, colE] = values[i];
    const nameRaw = (company || '').toString().trim();
    const q = encodeURIComponent(nameRaw);
    const qjbPath = encodeURIComponent(nameRaw + 'の仕事');

    let urlD = null;
    let urlE = null;
    const textD = (colD || '').toString();
    const textE = (colE || '').toString();

    if (textD) {
      if (textD.includes('マイナビ')) {
        urlD = `https://tenshoku.mynavi.jp/list/kw${q}/?jobsearchType=4&searchType=8`;
      } else if (textD.includes('工場ワークス')) {
        urlD = `https://04510.jp/jobs/areas/kanto/?keyword=${q}`;
      } else if (/DODA/i.test(textD)) {
        urlD = `https://doda.jp/DodaFront/View/JobSearchList.action?k=${q}&kwc=1&ss=1&pic=1&ds=0&tp=1&bf=1&mpsc_sid=10&oldestDayWdtno=0&leftPanelType=1&usrclk_searchList=PC-logoutJobSearchList_searchConditionArea_searchButton-kwdInclude`;
      } else if (textD.includes('エン転職')) {
        urlD = `https://employment.en-japan.com/search/search_list/?keywordtext=${q}%A4&aroute=0`;
      } else if (textD.includes('求人ボックス')) {
        urlD = `https://xn--pckua2a7gp15o89zb.com/${qjbPath}`;
      }
    }

    if (textE) {
      if (textE.includes('Indeed PLUS')) {
        const pref = extractPrefecture(String(address || ''));
        urlE = `https://jp.indeed.com/jobs?q=${q}&l=${encodeURIComponent(pref || '')}`;
      } else if (textE.includes('バイトル')) {
        urlE = `https://www.baitoru.com/kw/${q}/`;
      } else if (/DODA/i.test(textE)) {
        urlE = `https://doda.jp/DodaFront/View/JobSearchList.action?k=${q}&kwc=1&ss=1&pic=1&ds=0&tp=1&bf=1&mpsc_sid=10&oldestDayWdtno=0&leftPanelType=1&usrclk_searchList=PC-logoutJobSearchList_searchConditionArea_searchButton-kwdInclude`;
      } else if (textE.includes('エン転職')) {
        urlE = `https://employment.en-japan.com/search/search_list/?keywordtext=${q}%A4&aroute=0`;
      } else if (textE.includes('求人ボックス')) {
        urlE = `https://xn--pckua2a7gp15o89zb.com/${qjbPath}`;
      }
    }

    const dNow = curDE[i][0];
    const eNow = curDE[i][1];

    const needD = !dNow || dNow.getText() !== textD || dNow.getLinkUrl() !== (urlD || null);
    const needE = !eNow || eNow.getText() !== textE || eNow.getLinkUrl() !== (urlE || null);

    newD.push([needD ? SpreadsheetApp.newRichTextValue().setText(textD).setLinkUrl(urlD || null).build() : dNow]);
    newE.push([needE ? SpreadsheetApp.newRichTextValue().setText(textE).setLinkUrl(urlE || null).build() : eNow]);

    dirty = dirty || needD || needE;
  }

  if (dirty) {
    sh.getRange(2, 4, newD.length, 1).setRichTextValues(newD);
    sh.getRange(2, 5, newE.length, 1).setRichTextValues(newE);
  }
}


/* =========================================
   外部照合 改善版
========================================= */

function safeHighlightMatchedCustomers_(targetSheet, startMs, limitMs) {
  if (isTimeLimitReached_(startMs, limitMs)) {
    writeExecLog_('外部照合スキップ: 開始前に時間上限到達', 'WARN');
    return;
  }

  try {
    highlightMatchedCustomers_Improved(targetSheet, {
      headerRows: PHASE1_HEADER_ROWS,
      sourceSpreadsheetId: HL_SOURCE_SPREADSHEET_ID,
      sourceColIndex: HL_SOURCE_COL_INDEX,
      highlightColor: HL_COLOR,
      sheetNameRegex: HL_SHEET_NAME_REGEX,
      maxSheets: EXTERNAL_HIGHLIGHT_MAX_SHEETS,
      maxRowsPerSheet: EXTERNAL_HIGHLIGHT_MAX_ROWS_PER_SHEET,
      startMs: startMs,
      limitMs: limitMs
    });
  } catch (e) {
    if (String(e && e.message || '').indexOf('TIME_GUARD:') === 0) {
      writeExecLog_('外部照合中断: ' + e.message, 'WARN');
      return;
    }
    writeExecLog_('外部照合失敗: ' + (e && e.stack ? e.stack : e), 'WARN');
  }
}

function getExternalHighlightMap_(opt) {
  const { sourceSpreadsheetId, sourceColIndex, sheetNameRegex, maxSheets, maxRowsPerSheet, startMs, limitMs } = opt || {};

  const cache = CacheService.getScriptCache();
  const cacheKey = ['extHL2', sourceSpreadsheetId, sourceColIndex, maxSheets, maxRowsPerSheet].join(':');
  const cached = cache.get(cacheKey);
  if (cached) return JSON.parse(cached);

  if (isTimeLimitReached_(startMs, limitMs)) {
    throw new Error('TIME_GUARD: 外部照合: 辞書生成前 の途中で自主停止');
  }

  const src = SpreadsheetApp.openById(sourceSpreadsheetId);
  const monthSheets = src.getSheets()
    .filter(sh => sheetNameRegex.test(sh.getName()))
    .sort((a, b) => a.getName() < b.getName() ? 1 : -1)
    .slice(0, maxSheets);

  const exactMap = {};
  const colLetter = columnIndexToLetter_(sourceColIndex);

  monthSheets.forEach(sh => {
    if (isTimeLimitReached_(startMs, limitMs)) {
      throw new Error('TIME_GUARD: 外部照合: 月シート走査 の途中で自主停止');
    }

    const gid = sh.getSheetId();
    const rMax = Math.min(sh.getLastRow(), maxRowsPerSheet);
    if (rMax < 1) return;

    const col = sh.getRange(1, sourceColIndex, rMax, 1).getDisplayValues();
    for (let r = 0; r < rMax; r++) {
      const text = (col[r][0] ?? '').toString().trim();
      if (!text) continue;
      const aft = extractAfterFor(text);
      const core = aft && normalizeCompanyCore_(aft);
      if (!core || exactMap[core]) continue;
      const url = `https://docs.google.com/spreadsheets/d/${sourceSpreadsheetId}/edit#gid=${gid}&range=${encodeURIComponent(colLetter + (r + 1))}`;
      exactMap[core] = { url, raw: String(aft).trim() };
    }
  });

  const result = { exactMap };
  try { cache.put(cacheKey, JSON.stringify(result), 60 * 10); } catch (e) {} // 100KB超は保存しないだけ
  return result;
}

/* =========================================
   色行ソート 軽量版
========================================= */

function sortColoredRowsToTop_Light_(sheet, headerRows, opt) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow <= headerRows) return;

  const priority = (opt && opt.priorityColors) || [
    '#d6eaff',
    HL_COLOR,
    DUP_PHONE_COLOR,
    YELLOW_STORE_COLOR,
    PINK_BAITORU_ONLY,
    GREEN_DEFAULT_FILL
  ];

  const whiteSet = new Set(['#ffffff', '#fff', '', null]);
  const LAVENDER = LAVENDER_FIXED_COLOR;

  const numRows = lastRow - headerRows;
  const range = sheet.getRange(headerRows + 1, 1, numRows, lastCol);

  const values = range.getValues();
  const bgs = range.getBackgrounds();

  function rowColor(bgRow) {
    for (let c = 0; c < bgRow.length; c++) {
      const col = (bgRow[c] || '').toLowerCase();
      if (!whiteSet.has(col)) return col;
    }
    return '';
  }

  function colorScore(color) {
    if (!color) return priority.length;
    const idx = priority.indexOf(color);
    return idx >= 0 ? idx : priority.length + 1;
  }

  const rows = [];
  for (let i = 0; i < numRows; i++) {
    const rep = rowColor(bgs[i]);
    rows.push({
      fixed: rep === LAVENDER,
      score: colorScore(rep),
      orig: i,
      values: values[i],
      bg: bgs[i],
    });
  }

  const movable = rows.filter(r => !r.fixed);
  movable.sort((a, b) => (a.score !== b.score) ? (a.score - b.score) : (a.orig - b.orig));

  const outValues = new Array(numRows);
  const outBG = new Array(numRows);
  let k = 0;

  for (let i = 0; i < numRows; i++) {
    if (rows[i].fixed) {
      outValues[i] = rows[i].values;
      outBG[i] = rows[i].bg;
    } else {
      outValues[i] = movable[k].values;
      outBG[i] = movable[k].bg;
      k++;
    }
  }

  range.setValues(outValues);
  range.setBackgrounds(outBG);
}


/* =========================================
   補助関数
========================================= */

function normalizeCompanyCore_(raw) {
  if (!raw) return '';
  const key = String(raw);
  if (_CORE_NAME_CACHE.has(key)) return _CORE_NAME_CACHE.get(key);

  let s = key
    .replace(/[Ａ-Ｚａ-ｚ０-９]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xFEE0))
    .replace(/　/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

  s = s.replace(/[．・.\-\/＿_]/g, '');
  const legal = /(株式会社|（株）|\(株\)|（有）|\(有\)|有限会社|合同会社|合名会社|合資会社|一般社団法人|一般財団法人|特定非営利活動法人|ＮＰＯ法人|NPO法人|LLC|Inc\.?|Co\.?\,? ?Ltd\.?|Ltd\.?)/gi;
  s = s.replace(legal, '').replace(/\s+/g, ' ').trim();
  s = s.replace(/\(.*?\)|（.*?）/g, '').trim();
  s = s.replace(/\s+/g, '');

  const out = s.toLowerCase();
  _CORE_NAME_CACHE.set(key, out);
  return out;
}

function hasLegalCompanyMarker_(raw) {
  if (!raw) return false;
  const s = String(raw);
  const legal = /(株式会社|（株）|\(株\)|（有）|\(有\)|有限会社|合同会社|合名会社|合資会社|一般社団法人|一般財団法人|特定非営利活動法人|ＮＰＯ法人|NPO法人|LLC|Inc\.?|Co\.?\,? ?Ltd\.?|Ltd\.?)/i;
  return legal.test(s);
}

function extractAfterFor(text) {
  const normalized = String(text).replace(/　/g, ' ').replace(/\s+/g, ' ').trim();
  const lower = normalized.toLowerCase();
  const idx = lower.lastIndexOf(' for ');
  if (idx >= 0) return normalized.substring(idx + 5).trim();
  const idx2 = lower.lastIndexOf('for');
  if (idx2 >= 0) return normalized.substring(idx2 + 3).trim();
  return normalized;
}

function columnIndexToLetter_(idx) {
  let n = idx, s = '';
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function resetDataRowStyle_(sheet, headerRows) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow <= headerRows || lastCol < 1) return;
  const range = sheet.getRange(headerRows + 1, 1, lastRow - headerRows, lastCol);
  try { range.setBackground(null); } catch (e) {}
}

function extractPrefecture(addr) {
  let s = '';
  if (addr == null) s = '';
  else if (typeof addr === 'string') s = addr;
  else if (typeof addr.getText === 'function') {
    try { s = addr.getText(); } catch (e) { s = String(addr); }
  } else s = String(addr);

  if (!s) return '';

  const prefs = ['北海道','青森県','岩手県','宮城県','秋田県','山形県','福島県','茨城県','栃木県','群馬県','埼玉県','千葉県','東京都','神奈川県','新潟県','富山県','石川県','福井県','山梨県','長野県','岐阜県','静岡県','愛知県','三重県','滋賀県','京都府','大阪府','兵庫県','奈良県','和歌山県','鳥取県','島根県','岡山県','広島県','山口県','徳島県','香川県','愛媛県','高知県','福岡県','佐賀県','長崎県','熊本県','大分県','宮崎県','鹿児島県','沖縄県'];
  for (const p of prefs) if (s.includes(p)) return p;
  return '';
}

function cleanAddress(str) {
  if (!str) return '';
  let s = String(str);
  s = s.replace(/(ＴＥＬ|TEL|電話|ＦＡＸ|FAX).*$/i, '');
  s = s.replace(/[０-９]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xFEE0));
  s = s.replace(/[‐-–—−－﹣―]/g, '-');
  s = s.replace(/[（([\[]?\s*〒\s*\d{3}\s*-?\s*\d{4}\s*[）)】\]]?/g, '');
  s = s.replace(/\b\d{3}\s*-?\s*\d{4}\b/g, '');
  s = s.replace(/\b\d{7}\b/g, '');
  s = s.replace(/本社\s*[\/／]/g, '');
  s = s.replace(/^【[^】]*】\s*/g, '').replace(/^■+\s*/g, '').replace(/／\s*/g, '／ ')
       .replace(/^[^\/]*\/\s*/, '').replace(/^[^：]*：\s*/, '');

  if (!(s.includes('東京都') || s.includes('埼玉県'))) {
    const tokyoW = /^(足立区|荒川区|板橋区|江戸川区|大田区|葛飾区|北区|江東区|品川区|渋谷区|新宿区|杉並区|墨田区|世田谷区|台東区|中央区|千代田区|豊島区|中野区|練馬区|文京区|港区|目黒区)/;
    const saitama = /^(さいたま市|川越市|川口市|所沢市|越谷市|春日部市|上尾市|草加市|蕨市|戸田市|入間市|朝霞市|志木市|和光市|新座市|八潮市|富士見市|三郷市|ふじみ野市)/;
    if (tokyoW.test(s)) s = '東京都' + s;
    else if (saitama.test(s)) s = '埼玉県' + s;
  }

  s = s.replace(/丁目/g, '－').replace(/番地の/g, '－').replace(/番地/g, '－')
       .replace(/番/g, '－').replace(/号/g, '－').replace(/の/g, '－').replace(/[-－]+/g, '－');

  s = s.replace(/[（([\[]\s*[）)】\]]/g, '').replace(/\s{2,}/g, ' ')
       .replace(/^[－\s、，\.。・]+|[－\s、，\.。・]+$/g, '').trim();

  return s;
}

function formatPhoneNumber(input) {
  if (input == null) return '';
  let s = String(input).replace(/[０-９]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xFEE0));
  let digits = s.replace(/\D/g, '');
  if (!digits) return '';

  if (digits.length > 11) {
    const parts = partitionDigitsFlexible_(digits);
    if (parts) return parts.map(formatOneNumber_).join('/');
  }
  return formatOneNumber_(digits);

  function formatOneNumber_(d) {
    if (d.length === 9 && d[0] !== '0') d = '0' + d;
    if (d.length === 10 && /^(?:70|80|90|50)/.test(d)) d = '0' + d;

    if (d.length === 11) return d.slice(0, 3) + '-' + d.slice(3, 7) + '-' + d.slice(7);
    if (d.length === 10) {
      if (d.startsWith('03') || d.startsWith('06')) return d.slice(0, 2) + '-' + d.slice(2, 6) + '-' + d.slice(6);
      if (d.startsWith('0120') || d.startsWith('0570')) return d.slice(0, 4) + '-' + d.slice(4, 7) + '-' + d.slice(7);
      return d.slice(0, 3) + '-' + d.slice(3, 6) + '-' + d.slice(6);
    }
    return d;
  }
}

function partitionDigitsFlexible_(digits) {
  const out = [];
  let i = 0;
  while (i < digits.length) {
    let len = null;
    if (i + 11 <= digits.length) len = 11;
    else if (i + 10 <= digits.length) len = 10;
    else if (i + 9 <= digits.length) len = 9;
    else return null;
    out.push(digits.slice(i, i + len));
    i += len;
  }
  return out;
}

function ensureHeaderRow_(sheet, headerRows) {
  if (headerRows >= 1) {
    sheet.setFrozenRows(headerRows);
    const hdr = sheet.getRange(1, 1, 1, sheet.getLastColumn());
    hdr.setFontWeight('bold');
    if (!sheet.getRange(1, 1).getValue()) sheet.getRange(1, 1).setValue('社名');
    if (!sheet.getRange(1, 2).getValue()) sheet.getRange(1, 2).setValue('電話');
    if (!sheet.getRange(1, 3).getValue()) sheet.getRange(1, 3).setValue('住所');
  }
}

function ensureMinColumns_(sheet, minCols) {
  const maxCols = sheet.getMaxColumns();
  if (maxCols < minCols) sheet.insertColumnsAfter(maxCols, minCols - maxCols);
}

function addCheckboxesFGH_(sheet, headerRows) {
  ensureMinColumns_(sheet, 11);
  const lastRow = sheet.getLastRow();
  if (lastRow <= headerRows) return;

  const numRows = lastRow - headerRows;
  const rule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  const targetCols = [6, 7, 8, 11];

  for (const col of targetCols) {
    const r = sheet.getRange(headerRows + 1, col, numRows, 1);
    const validations = Array.from({ length: numRows }, () => [rule]);
    r.setDataValidations(validations);
  }

  const headerMap = new Map([[6,'フラグ1'],[7,'フラグ2'],[8,'フラグ3'],[11,'フラグ4']]);
  for (const [col, label] of headerMap.entries()) {
    const cell = sheet.getRange(1, col);
    if (!cell.getValue()) cell.setValue(label);
  }
}

function applyBoldLeftAlign_(sheet, headerRows) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) return;

  sheet.getRange(1, 1, Math.min(headerRows, lastRow), lastCol)
    .setFontWeight('bold')
    .setHorizontalAlignment('left');

  if (lastRow > headerRows) {
    sheet.getRange(headerRows + 1, 1, lastRow - headerRows, lastCol)
      .setHorizontalAlignment('left');
  }
}

function containsMizuShobaiKeyword_(raw) {
  if (!raw) return false;

  const s = String(raw)
    .replace(/[Ａ-Ｚａ-ｚ０-９]/g, c => String.fromCharCode(c.charCodeAt(0)-0xFEE0))
    .replace(/[’´`]/g, "'")
    .toLowerCase()
    .trim();

  const KEY_PATTERNS = [
    /girls?\s*bar\b/, /girls?bar\b/,
    /ガールズ?バー/, /ガルバ/,
    /girls?\s*lounge\b/, /ガールズ?ラウンジ/,
    /bunny\s*(?:lounge|bar)\b/, /バニー(?:ラウンジ|バー)/,
    /キャバクラ/, /(?:朝|昼|熟女)キャバクラ/,
    /スナック/,
    /ニュークラブ/,
    /girls?\s*caf[eé]\b/, /ガールズ?カフェ/
  ];
  if (KEY_PATTERNS.some(re => re.test(s))) return true;

  const LOOSE_JA = [
    /ガール(?:ズ)?[^一-龠ぁ-んァ-ンa-z0-9]{0,3}バー/,
    /ガール(?:ズ)?[^一-龠ぁ-んァ-ンa-z0-9]{0,3}ラウンジ/,
    /ガール(?:ズ)?[^一-龠ぁ-んァ-ンa-z0-9]{0,3}カフェ/
  ];
  const LOOSE_EN = [
    /girls?[^a-z0-9]{0,3}bar\b/,
    /girls?[^a-z0-9]{0,3}lounge\b/,
    /girls?[^a-z0-9]{0,3}caf[eé]\b/
  ];
  if (LOOSE_JA.some(re => re.test(s)) || LOOSE_EN.some(re => re.test(s))) return true;

  const NEGATIVE = /\b(?:golf|sports?|tennis|health|fitness|school|after[-\s]?school|clubhouse|culture|カルチャー|スポーツ|テニス|フィットネス|ヘルス|少年|少女|子ども|子供)\b/;
  if (/club\b/i.test(s) && !NEGATIVE.test(s)) return true;

  return false;
}

function normalizePhoneForKey_(input) {
  if (input == null) return '';
  return String(input)
    .replace(/[０-９]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xFEE0))
    .replace(/\D/g, '');
}

function buildPhoneCountMap_(values, phoneColIndex1based) {
  const idx = phoneColIndex1based - 1;
  const map = new Map();
  for (let i = 0; i < values.length; i++) {
    const raw = values[i][idx];
    const key = normalizePhoneForKey_(raw);
    if (!key) continue;
    map.set(key, (map.get(key) || 0) + 1);
  }
  return map;
}
function isTimeLimitReached_(startMs, limitMs) {
  return (Date.now() - startMs) > limitMs;
}

function highlightMatchedCustomers_Improved(targetSheet, opt) {
  const headerRows = opt.headerRows || 1;
  const highlightColor = opt.highlightColor || '#ffcccc';

  const lastRow = targetSheet.getLastRow();
  const lastCol = targetSheet.getLastColumn();
  if (lastRow <= headerRows) return;

  const numRows = lastRow - headerRows;
  const map = getExternalHighlightMap_(opt);
  const exactMap = map.exactMap || {};

  const nameRange = targetSheet.getRange(headerRows + 1, 1, numRows, 1);
  const names = nameRange.getDisplayValues();

  for (let i = 0; i < numRows; i++) {
    const rawName = (names[i][0] || '').toString().trim();
    const core = normalizeCompanyCore_(rawName);
    const matched = core && exactMap[core];

    if (matched && matched.url) {
      const rowNo = headerRows + 1 + i;

      targetSheet.getRange(rowNo, 1).setRichTextValue(
        SpreadsheetApp.newRichTextValue()
          .setText(rawName)
          .setLinkUrl(matched.url)
          .build()
      );

      targetSheet.getRange(rowNo, 1, 1, lastCol).setBackground(highlightColor);
    }
  }
}

/***** 効果最大化・H1起点対応 完全最適化版（加重件数最大化＋道路っぽい距離 ＝ APIなし） by GAS和尚 *****/

// ≪統合先シート≫
const DEST_SHEET_MAIN = "飛び込みリスト";

// ≪パイプライン（全部このシートへ集約）≫
const PIPELINES = {
  leads:     { dest: DEST_SHEET_MAIN, sources: ["営業先リスト/新規", "営業先リスト/現S"] },
  published: { dest: DEST_SHEET_MAIN, sources: ["掲載開始顧客"] }
};

// ≪列定義（1-indexed）≫ O列(=15) はセグメント用（PUB/NEW/CUR）
const COL = { A:1,B:2,C:3,D:4,E:5,F:6,G:7,H:8,I:9,J:10,K:11,L:12,M:13,N:14,O:15 };

// ≪閾値等≫
const THRESH = {
  clusterKm: 0.5,        // クラスタ距離（km）
  clusterMinSize: 3,     // クラスタ最小件数
  recentDays: 30,        // 直近予定で黒●外し
  cacheSecs: 21600,      // 6h
  cacheKeyMax: 230
};

const AUTO_CONF = {
  speedKmPerHour: 12         // 徒歩＋小移動の想定（移動分数の表示用）
};

// === “道路っぽい距離” 近似（APIなし） ======================================
function haversineKm(lat1,lng1,lat2,lng2){
  const R=6371,rad=Math.PI/180;
  const dLat=(lat2-lat1)*rad, dLng=(lng2-lng1)*rad;
  const a=Math.sin(dLat/2)**2+Math.cos(lat1*rad)*Math.cos(lat2*rad)*Math.sin(dLng/2)**2;
  return 2*R*Math.atan2(Math.sqrt(a),Math.sqrt(1-a));
}
function llToXY_(lat, lng, lat0){
  const kmPerDegLat = 110.574; // 緯度方向
  const kmPerDegLng = 111.320 * Math.cos((lat0*Math.PI)/180); // 経度方向（緯度依存）
  return { x: lng * kmPerDegLng, y: lat * kmPerDegLat };
}
function manhattanKm_(A, B, latRef){
  const a = llToXY_(A.lat, A.lng, latRef);
  const b = llToXY_(B.lat, B.lng, latRef);
  return Math.abs(a.x - b.x) + Math.abs(a.y - b.y);
}
function calibDetourFactor_(pts){
  if(pts.length < 4) return 1.20;
  const latRef = pts.reduce((s,p)=>s+p.lat,0)/pts.length;
  const samples=[];
  for(let i=0;i<Math.min(60, pts.length-1); i++){
    const a=pts[i];
    const b=pts[(i*17+11)%pts.length];
    const hv=haversineKm(a.lat,a.lng,b.lat,b.lng); if(hv<0.05) continue;
    const mk=manhattanKm_(a,b,latRef);
    samples.push(mk/hv);
  }
  if(!samples.length) return 1.22;
  samples.sort((x,y)=>x-y);
  const med=samples[Math.floor(samples.length/2)];
  return Math.min(1.60, Math.max(1.08, med));
}
function roadishKm_(A,B,latRef,detour){
  const hv=haversineKm(A.lat,A.lng,B.lat,B.lng);
  const mk=manhattanKm_(A,B,latRef);
  const blend=0.35*mk + 0.65*hv; // 直線寄りブレンド
  return Math.max(hv,blend)*detour;
}
// ＝＝＝ 小道具 ＝＝＝
function to2dBlank(n){ return Array.from({length:n}, ()=>[""]); } // N×1空配列
function normAddr(s){ return String(s||"").replace(/　/g," ").replace(/\s+/g," ").trim(); }
function countVisits(text){
  if(!text) return 0;
  const m=String(text).match(/\d+回/g);
  return m? m.reduce((s,t)=>s+parseInt(t,10),0):0;
}

/** 単発ジオコーディング（住所→{lat,lng}） */
function geocodeOne_(address){
  const a=normAddr(address||"");
  if(!a) return null;
  try{
    const geo=Maps.newGeocoder().geocode(a);
    if(geo.results && geo.results.length>0){
      return geo.results[0].geometry.location; // {lat,lng}
    }
  }catch(e){}
  return null;
}

/** 起点の取得（H1から住所 or "lat,lng" を解釈） */
function getStartFromH1_(sheetName){
  const sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if(!sh) throw new Error("シートなし: "+sheetName);
  const v = String(sh.getRange(1, COL.H).getValue()||"").trim(); // H1
  if(!v) return null;
  const m=v.match(/^\s*(-?\d+(?:\.\d+)?)\s*,\s*(-?\d+(?:\.\d+)?)\s*$/);
  if(m){ return {lat:parseFloat(m[1]), lng:parseFloat(m[2])}; }
  const loc=geocodeOne_(v);
  if(!loc) throw new Error("起点(H1)のジオコーディングに失敗: "+v);
  return loc;
}

// ＝＝＝ 小道具（電話番号の正規化＆先頭判定） ＝＝＝
function toHanDigits_(s){
  return String(s||"").replace(/[０-９]/g, ch => String.fromCharCode(ch.charCodeAt(0)-0xFEE0));
}
function normalizeTel_(v){
  let s = toHanDigits_(String(v||"").trim());
  s = s.replace(/[^\d+]/g, "");
  if (s.startsWith("+81")) s = "0" + s.slice(3);
  return s;
}
function telAllowed_(v){
  const s = normalizeTel_(v);
  return s.startsWith("048") || s.startsWith("03") || s.startsWith("042");
}

/**
 * 1) 転記＋K復元＋仮●＋背景色＋セグメント（O=PUB/NEW/CUR）
 *    ※電話番号フィルタ（048/03/042 以外は省く）
 */
function processLeadsAndJudge_(destName, sourceNames){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const dest=ss.getSheetByName(destName)||ss.insertSheet(destName);
  const today=new Date();
  const lastRow=Math.max(dest.getLastRow(),1);

  // 既存K列のメモ（会社名→Kの値）
  const exist= lastRow>1? dest.getRange(2,COL.A,lastRow-1,COL.K).getValues():[];
  const memoK={}; exist.forEach(r=>{ if(r[0]) memoK[r[0]] = r[COL.K-1]; });

  const outRows=[], outBg=[];

  sourceNames.forEach(name=>{
    const sh=ss.getSheetByName(name); if(!sh) return;
    const data=sh.getDataRange().getValues();

    for(let i=1;i<data.length;i++){
      const r=data[i];
      const company=r[0], tel=r[1], addr=r[2], keyman=r[3];
      const nextDate=r[9], skipFlag=r[11], visitsTxt=r[13], presence=r[15];

      if(!company||!addr) continue;

      // ★電話番号フィルタ：048 / 03 / 042 以外は省く（空欄も省く）
      if(!telAllowed_(tel)) continue;

      // 新規かつ「一旦追わない」は除外
      if(name.indexOf("新規")!==-1 && skipFlag==="一旦追わない") continue;

      // 背景色：掲載開始=薄緑 / 新規=薄紅 / 現S=水色
      let bg="#e5ffe5";
      if(name.indexOf("新規")!==-1) bg="#ffe5e5";
      else if(name.indexOf("現S")!==-1) bg="#e5f0ff";

      // 仮 黒●（新規/掲載開始）
      let mark="";
      const target=(name.indexOf("新規")!==-1 || name.indexOf("掲載開始")!==-1);
      if(target){
        const hasKeyman=!!String(keyman||"").trim();
        const v=countVisits(visitsTxt);
        const low=(presence==="中"||presence==="低");
        let recent=false;
        if(nextDate instanceof Date){
          const dd=Math.floor((nextDate-today)/(1000*60*60*24));
          recent=(dd>=0 && dd<=THRESH.recentDays);
        }
        if((hasKeyman||v>=5||low)&&!recent) mark="●";
      }

      // セグメント
      let seg="CUR";
      if(name.indexOf("掲載開始")!==-1) seg="PUB";
      else if(name.indexOf("新規")!==-1) seg="NEW";

      // 出力（A..O=15列）
      outRows.push([
        company,                 // A 会社名
        normalizeTel_(tel),      // B 電話
        addr,                    // C 住所
        keyman,                  // D キーマン
        "", "",                  // E-F 緯度・経度
        "","","","",             // G-J 効果最大化用
        memoK[company]||"",      // K 過去値復元
        mark,"","",              // L/M/N 黒●、クラスタID、訪問回数テキスト等
        seg                      // O セグメント
      ]);

      // 背景（O列幅に合わせる）
      outBg.push(Array.from({length:COL.O}, (_,j)=>{
        return (j===COL.O-1)? bg : (j<COL.K? bg : "");
      }));
    }
  });

  // 既存データを消去
  if(lastRow>1) dest.getRange(2,COL.A,lastRow-1,COL.O).clearContent().setBackground(null);

  // 転記＆背景色適用
  if(outRows.length){
    dest.getRange(2,COL.A,outRows.length,COL.O).setValues(outRows);
    dest.getRange(2,COL.A,outRows.length,COL.O).setBackgrounds(outBg);
  }
}

// 2) 住所→緯度経度（Maps.Geocoder＋キャッシュ）
function fillLatLng_(sheetName){
  const sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if(!sh) throw new Error("シートなし: "+sheetName);
  const lastRow=sh.getLastRow(); if(lastRow<=1) return;

  const addr=sh.getRange(2,COL.C,lastRow-1,1).getValues();
  const latlng=sh.getRange(2,COL.E,lastRow-1,2).getValues();
  const cache=CacheService.getScriptCache();

  for(let i=0;i<addr.length;i++){
    if(!addr[i][0]) continue;
    if(latlng[i][0] && latlng[i][1]) continue;

    const a=normAddr(addr[i][0]);
    const useCache = a.length<=THRESH.cacheKeyMax;
    const key = useCache ? `geo:${a}` : null;

    let loc=null;
    if(useCache){
      const c=cache.get(key);
      if(c){ try{ loc=JSON.parse(c); }catch(e){} }
    }
    if(!loc){
      try{
        const geo=Maps.newGeocoder().geocode(a);
        if(geo.results&&geo.results.length>0){
          loc=geo.results[0].geometry.location;
          if(useCache){ try{ cache.put(key,JSON.stringify(loc),THRESH.cacheSecs);}catch(e){} }
        }else{ continue; }
        Utilities.sleep(150);
      }catch(e){ continue; }
    }
    latlng[i][0]=loc.lat; latlng[i][1]=loc.lng;
  }
  sh.getRange(2,COL.E,lastRow-1,2).setValues(latlng);
}

// 4) 並べ替え（G昇順）
function sortByG_(sheetName){
  const sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if(!sh) throw new Error("シートなし: "+sheetName);
  const lastRow=sh.getLastRow(); if(lastRow<=1) return;
  sh.getRange(2,COL.A,lastRow-1,COL.O).sort({column:COL.G, ascending:true});
}

// 5) 効率順（近い順）に「全部」並び替えるルート生成
//    segMode: "pub_new"（既定） / "all" / "cur"
function optimizeRoute_Auto_Core_(sheetName, opt, startPoint, segMode){
  const o = Object.assign({
    speedKmPerHour: 12  // 分数表示用にだけ使う
  }, opt || {});

  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sh) throw new Error("シートなし: " + sheetName);
  const lastRow = sh.getLastRow();
  if (lastRow <= 1) return;

  const N     = lastRow - 1;
  const width = COL.O;
  const rows  = sh.getRange(2, COL.A, N, width).getValues();

  // 先に G:J をクリア
  sh.getRange(2, COL.G, N, 4).clearContent();

  // 出力バッファ（N×1）
  const colG = to2dBlank(N);
  const colH = to2dBlank(N);
  const colI = to2dBlank(N);
  const colJ = to2dBlank(N);

  // 緯度経度が入っている行だけを点集合にする
  const ptsAll = rows.map((r, i) => ({
    idx:  i,
    row:  i + 2,
    name: r[COL.A - 1],
    lat:  r[COL.E - 1],
    lng:  r[COL.F - 1],
    seg:  String(r[COL.O - 1] || "CUR").trim() || "CUR"
  })).filter(p => typeof p.lat === "number" && typeof p.lng === "number");

  if (ptsAll.length === 0) return;

  // 対象セグメントを決定
  const mode = segMode || "pub_new";
  let pts;
  if (mode === "all") {
    pts = ptsAll.slice();
  } else if (mode === "cur") {
    pts = ptsAll.filter(p => p.seg === "CUR");
  } else { // "pub_new"
    pts = ptsAll.filter(p => p.seg === "PUB" || p.seg === "NEW");
  }
  if (pts.length === 0) return;

  // 距離関数（道路っぽい距離）の準備
  const latRef = pts.reduce((s, p) => s + p.lat, 0) / pts.length;
  const detour = calibDetourFactor_(pts);
  const distKm = (A, B) => roadishKm_(A, B, latRef, detour);

  // ===== ルート生成：近い順に「全部」つなぐ =====

  const usedIndex = new Set();   // pts 内で使ったインデックス
  const orderIdx  = [];          // pts 内の順番（0..pts.length-1）

  // スタート地点を決める
  let currentIdx;
  if (startPoint && typeof startPoint.lat === "number" && typeof startPoint.lng === "number") {
    // 起点から最も近い点をスタートに
    let best = Infinity;
    let bestIdx = 0;
    for (let i = 0; i < pts.length; i++) {
      const d = distKm(startPoint, pts[i]);
      if (d < best) {
        best    = d;
        bestIdx = i;
      }
    }
    currentIdx = bestIdx;
  } else {
    // 起点なし：とりあえず最初の点をスタートに
    currentIdx = 0;
  }

  // 近傍法で全件つなぐ
  while (orderIdx.length < pts.length) {
    usedIndex.add(currentIdx);
    orderIdx.push(currentIdx);

    // 次の候補を探す
    let nextIdx = null;
    let best = Infinity;

    for (let i = 0; i < pts.length; i++) {
      if (usedIndex.has(i)) continue;
      const d = distKm(pts[currentIdx], pts[i]);
      if (d < best) {
        best    = d;
        nextIdx = i;
      }
    }

    if (nextIdx === null) {
      // 使っていない点がもう無ければ終了
      break;
    }
    currentIdx = nextIdx;
  }

  // ===== G〜J列を埋める（対象セグメント分） =====

  let seq = 1;
  let prevPoint = null;

  for (const idxInPts of orderIdx) {
    const p = pts[idxInPts];

    let dkm = 0;
    if (prevPoint) {
      dkm = distKm(prevPoint, p);
    }
    const minutes = (dkm / o.speedKmPerHour) * 60;

    colG[p.idx][0] = seq++;
    colH[p.idx][0] = Number(dkm.toFixed(2));
    colI[p.idx][0] = `${Math.round(minutes)} 分`;
    colJ[p.idx][0] = `=HYPERLINK("https://www.google.com/maps/search/?api=1&query=" & E${p.row} & "," & F${p.row}, "地図")`;

    prevPoint = p;
  }

  // ===== 対象外の行（緯度なし or セグメント外）にも G の後番を振る =====
  for (let r = 0; r < N; r++) {
    if (!colG[r][0]) {
      colG[r][0] = seq++;
    }
  }

  // シートに反映（並び替え自体はラッパの sortByG_ で実施）
  sh.getRange(2, COL.G, N, 1).setValues(colG);
  sh.getRange(2, COL.H, N, 1).setValues(colH);
  sh.getRange(2, COL.I, N, 1).setValues(colI);
  sh.getRange(2, COL.J, N, 1).setFormulas(colJ);
}




// 6) 黒●最終判定＋クラスタID（L/M）
function markTargets_(sheetName){
  const sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if(!sh) throw new Error("シートなし: "+sheetName);
  const lastRow=sh.getLastRow(); if(lastRow<=1) return;
  const width=COL.O;
  const data=sh.getRange(2,COL.A,lastRow-1,width).getValues();
  const today=new Date();

  const pts=data.map((r,i)=>({i,lat:r[COL.E-1],lng:r[COL.F-1],neigh:[],gid:null}));
  for(let a=0;a<pts.length;a++){
    if(pts[a].lat===""||pts[a].lng===""||pts[a].lat==null||pts[a].lng==null) continue;
    for(let b=a+1;b<pts.length;b++){
      if(pts[b].lat===""||pts[b].lng===""||pts[b].lat==null||pts[b].lng==null) continue;
      if(haversineKm(pts[a].lat,pts[a].lng,pts[b].lat,pts[b].lng)<=THRESH.clusterKm){
        pts[a].neigh.push(b); pts[b].neigh.push(a);
      }
    }
  }

  const visited=new Set(); let temp=1; const clusterMap={};
  for(let s=0; s<pts.length; s++){
    if(visited.has(s)||pts[s].neigh.length===0) continue;
    const q=[s], set=new Set([s]); visited.add(s);
    while(q.length){
      const cur=q.pop();
      for(const nb of pts[cur].neigh) if(!visited.has(nb)){ visited.add(nb); set.add(nb); q.push(nb); }
    }
    if(set.size>=THRESH.clusterMinSize){ set.forEach(i=>pts[i].gid=temp); clusterMap[temp]=Array.from(set); temp++; }
  }
  const order=Object.entries(clusterMap).sort((a,b)=>b[1].length-a[1].length);
  const idMap={}; order.forEach(([tid],i)=>idMap[tid]=i+1);

  const colL=to2dBlank(lastRow-1), colM=to2dBlank(lastRow-1);
  for(let i=0;i<data.length;i++){
    const keyman=data[i][COL.D-1];
    const visitsTxt=data[i][COL.N-1];
    const schedule=data[i][COL.J-1];
    const hasKeyman=!!String(keyman||"").trim();
    const v=countVisits(visitsTxt);
    let recent=false;
    if(schedule instanceof Date){
      const dd=Math.floor((schedule-today)/(1000*60*60*24));
      recent=(dd>=0 && dd<=THRESH.recentDays);
    }
    colL[i][0]=((hasKeyman||v>=5)&&!recent)?"●":"";
    const finalId=idMap[pts[i].gid];
    colM[i][0]=finalId?`#${String(finalId).padStart(2,"0")}`:"";
  }
  sh.getRange(2,COL.L,lastRow-1,1).setValues(colL);
  sh.getRange(2,COL.M,lastRow-1,1).setValues(colM);
}

// 7) メニュー用ラッパ関数（5個だけ）

// 1) 最初（転記＋緯度経度取得）
function menu_initAll(){
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const dest = ss.getSheetByName(DEST_SHEET_MAIN) || ss.insertSheet(DEST_SHEET_MAIN);
  const lr   = Math.max(dest.getLastRow(), 1);

  if (lr > 1) {
    dest.getRange(2, COL.A, lr - 1, COL.O).clearContent().setBackground(null);
  }

  processLeadsAndJudge_(DEST_SHEET_MAIN, PIPELINES.leads.sources.concat(PIPELINES.published.sources));
  fillLatLng_(DEST_SHEET_MAIN);
}

// 2) 全部並び替え（起点アリ / H1起点）
function menu_sortAll_WithStartFromH1(){
  const start = getStartFromH1_(DEST_SHEET_MAIN);
  if(!start){
    Browser.msgBox("H1が空です。起点住所か「緯度,経度」を入力してください。");
    return;
  }
  optimizeRoute_Auto_Core_(DEST_SHEET_MAIN, AUTO_CONF, start, "all");
  sortByG_(DEST_SHEET_MAIN);
  markTargets_(DEST_SHEET_MAIN);
}

// 3) 全部並び替え（起点ナシ）
function menu_sortAll_NoStart(){
  optimizeRoute_Auto_Core_(DEST_SHEET_MAIN, AUTO_CONF, null, "all");
  sortByG_(DEST_SHEET_MAIN);
  markTargets_(DEST_SHEET_MAIN);
}

// 4) 赤と緑並び替え（起点ナシ）＝ PUB＋NEW
function menu_sortRedGreen_NoStart(){
  optimizeRoute_Auto_Core_(DEST_SHEET_MAIN, AUTO_CONF, null, "pub_new");
  sortByG_(DEST_SHEET_MAIN);
  markTargets_(DEST_SHEET_MAIN);
}

// 5) 青並び替え（起点ナシ）＝ CUR
function menu_sortBlue_NoStart(){
  optimizeRoute_Auto_Core_(DEST_SHEET_MAIN, AUTO_CONF, null, "cur");
  sortByG_(DEST_SHEET_MAIN);
  markTargets_(DEST_SHEET_MAIN);
}

/******************************************************
 * Google Apps Script - 社名自動職種分類システム 最終版
 * 対象: 営業先リスト/新規（A列→O列）
 * 辞書: DICT（無ければ自動生成）
 * 出力: O列（分類不可は「分類不可」と出力）
 ******************************************************/

const TARGET_SHEET = '営業先リスト/新規'; // 対象シート
const DICT_SHEET = 'DICT';                // 辞書シート
const OUT_COL = 15;                       // 出力列（O列=15）

/* ===== メイン関数 ===== */
function classifyTargets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(TARGET_SHEET);
  if (!sh) throw new Error(`「${TARGET_SHEET}」が見当たらん…`);

  ensureDict_(); // DICTがなければ自動生成
  const dict = loadDict_();

  const last = sh.getLastRow();
  if (last < 2) return Logger.log('データ行が無いのう…');

  const names = sh.getRange(2, 1, last - 1, 1).getValues().map(r => r[0]);
  const out = names.map(n => [classifyCompany_(n, dict)]);
  sh.getRange(2, OUT_COL, out.length, 1).setValues(out);

  Logger.log(`${out.length}件の分類完了`);
}

/* ===== DICT自動生成 ===== */
function ensureDict_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(DICT_SHEET);
  if (sh) return;

  sh = ss.insertSheet(DICT_SHEET);
  sh.getRange(1, 1, 1, 4).setValues([['pattern', 'category', 'matchType', 'priority']]);

  const dictData = [
    ['建設|土木|電気工事|配管|内装|塗装|建築|電工|工事|施工|設計|エンジニアリング|舗装|アスコン|外構|左官|解体|足場|とび|型枠', '建設・工事', 'regex', 9],
    ['学校|塾|保育園|幼稚園|専門学校|キッズ|スクール|こども|教室|英会話|予備校|アカデミー|Kids', '教育・スクール', 'regex', 9],
    ['IT|ソフト|ソフトウェア|システム|Web|SaaS|クラウド|アプリ|ネット|デジタル|AI|データ|プログラム|ITサービス|Think\\s*Base|R\\s*3|ZERO\\s*カンパニー', 'IT/Web', 'regex', 8],
    ['温泉|ホテル|旅館|観光|レジャー|スポーツクラブ|ジム|フィットネス|HOTEL|スポーツ|ゴルフ|テニス|ブライダル|ウェディング|Plan\\s*･?\\s*Do\\s*･?\\s*See|アート|SPORTS', 'アミューズメント業', 'regex', 8],
    ['弁護士|税理士|社労士|一般社団法人|一般財団法人|協会|組合|連合会|医療生活協同組合|商工会|医師会|行政法人|機構|グループ|事務所', '士業・団体', 'regex', 10],
    ['運送|運輸|輸送|配送|配車|路線便|トラック|ロジ|物流|チャーター|宅配|運搬|観光バス|ライン|タクシー|ハイヤー|旅客', '運送業', 'regex', 7],
    ['倉庫|保管|庫内|3PL|物流センター|流通センター|配送センター|入出庫|在庫|ピッキング|仕分け|倉庫業', '倉庫業', 'regex', 10],
    ['居酒屋|焼肉|ホルモン|レストラン|カフェ|寿司|ラーメン|酒場|バル|日本酒|唐揚げ|ハイボール|ワイン|肉|魚|料理|酒|炭|焼き鳥|そば|蕎麦|割烹|イタリアン|サカバ', '飲食業', 'regex', 10],
    ['コンビニ|ファミリーマート|セブンイレブン|セブン\\-?イレブン|ローソン|スーパー|家電量販|ドラッグストア|書店|ホームセンター|ASA|デパート|商事|商会|産商|タカラ|アップル|イワタニ|筑波産商|ガレージ|商店|カーショップ|ディーラー', '小売業', 'regex', 10],
    ['ビルメン|設備管理|清掃|警備|リフォーム|造園|害虫駆除|グリーン|ガーデン|クリーン|テックエージェンシー|環境|美化|レンタル|リース|フォークリフト|保守|点検|サービス|メカトロ', '現場系サービス', 'regex', 8],
    ['病院|クリニック|診療所|薬局|介護|特別養護老人ホーム|老健|デイサービス|訪問看護|福祉|メディカル|歯科|生活|福祉施設', '医療・福祉', 'regex', 10],
    ['人材|派遣|職業紹介|BPO|RPO|コールセンター|営業代行|求人|採用|リクルート', '人材・採用', 'regex', 9],
    ['不動産|賃貸|売買|仲介|管理|デベロ|ホーム|区画整理|開発|建物|住宅|ハウス|オフィス|ビジネスセンター|興発', '不動産業', 'regex', 9],
    ['美容室|理容|ヘアサロン|ネイル|エステ|脱毛|リラク|サロン|まつげ|リラクゼーション|整体|ボディケア', '美容・サロン', 'regex', 9],
    ['工業|製作所|製造|メーカー|印刷|金属|樹脂|化学|食品工業|機械|部品|工場|理工|電化|デバイス|ヤクシン|テック|ホシザキ|トランステック|製版|加工|プレス|塗装|鋳造|パッケージ|フロロ|メックス|エバテック', '製造業', 'regex', 9],
  ];

  sh.getRange(2, 1, dictData.length, 4).setValues(dictData);
  sh.autoResizeColumns(1, 4);
  Logger.log('DICTを自動生成したぞい。');
}

/* ===== 社名分類 ===== */
function classifyCompany_(rawName, dict) {
  if (!rawName) return '分類不可';
  const name = normalizeName_(String(rawName));
  const nameC = compress_(name);
  let best = { cat: '', pri: -Infinity, hitLen: -1 };

  for (const d of dict) {
    if (!d.pattern || !d.category) continue;
    if (matchByType_(name, nameC, d.pattern, d.matchType)) {
      const len = String(d.pattern).length;
      if (d.priority > best.pri || (d.priority === best.pri && len > best.hitLen))
        best = { cat: d.category, pri: d.priority, hitLen: len };
    }
  }
  return best.cat || '分類不可';
}

/* ===== マッチ判定 ===== */
function matchByType_(name, nameC, pattern, type) {
  const pn = normalizeName_(pattern);
  const pnC = compress_(pn);
  switch ((type || '').toLowerCase()) {
    case 'exact': return name === pn || nameC === pnC;
    case 'prefix': return name.startsWith(pn) || nameC.startsWith(pnC);
    case 'contains': return name.includes(pn) || nameC.includes(pnC);
    case 'regex':
      try {
        const re = new RegExp(pattern);
        return re.test(name) || re.test(nameC);
      } catch (e) { return false; }
    default: return name.includes(pn) || nameC.includes(pnC);
  }
}

/* ===== 正規化 ===== */
function normalizeName_(s) {
  let t = s.normalize('NFKC').trim();
  const legal = ['株式会社','（株）','(株)','㈱','有限会社','（有）','(有)','合同会社','（同）',
                 'Inc.','Co., Ltd.','Ltd.','LLC','有限責任事業組合'];
  legal.forEach(k => t = t.replace(new RegExp(k, 'gi'), ''));
  t = t.replace(/\b(本社|本部|支店|営業所|出張所|センター|物流センター|流通センター|配送センター|事業本部|工場|倉庫|配車管理|店)\b/g, '');
  t = t.replace(/\b(東京|埼玉|神奈川|千葉|大阪|愛知|名古屋|福岡|仙台|札幌|関東|関西|東海|東日本|西日本|大宮|川口|越谷|志木|朝霞)\b/g, '');
  t = t.replace(/[［\[\]（）()｢｣「」・—―‐\-･・]/g, ' ')
       .replace(/\s+/g, ' ')
       .trim();
  return t;
}

/* ===== 空白削除版 ===== */
function compress_(s) {
  return String(s || '').replace(/\s+/g, '');
}

/* ===== DICT読み込み ===== */
function loadDict_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DICT_SHEET);
  const rows = Math.max(0, sh.getLastRow() - 1);
  if (!rows) return [];
  return sh.getRange(2, 1, rows, 4).getValues().map(r => ({
    pattern: String(r[0] || '').trim(),
    category: String(r[1] || '').trim(),
    matchType: String(r[2] || 'contains').toLowerCase(),
    priority: Number(r[3] || 0)
  }));
}

/***** 架電記録用：ユニーク命名版（他ファイルと衝突しない） *****/
// 設定（必要に応じ変更可）

/**
 * 架電記録/新規：I24に今日から5営業日後、I25に今日から3営業日後を入力
 * 15:00以降は翌日を起点に計算（JST固定）
 * 対象シート: 架電記録/新規
 * 営業日定義: 土日＋日本の祝日を除外
 */

const TARGET_SHEET_NAME = '架電記録/新規';
const TZ = 'Asia/Tokyo';
const CUTOFF_HOUR = 15;
const HOL_CAL_ID = 'ja.japanese#holiday@group.v.calendar.google.com';

/** 実行本体 */
function fillBusinessDays() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(TARGET_SHEET_NAME);
  if (!sh) return;

  const base = getBaseDate_();
  const date5 = addBusinessDays_(base, 5);
  const date3 = addBusinessDays_(base, 3);

  sh.getRange('I24').setValue(date5);
  sh.getRange('I25').setValue(date3);
}

/** 15時締めルール：15時以降は翌日を起点に */
function getBaseDate_() {
  const now = new Date();
  const jst = new Date(Utilities.formatDate(now, TZ, 'yyyy/MM/dd HH:mm:ss'));
  const base = new Date(jst.getFullYear(), jst.getMonth(), jst.getDate());
  if (jst.getHours() >= CUTOFF_HOUR) base.setDate(base.getDate() + 1);
  return base;
}

/** 日本の祝日リスト取得（キャッシュあり） */
function getHolidaySet_(year) {
  const cache = CacheService.getScriptCache();
  const key = `jp_holidays_${year}`;
  const cached = cache.get(key);
  if (cached) return new Set(JSON.parse(cached));

  const cal = CalendarApp.getCalendarById(HOL_CAL_ID);
  const from = new Date(year, 0, 1);
  const to   = new Date(year, 11, 31);
  const events = cal.getEvents(from, to);
  const ymd = d => Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
  const set = new Set(events.map(e => ymd(e.getStartTime())));
  cache.put(key, JSON.stringify([...set]), 21600); // 6時間キャッシュ
  return set;
}

/** 営業日加算（祝日＋土日を除外） */
function addBusinessDays_(startDate, daysToAdd) {
  const dir = daysToAdd >= 0 ? 1 : -1;
  let remaining = Math.abs(daysToAdd);
  let d = new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate());
  const holidaySets = new Map();

  while (remaining > 0) {
    d.setDate(d.getDate() + dir);
    const y = d.getFullYear();
    if (!holidaySets.has(y)) holidaySets.set(y, getHolidaySet_(y));
    const ymd = Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
    const day = d.getDay();
    const isWeekend = (day === 0 || day === 6);
    const isHoliday = holidaySets.get(y).has(ymd);
    if (!isWeekend && !isHoliday) remaining--;
  }
  return d;
}

//=== 新規ホット：営業先リスト/新規 → “当日ちょうど”で毎日追記 ===========================
// ・L列が Ⓐ/Ⓑ/Ⓒ の行のみ対象
// ・J列(最終有効接触日)から 7/14/21/28/60/90 日“ちょうど”の日だけ追記
// ・D列は "indeed" テキストで社名検索リンクを付与（住所は使わない）
// ・同日・同社名・同段階の重複は防止
// ・追記後に「新規ホット」と「会話履歴/新規」の社名一致で双方向リンクを作成
// ・★3カ月の通知を追記した行の元データ（営業先リスト/新規）のL列を「判断がつかない」に変更
function appendShinkiHotDaily(){
  const SS = SpreadsheetApp.getActive();
  const SRC = SS.getSheetByName('営業先リスト/新規');
  if(!SRC) throw new Error('シートが見当たらん: 営業先リスト/新規');
  const DST = SS.getSheetByName('新規ホット') || SS.insertSheet('新規ホット');
  ensureHotHeader_(DST);

  const last = SRC.getLastRow();
  if(last < 2) return;

  const rows = SRC.getRange(2,1,last-1,12).getValues();  // A〜L
  const tz = Session.getScriptTimeZone();
  const today = new Date();
  const todayStr = Utilities.formatDate(today, tz, 'yyyy/MM/dd');

  // 同日・同社名・同段階の重複防止
  const already = buildTodayIndex_(DST, todayStr);

  const outVals = [];  // A:社名 B:段階 C:通知日 D:(後でリンク) E:電話 F:キーマン
  const outRich = [];  // D列のRichTextValue
  const addrList = []; // ★K列へ書く住所の一時保持

  // ★3カ月該当でL列を書き換える対象の行番号（元シートの行番号）を溜める
  const undecidedRows = [];

  for(let i=0; i<rows.length; i++){
    const r = rows[i];
    const srcRow = i + 2;                 // 元シートの行番号
    const name = String(r[0]||'').trim(); // A 社名
    const tel  = String(r[1]||'').trim(); // B 電話
    const addr = String(r[2]||'').trim(); // ★C 住所
    const key  = String(r[3]||'').trim(); // D キーマン
    const jVal = r[9];                    // J 最終有効接触日
    const stat = String(r[11]||'').trim();// L ステータス
    if(!name) continue;
    if(!isHotStatus_(stat)) continue;

    const base = toDate_(jVal); if(!base) continue;
    const days = diffDays_(base, today);
    const bucket = stageFromDays_(days);  // 当日ちょうどのみ返す
    if(!bucket) continue;

    const dupKey = name + '|' + bucket;
    if(already.has(dupKey)) continue;

    // D列：indeedリンク（社名のみ）
    const url = `https://jp.indeed.com/jobs?q=${encodeURIComponent(name)}`;
    const link = SpreadsheetApp.newRichTextValue().setText('indeed').setLinkUrl(url).build();

    outVals.push([name, bucket, todayStr, '', tel, key]);
    outRich.push(link);
    addrList.push(addr); // ★住所を同じ順序で保持

    // ★3カ月通知を追記するタイミングで、元シートL列を「判断がつかない」に変更
    if(bucket === '3カ月'){
      undecidedRows.push(srcRow);
    }
  }

  appendHotRows_(DST, outVals, outRich, addrList);

  // ★ 元シートのL列一括更新（非連続セルをRangeListでまとめて更新）
  if(undecidedRows.length){
    markStatusUndecidedRows_(SRC, undecidedRows);
  }

  // ★ 追記後：社名一致でホット⇄会話履歴の相互リンクを作成
  linkHotAndHistory();
  sortHotByDateDesc_();

}


//=== 初回用：過去ぶん一括取り込み（通知日は過去日で記録） ============================
// ・L列が Ⓐ/Ⓑ/Ⓒ の行のみ対象
// ・J列 + [7,14,21,28,60,90]日の各マイルストーンが “今日より過去” のものを全部追記
// ・C列（通知日）はその過去日、D列は "indeed" リンク
// ・既存重複は「社名|段階|通知日」で防止
// ・取り込み後にホット⇄会話履歴の相互リンクを作成
// ・★3カ月（90日）を追記した場合、元シートL列を「判断がつかない」に変更
function backfillShinkiHotAllPast(){
  const SS  = SpreadsheetApp.getActive();
  const SRC = SS.getSheetByName('営業先リスト/新規');
  if(!SRC) throw new Error('シートが見当たらん: 営業先リスト/新規');
  const DST = SS.getSheetByName('新規ホット') || SS.insertSheet('新規ホット');
  ensureHotHeader_(DST);

  const last = SRC.getLastRow();
  if(last < 2) return;

  const rows  = SRC.getRange(2,1,last-1,12).getValues(); // A〜L
  const tz    = Session.getScriptTimeZone();
  const today = stripTime_(new Date());
  const exIdx = buildExistingIndexByDate_(DST); // Set<name|bucket|yyyy/MM/dd>
  const OFFSETS = [7,14,21,28,60,90];

  const outVals = [];
  const outRich = [];
  const addrList = []; // ★K列へ書く住所の一時保持
  const undecidedRows = []; // ★3カ月該当でL列書き換え対象

  for(let i=0; i<rows.length; i++){
    const r = rows[i];
    const srcRow = i + 2;
    const name = String(r[0]||'').trim();
    const tel  = String(r[1]||'').trim();
    const addr = String(r[2]||'').trim(); // ★C 住所
    const key  = String(r[3]||'').trim();
    const jVal = r[9];
    const stat = String(r[11]||'').trim();
    if(!name) continue;
    if(!isHotStatus_(stat)) continue;

    const base = toDate_(jVal); if(!base) continue;
    const base0 = stripTime_(base);

    for(const d of OFFSETS){
      const notifyDate = addDays_(base0, d);
      if(notifyDate > today) continue; // 未来は除外
      const bucket = stageFromDays_(d);
      const dateStr = Utilities.formatDate(notifyDate, tz, 'yyyy/MM/dd');
      const dedupKey = `${name}|${bucket}|${dateStr}`;
      if(exIdx.has(dedupKey)) continue;

      const url = `https://jp.indeed.com/jobs?q=${encodeURIComponent(name)}`;
      const link = SpreadsheetApp.newRichTextValue().setText('indeed').setLinkUrl(url).build();

      outVals.push([name, bucket, dateStr, '', tel, key]);
      outRich.push(link);
      addrList.push(addr); // ★住所を同じ順序で保持

      // ★3カ月のレコードを新規追加する場合のみ、元シートL列を更新対象に
      if(d === 90){
        undecidedRows.push(srcRow);
      }
    }
  }

  appendHotRows_(DST, outVals, outRich, addrList);

  // ★ 元シートL列を一括で「判断がつかない」に更新
  if(undecidedRows.length){
    markStatusUndecidedRows_(SRC, undecidedRows);
  }

  // ★ 取り込み後：社名一致でホット⇄会話履歴の相互リンクを作成
  linkHotAndHistory();

  SpreadsheetApp.getUi().alert(`バックフィル完了：${outVals.length} 行を追加しました`);
}

function linkHotAndHistory(){
  const SS = SpreadsheetApp.getActive();
  const SH_HOT = SS.getSheetByName('新規ホット');
  const SH_HIS = SS.getSheetByName('会話履歴/新規');
  if(!SH_HOT || !SH_HIS) return;

  const hotLast = SH_HOT.getLastRow();
  const hisLast = SH_HIS.getLastRow();
  if(hotLast < 2 || hisLast < 2) return;

  const hotVals = SH_HOT.getRange(2,1,hotLast-1,1).getDisplayValues().map(r => String(r[0]).trim());
  const hisVals = SH_HIS.getRange(2,1,hisLast-1,1).getDisplayValues().map(r => String(r[0]).trim());

  const hisMap = new Map();
  for(let i=0;i<hisVals.length;i++){
    const nm = hisVals[i];
    if(nm) hisMap.set(nm, i+2);
  }

  let cnt = 0;
  for(let i=0;i<hotVals.length;i++){
    const name = hotVals[i];
    if(!name) continue;

    const hisRow = hisMap.get(name);
    if(!hisRow) continue;

    const hotRow = i + 2;
    const hisUrl = `#gid=${SH_HIS.getSheetId()}&range=A${hisRow}`;

    // ★ホット → 履歴 のみ作成
    const hotLink = SpreadsheetApp.newRichTextValue().setText(name).setLinkUrl(hisUrl).build();
    SH_HOT.getRange(hotRow,1).setRichTextValue(hotLink);

    cnt++;
  }

  Logger.log(`linkHotAndHistory: ${cnt} 件リンク更新`);
}


//=== 新規ホットへ追記：A列の最終データ行の“次の行”から A〜F・D列リンク・K列住所をまとめて書く ====
function appendHotRows_(dst, outVals, outRich, addrList){
  if(!outVals.length) return;
  const start = getNextAppendRowByColA_(dst);
  const n = outVals.length;
  dst.getRange(start,1,n,6).setValues(outVals);
  dst.getRange(start,4,n,1).setRichTextValues(outRich.map(r => [r]));        // D列にリンク
  dst.getRange(start,11,n,1).setValues(addrList.map(a => [a || '']));        // K列に住所
}

//=== ステータス判定（Ⓐ/Ⓑ/Ⓒ のみ） ====================================================
function isHotStatus_(s){
  return /Ⓐ1か月以内ニーズ有 × 選んでくれる|Ⓑ1か月以内ニーズ有 × 選ぶか不明|Ⓒ1か月以内ニーズ有 × 選ばなそう/.test(s||'');
}

//=== 経過日 → 段階（当日ちょうどのみ） ================================================
function stageFromDays_(d){
  if(d===7)  return '1週間';
  if(d===14) return '2週間';
  if(d===21) return '3週間';
  if(d===28) return '4週間';
  if(d===60) return '2カ月';
  if(d===90) return '3カ月';
  return null;
}
//=== オフセット → 段階名（バックフィル用） =============================================
//=== “今日”すでにある (社名|段階) のセット（重複防止：毎日追記用） =====================
function buildTodayIndex_(dst, todayStr){
  const set = new Set();
  const last = dst.getLastRow();
  if(last < 2) return set;
  const vals = dst.getRange(2,1,last-1,3).getValues(); // A:社名 B:段階 C:通知日
  for(const [n,b,d] of vals){
    if(String(d||'').trim() === todayStr){
      set.add(String(n||'').trim() + '|' + String(b||'').trim());
    }
  }
  return set;
}

//=== 既存の (社名|段階|通知日) セット（重複防止：バックフィル用） ========================
function buildExistingIndexByDate_(dst){
  const set = new Set();
  const last = dst.getLastRow();
  if(last < 2) return set;
  const vals = dst.getRange(2,1,last-1,3).getValues(); // A:社名 B:段階 C:通知日
  for(const [n,b,d] of vals){
    const name = String(n||'').trim();
    const bucket = String(b||'').trim();
    const dateStr = String(d||'').trim();
    if(name && bucket && dateStr) set.add(`${name}|${bucket}|${dateStr}`);
  }
  return set;
}

//=== ヘッダー整備 ======================================================================
function ensureHotHeader_(sh){
  if(sh.getLastRow()===0){
    // 既存のA〜Fヘッダ
    sh.getRange(1,1,1,6).setValues([[
      '社名','ステータス','通知日','indeed','電話番号','キーマン'
    ]]);
    sh.setFrozenRows(1);
  }else{
    const h = sh.getRange(1,1,1,6).getValues()[0];
    if(!h[0]) sh.getRange(1,1).setValue('社名');
    if(!h[1]) sh.getRange(1,2).setValue('ステータス');
    if(!h[2]) sh.getRange(1,3).setValue('通知日');
    if(!h[3]) sh.getRange(1,4).setValue('indeed');
    if(!h[4]) sh.getRange(1,5).setValue('電話番号');
    if(!h[5]) sh.getRange(1,6).setValue('キーマン');
    sh.setFrozenRows(1);
  }
  // ★K列(11列目)のヘッダーに「住所」を補完（空なら）
  const kHeader = String(sh.getRange(1,11).getDisplayValue()||'').trim();
  if(!kHeader) sh.getRange(1,11).setValue('住所');
}

//=== A列の最終データ行の“次の行”を返す（ヘッダ1行を想定） ==============================
function getNextAppendRowByColA_(sh){
  const last = sh.getLastRow();
  if(last < 2) return 2; // ヘッダのみなら2行目から
  const colA = sh.getRange(2,1,last-1,1).getDisplayValues();
  for(let i = colA.length - 1; i >= 0; i--){
    if(String(colA[i][0]||'').trim()){
      return 2 + i + 1; // 見つかった最終データ行の“次の行”
    }
  }
  return 2;
}

//=== 日付系ユーティリティ ===============================================================
function toDate_(v){
  if(v instanceof Date) return v;
  const s = String(v||'').trim(); if(!s) return null;
  const m = s.match(/(\d{4})[\/\-\.](\d{1,2})[\/\-\.](\d{1,2})/);
  if(m){ const d=new Date(+m[1],+m[2]-1,+m[3]); return isNaN(d)?null:d; }
  if(!isNaN(s)){ const n=Number(s); const base=new Date(1899,11,30); return new Date(base.getTime()+n*86400000); }
  return null;
}
function stripTime_(dt){ return new Date(dt.getFullYear(), dt.getMonth(), dt.getDate()); }
function addDays_(dt, days){ const d=new Date(dt); d.setDate(d.getDate()+days); return stripTime_(d); }
function diffDays_(from,to){
  const f=new Date(from.getFullYear(),from.getMonth(),from.getDate());
  const t=new Date(to.getFullYear(),to.getMonth(),to.getDate());
  return Math.floor((t-f)/86400000);
}

//=== （任意）毎日9時のトリガーを作成 ====================================================
function installDailyTriggerForHotAppend(){
  const fn='appendShinkiHotDaily';
  ScriptApp.getProjectTriggers().forEach(t=>{ if(t.getHandlerFunction()===fn) ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger(fn).timeBased().atHour(9).everyDays(1).inTimezone(Session.getScriptTimeZone()).create();
}

//=== ★補助：元シートの任意行（L列）を一括で「判断がつかない」に更新 ======================
function markStatusUndecidedRows_(sheet, rowNumbers){
  // RangeListを使って非連続セルを一括更新
  const a1s = rowNumbers.map(r => `L${r}`);
  sheet.getRangeList(a1s).setValue('判断がつかない');
}

//=== ★「新規ホット」をC列(通知日)の降順で並び替える ================================
function sortHotByDateDesc_(){
  const SS = SpreadsheetApp.getActive();
  const SH = SS.getSheetByName('新規ホット');
  if(!SH) return;
  const last = SH.getLastRow();
  if(last < 3) return; // データが1件以下ならスキップ
  // ヘッダーを除いた範囲をC列(3列目)で降順ソート
  SH.getRange(2, 1, last - 1, SH.getLastColumn()).sort({ column: 3, ascending: false });
}

// シート「新規ホット」の色を毎朝リセットし、チェックあり行を灰色にする
// さらに G列（7列目）を2行目以降でクリア
function resetAndHighlightRows() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("新規ホット");
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return; // データなしなら終了

  // G列をクリア（見出し行は残す・書式は残す）
  sheet.getRange(2, 7, lastRow - 1, 1).clearContent();

  // 背景をリセットし、I列（9列目）にチェックがある行を灰色に
  const lastCol = sheet.getLastColumn();
  const range = sheet.getRange(2, 1, lastRow - 1, lastCol);
  const checks = sheet.getRange(2, 9, lastRow - 1, 1).getValues();
  range.setBackgrounds(checks.map(c => Array(lastCol).fill(c[0] === true ? "#d3d3d3" : null)));
}

// 初回に実行してトリガーを設定する関数（毎朝9:30実行）
function createDailyTrigger() {
  // 既存のトリガーをクリア（重複防止）
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === "resetAndHighlightRows") {
      ScriptApp.deleteTrigger(t);
    }
  });

  // ★9:30に実行（日本時間はプロジェクトのタイムゾーン設定に従う）
  ScriptApp.newTrigger("resetAndHighlightRows")
           .timeBased()
           .everyDays(1)
           .atHour(9)
           .nearMinute(30)
           .create();
}

// ===== メニュー（既存＋時間報メニュー） =====
function onOpen() {
  SpreadsheetApp.getUi().createMenu('メニュー')
    .addItem('新規リフレッシュ', 'runAllSafe_Final')
    .addItem('リンク再生', 'runAllProcessesCombined')
    .addItem('今日の掛け先をハイライト', 'updateAllSheetsAndResetColors_Safe')
    .addItem('データ補正', 'processPhoneNumbers')
    .addItem('顧客データ追加', 'transfer営業先リスト完全版')
    .addItem('NG顧客アラート', 'highlightRows')
    .addSeparator()
    .addItem('① 最初（転記＋緯度経度取得）', 'menu_initAll')
    .addItem('② 全部並び替え（起点=H1）', 'menu_sortAll_WithStartFromH1')
    .addItem('③ 全部並び替え（起点なし）', 'menu_sortAll_NoStart')
    .addItem('④ 赤と緑並び替え（起点なし）', 'menu_sortRedGreen_NoStart')
    .addItem('⑤ 青並び替え（起点なし）', 'menu_sortBlue_NoStart')
    .addSeparator()
    .addItem('掲載開始顧客を整形', 'cleanKeisaiKaishiKokyaku')
    .addItem('A〜D列を削除して左詰め', 'deleteABCD_andLeftShift')
    .addItem('繋がらなかった顧客を転記', 'runLavenderTransfer')
    .addToUi();
}


/***********************
 * 安全一括実行（改訂版）
 * 並び替え → 色付け → リンク再生成
 ***********************/
function runAllSafe_Final() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // 時間主導のきっかけでは画面がないため getUi() は使えない。使える時だけ取得する
  var ui = null;
  try { ui = SpreadsheetApp.getUi(); } catch (uiErr) { ui = null; }

  try {
    // ① 先に集計・逆転記・必要な更新だけ実行
    //    ※色付け・リンク再生成はここではまだやらない
    transferStatusToHistory_(ss, "会話履歴/新規", "営業先リスト/新規");
    transferStatusToHistory_(ss, "会話履歴/現S",  "営業先リスト/現S");
    clearDataSheetColumns();
    countStatusSummary();

    // ② 並び替え
    // 新規はスコア計算＋並び替え
    calcSalesPriorityScore();

    // 掲載開始顧客は色優先で並び替え
    colorKeisaiRowsUnified_({
      sheetName: "掲載開始顧客",
      headerRows: 1,
      applySort: true
    });

    // ③ 並び替え後に色付け
    // 今日色
    resetAndPaintTodayRows_(ss, "営業先リスト/新規");
    resetAndPaintTodayRows_(ss, "営業先リスト/現S");

    // 新規：今日電話した＝薄紫／今日かけるべき＝オレンジ（点数と同じ判断）／それ以外＝白
    paintShinkiByPriority_(ss);
    // 現S：期限超過オレンジ
    highlightCellsByStatus_Param("営業先リスト/現S");

    // 現Sの薄黄色
    highlightRowsJNotThisMonth();

    // 掲載開始顧客は並び替え後にもう一度色を確定
    colorKeisaiRowsUnified_({
      sheetName: "掲載開始顧客",
      headerRows: 1,
      applySort: false
    });

    // 営業日補完（fillBusinessDays）
    runFillBusinessDaysSafely_();

    // ④ 一番最後にリンク再生成
    regenerateLinks(ss);

  } catch (e) {
    Logger.log("runAllSafe_Final エラー: " + (e && e.stack ? e.stack : e));
    if (ui) ui.alert("エラー：" + e.message);
    throw e;
  }
}

// 手動実行用：掲載開始顧客の薄紫(#E6E6FA)行を転記
// ・営業先リスト/新規 Upsert（※新規作成時の行コピーは D・E を除外、L列=「キーマンと繋がってない」、N列ステータス整形）
// ・会話履歴/新規 C列へ追記（B列は変更しない、新規も空欄） ※既に同一最終行があれば重複追記しない
// ・相互リンク付与、原票をグレー化
function runLavenderTransfer() {
  const ss     = SpreadsheetApp.getActive();
  const keisai = ss.getSheetByName('掲載開始顧客');
  const sales  = ss.getSheetByName('営業先リスト/新規');
  const hist   = ss.getSheetByName('会話履歴/新規');
  if (!keisai || !sales || !hist) { SpreadsheetApp.getUi().alert('必要シートが見つかりません'); return; }

  const LABEL = 'キーマンと繋がってない';                  // 営業先/新規 L列(12)へ入れる
  const LAVENDER = '#E6E6FA';
  const tz = Session.getScriptTimeZone();
  const now = new Date();
  const ymd = Utilities.formatDate(now, tz, 'yyyy/MM/dd');
  const hh  = Utilities.formatDate(now, tz, 'HH時');

  const lastRow = keisai.getLastRow();
  const lastCol = keisai.getLastColumn();
  const copyCols = Math.min(lastCol, sales.getLastColumn());
  let count = 0;

  if (lastRow >= 2 && lastCol >= 1) {
    // 掲載開始顧客は1回でまとめて読む
    const rng = keisai.getRange(2, 1, lastRow - 1, lastCol);
    const bgs = rng.getBackgrounds();
    const disp = rng.getDisplayValues();
    const vals = rng.getValues();

    for (let i = 0; i < bgs.length; i++) {
      // 対象：行全体が薄紫（処理済みは灰色になっているので対象外）
      if (!bgs[i].every(c => String(c || '').toUpperCase() === LAVENDER)) continue;

      const company = String(disp[i][0] || '').trim();
      if (!company) continue;

      // 1) 会話本文：L列の最終行 > I列 > 定型（先頭の「yyyy/MM/dd HH時 - 」の重なりは削る）
      const Ltext = String(disp[i][11] || '').trim();
      const Itext = String(disp[i][8] || '').trim();
      const plain = String((Ltext ? Ltext.split('\n').pop() : '') || Itext)
        .replace(/^(?:\d{4}\/\d{2}\/\d{2}\s+\d{2}時\s*-\s*)+/, '').trim() || '掲載開始顧客（薄紫）から自動転記';
      const body = `${ymd} ${hh} - ${plain}`;

      // 2) 営業先/新規 Upsert（新規作成時のみ行コピー。D・E 列は転記しない）
      let salesRow = findRowInColA_(sales, company);
      if (salesRow === 0) {
        salesRow = sales.getLastRow() + 1;
        const src = vals[i].slice(0, copyCols);
        sales.getRange(salesRow, 1, 1, Math.min(3, copyCols)).setValues([src.slice(0, 3)]);
        if (copyCols >= 6) sales.getRange(salesRow, 6, 1, copyCols - 5).setValues([src.slice(5)]);
      }
      if (sales.getLastColumn() >= 12) sales.getRange(salesRow, 12).setValue(LABEL);
      if (sales.getLastColumn() >= 14) {
        const nCell = sales.getRange(salesRow, 14);
        nCell.setValue(getOrInitStatus_(nCell.getValue()));
      }

      // 3) 会話履歴/新規 Upsert（B列は触らない・直前と同じ行なら追記しない）
      let histRow = findRowInColA_(hist, company);
      if (histRow > 0) {
        const c = hist.getRange(histRow, 3);
        const cur = String(c.getValue() || '');
        if ((cur ? cur.split('\n').pop() : '') !== body) c.setValue((cur ? cur + '\n' : '') + body);
      } else {
        hist.appendRow([company, '', body]);
        histRow = hist.getLastRow();
      }

      // 4) 相互リンク　5) 原票をグレー化（再実行防止）
      createHyperlinks(sales, hist, salesRow, histRow, company);
      keisai.getRange(i + 2, 1, 1, lastCol).setBackground('#D3D3D3');
      count++;
    }
  }

  SpreadsheetApp.getUi().alert(`薄紫行の転記が完了しました：${count}件（履歴の二重ヘッダ防止、DE列は転記せず）`);
}
