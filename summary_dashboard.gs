// 売上データを月ごとに集計し、「月次サマリー」シートと棒グラフを更新するスクリプト
// スプレッドシートにバインドされたスクリプトとして使う想定

// シート名の設定
const SOURCE_SHEET_NAME = "売上データ";
const SUMMARY_SHEET_NAME = "月次サマリー";

// トリガーで自動実行する時刻（時）
const TRIGGER_HOUR = 9;

// メイン処理：売上データを月ごとに集計し、サマリーとグラフを書き直す
function runDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(SOURCE_SHEET_NAME);
  if (!sourceSheet) {
    throw new Error(`「${SOURCE_SHEET_NAME}」シートが見つかりません。`);
  }

  const timeZone = ss.getSpreadsheetTimeZone();
  const monthlyTotals = aggregateByMonth(sourceSheet, timeZone);

  const summarySheet = ss.getSheetByName(SUMMARY_SHEET_NAME) || ss.insertSheet(SUMMARY_SHEET_NAME);
  writeSummary(summarySheet, monthlyTotals);
  drawChart(summarySheet, monthlyTotals.length);
}

// 毎朝9時に runDashboard を自動実行するトリガーを登録する（最初に1回だけ手動で実行する）
function setDailyTrigger() {
  // 同じトリガーが重複しないよう、既存の runDashboard のトリガーを削除してから登録する
  ScriptApp.getProjectTriggers()
    .filter((trigger) => trigger.getHandlerFunction() === "runDashboard")
    .forEach((trigger) => ScriptApp.deleteTrigger(trigger));

  // GAS の仕様上、実行時刻は9時〜10時の間で多少前後する
  ScriptApp.newTrigger("runDashboard")
    .timeBased()
    .everyDays(1)
    .atHour(TRIGGER_HOUR)
    .nearMinute(0)
    .create();
}

// 「売上データ」シートの全行を読み込み、月ごとの合計売上と件数を返す（古い月から順に並べる）
function aggregateByMonth(sheet, timeZone) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }

  // 1行目は見出しとして読み飛ばし、A〜D列をまとめて読み込む
  const rows = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  const totals = {};

  rows.forEach((row) => {
    const date = toDate(row[0]);
    const amount = Number(row[3]);
    // 日付や金額が不正な行（空行など）は集計しない
    if (!date || row[3] === "" || isNaN(amount)) {
      return;
    }

    // 並べ替え用のキー（例: 2026-01）
    const key = Utilities.formatDate(date, timeZone, "yyyy-MM");
    if (!totals[key]) {
      totals[key] = {
        label: Utilities.formatDate(date, timeZone, "yyyy年M月"),
        total: 0,
        count: 0
      };
    }
    totals[key].total += amount;
    totals[key].count += 1;
  });

  return Object.keys(totals)
    .sort()
    .map((key) => totals[key]);
}

// セルの値を日付に変換する（日付型でも「2026/01/05」のような文字列でも受け付ける）
function toDate(value) {
  if (value instanceof Date) {
    return isNaN(value.getTime()) ? null : value;
  }
  if (typeof value === "string" && value.trim() !== "") {
    const parsed = new Date(value.trim().replace(/-/g, "/"));
    return isNaN(parsed.getTime()) ? null : parsed;
  }
  return null;
}

// 「月次サマリー」シートをクリアして、見出しと集計結果を書き込む
function writeSummary(sheet, monthlyTotals) {
  sheet.clear();

  const values = [["月", "合計売上", "件数"]].concat(
    monthlyTotals.map((item) => [item.label, item.total, item.count])
  );
  const range = sheet.getRange(1, 1, values.length, 3);

  // 「2026年1月」が日付に自動変換されないよう、A列を文字列の書式にしてから書き込む
  sheet.getRange(1, 1, values.length, 1).setNumberFormat("@");
  range.setValues(values);

  // 見ための調整
  sheet.getRange(1, 1, 1, 3).setFontWeight("bold");
  if (monthlyTotals.length > 0) {
    sheet.getRange(2, 2, monthlyTotals.length, 1).setNumberFormat("#,##0");
  }
  sheet.autoResizeColumns(1, 3);
}

// 「月次サマリー」シートに月次推移の棒グラフを作り直す
function drawChart(sheet, monthCount) {
  // clear() ではグラフが消えないため、既存のグラフを削除してから作り直す
  sheet.getCharts().forEach((chart) => sheet.removeChart(chart));

  if (monthCount === 0) {
    return;
  }

  // A列（月）と B列（合計売上）を見出し行込みで使う
  const chart = sheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(sheet.getRange(1, 1, monthCount + 1, 2))
    .setNumHeaders(1)
    .setPosition(1, 5, 0, 0)
    .setOption("title", "月次売上推移")
    .setOption("legend", { position: "none" })
    .setOption("hAxis", { title: "月" })
    .setOption("vAxis", { title: "合計売上", format: "#,##0" })
    .build();

  sheet.insertChart(chart);
}
