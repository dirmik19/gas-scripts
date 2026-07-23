/**
 * 売上データ集計ダッシュボード
 *
 * 「売上データ」シートの明細を月ごとに集計し、「月次サマリー」シートへ書き出したうえで
 * 月次推移を示す棒グラフを同じスプレッドシート内に作成する。
 */

// シート名は変更されやすいため定数化しておく
var SOURCE_SHEET_NAME = '売上データ';
var SUMMARY_SHEET_NAME = '月次サマリー';

/**
 * メインの集計処理。
 * 「売上データ」シートを読み込み、月次サマリーを作成してグラフを更新する。
 */
function runDashboard() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = spreadsheet.getSheetByName(SOURCE_SHEET_NAME);

  if (!sourceSheet) {
    throw new Error('「' + SOURCE_SHEET_NAME + '」シートが見つかりません。');
  }

  var monthlyData = aggregateMonthlySales(sourceSheet, spreadsheet.getSpreadsheetTimeZone());
  var summarySheet = getOrCreateSummarySheet(spreadsheet);

  writeSummary(summarySheet, monthlyData);
  drawMonthlyChart(summarySheet, monthlyData.length);
}

/**
 * 「売上データ」シートを読み込み、月ごとの合計売上・件数を集計する。
 * @param {Sheet} sourceSheet 売上データシート
 * @param {string} timeZone 集計に使うタイムゾーン（スプレッドシートの設定に合わせる）
 * @return {Array<Object>} 月順にソートされた集計結果の配列
 */
function aggregateMonthlySales(sourceSheet, timeZone) {
  var lastRow = sourceSheet.getLastRow();

  // ヘッダー行のみ、またはデータなしの場合は空配列を返す
  if (lastRow < 2) {
    return [];
  }

  // A〜D列（日付・担当者名・商品名・金額）を一括取得
  var values = sourceSheet.getRange(2, 1, lastRow - 1, 4).getValues();

  // 月ごとの集計結果を保持するマップ（キー：yyyy-MM、ソート用）
  var monthlyMap = {};

  values.forEach(function (row) {
    var date = row[0];
    var amount = row[3];

    // 日付が空、または不正な行はスキップする
    if (!(date instanceof Date) || isNaN(date.getTime())) {
      return;
    }

    var sortKey = Utilities.formatDate(date, timeZone, 'yyyy-MM');
    var label = Utilities.formatDate(date, timeZone, 'yyyy年M月');

    if (!monthlyMap[sortKey]) {
      monthlyMap[sortKey] = {
        sortKey: sortKey,
        label: label,
        total: 0,
        count: 0
      };
    }

    monthlyMap[sortKey].total += Number(amount) || 0;
    monthlyMap[sortKey].count += 1;
  });

  // 月の昇順（時系列順）に並び替えて返す
  return Object.keys(monthlyMap)
    .sort()
    .map(function (key) {
      return monthlyMap[key];
    });
}

/**
 * 「月次サマリー」シートを取得する。存在しない場合は新規作成する。
 * @param {Spreadsheet} spreadsheet 対象のスプレッドシート
 * @return {Sheet} 月次サマリーシート
 */
function getOrCreateSummarySheet(spreadsheet) {
  var summarySheet = spreadsheet.getSheetByName(SUMMARY_SHEET_NAME);

  if (!summarySheet) {
    summarySheet = spreadsheet.insertSheet(SUMMARY_SHEET_NAME);
  }

  return summarySheet;
}

/**
 * 「月次サマリー」シートを毎回クリアしてから、集計結果を書き込む。
 * @param {Sheet} summarySheet 月次サマリーシート
 * @param {Array<Object>} monthlyData 月ごとの集計結果
 */
function writeSummary(summarySheet, monthlyData) {
  // 既存の内容とグラフを一旦すべてクリアする
  summarySheet.clear();
  summarySheet.getCharts().forEach(function (chart) {
    summarySheet.removeChart(chart);
  });

  // ヘッダー行を書き込む
  summarySheet.getRange(1, 1, 1, 3).setValues([['月', '合計売上', '件数']]);

  if (monthlyData.length === 0) {
    return;
  }

  // 集計結果を書き込む（A：月、B：合計売上、C：件数）
  var rows = monthlyData.map(function (item) {
    return [item.label, item.total, item.count];
  });

  summarySheet.getRange(2, 1, rows.length, 3).setValues(rows);
}

/**
 * 「月次サマリー」シートのデータをもとに、月次推移の棒グラフを作成する。
 * @param {Sheet} summarySheet 月次サマリーシート
 * @param {number} dataRowCount 集計されたデータ行数（ヘッダー行を除く）
 */
function drawMonthlyChart(summarySheet, dataRowCount) {
  if (dataRowCount === 0) {
    return;
  }

  // 月（A列）と合計売上（B列）の範囲をグラフのデータ範囲とする
  var dataRange = summarySheet.getRange(1, 1, dataRowCount + 1, 2);

  var chart = summarySheet
    .newChart()
    .asColumnChart()
    .addRange(dataRange)
    .setPosition(2, 5, 0, 0)
    .setOption('title', '月次売上推移')
    .setOption('legend', { position: 'none' })
    .setOption('hAxis', { title: '月' })
    .setOption('vAxis', { title: '合計売上' })
    .build();

  summarySheet.insertChart(chart);
}

/**
 * 毎朝9時に runDashboard を自動実行するトリガーを設定する。
 * 実行するたびに古いトリガーを削除してから登録し直すため、重複登録は発生しない。
 */
function setDailyTrigger() {
  // runDashboard に紐づく既存のトリガーを削除する
  ScriptApp.getProjectTriggers().forEach(function (trigger) {
    if (trigger.getHandlerFunction() === 'runDashboard') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // 毎朝9時に runDashboard を実行するトリガーを新規作成する
  ScriptApp.newTrigger('runDashboard')
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .create();
}
