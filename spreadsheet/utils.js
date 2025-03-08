// Description: スプレッドシートのユーティリティ関数を定義する
// addIfErrorToDivisionFormulas: 割り算の数式に IFERROR を追加する
// 引数: なし
// 戻り値: なし
// 例外: なし
// 備考: 実行したいスプレッドシートをアクティブにしてからスクリプトを実行する

function addIfErrorToDivisionFormulas() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var range = sheet.getDataRange();
  var numRows = range.getNumRows();
  var numCols = range.getNumColumns();

  // 各セルを1行1列単位で処理
  for (var i = 1; i <= numRows; i++) {
    for (var j = 1; j <= numCols; j++) {
      var cell = sheet.getRange(i, j);
      var formula = cell.getFormula();
      // 数式が存在し、割り算が含まれていて、まだ IFERROR でラップされていない場合
      if (formula && formula.indexOf("/") !== -1 && formula.indexOf("IFERROR(") === -1) {
        // 最初の "=" を除いた式を IFERROR でラップ
        var newFormula = "=IFERROR(" + formula.substring(1) + ", 0)";
        cell.setFormula(newFormula);
      }
    }
  }
}