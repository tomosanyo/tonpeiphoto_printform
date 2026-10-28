//紐づけるフォームIDの取得
//スプレッドシートのIDはすでに紐づいているから設定不要
//新たにフォームを作り直すときは取得すること．
//docs.google.com/forms/d/.../editの...部分を指定すること．
const formID = '1AhKI-q39UNN1tEVmuo4__9Ej_4WdWSmFfw12bUbBUAw';

//スプレッドシートの取得
//シート名の取得，変えたら変更してね
const sheetName = 'フォームの回答 1';
const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);

//formID・sheetName・sheetはグローバル変数なので気を付けてね．
