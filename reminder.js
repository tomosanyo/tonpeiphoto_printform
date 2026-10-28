function reminderSearch() {
  //実行時刻を取得
  const now = new Date();
  //スプレッドシートの内容を取得
  const data = sheet.getDataRange().getValues();
  //デバッグ用
  let log_counter = 0;
  let log_user = [];
  //1行ずつ検証
  for (let i = 1; i < data
    .length; i++){
    let submitDate = data[i][0];
    let paymentStatus = data[i][20];

    Logger.log(`submitDate: ${submitDate}, user: ${data[i][1]}, paymentStatus: ${paymentStatus}`);
    //支払いステータスが未完了の場合
    if (paymentStatus === `未完了`){
        //提出日から現在までの日数を取得
        let diffDays = Math.floor((now - new Date(submitDate)) / (1000 * 60 * 60 * 24));
        //7日以上経過している場合
        if (diffDays >= 7) {
            sendReminderEmail(data, i)
            log_user.push(data[i][1]);
            log_counter++;
        }
    }
  }
  //デバッグ用
  Logger.log(`リマインダーメールを${log_counter}件送信しました。`);
  //リマインダを送信した場合、レポートを管理者に送信
  if(log_counter > 0){
    sendReport(log_counter, log_user);
  }
}

//リマインダメールの送信
function sendReminderEmail(data, i) {
  const address = data[i][2];
  const subject = `【東北大学写真部】用紙・印刷代金 支払いのお願い`
  const body = createRemainderEmailBody(data, i);
  const options = {name: `東北大学写真部（自動送信）`};
  MailApp.sendEmail(address, subject, body, options);
}

//リマインダメールの本文を作成
function createRemainderEmailBody(data, i) {
  if (data[i][14] === `はい`) {
    return `${data[i][1]} 様
東北大学写真部です。以下の内容で報告いただいた印刷代の支払いが未完了のままになっています。
早めのお支払いの上、以下のの編集リンクから支払い完了の報告をお願いします。
${data[i][27]}

報告内容
印刷日時：${data[i][0].toLocaleString('ja-JP', { timeZone: 'Asia/Tokyo' })}

1種類目
用途: ${data[i][3]}
用紙の種類: ${data[i][4]}
用紙サイズ: ${data[i][5]}
枚数: ${data[i][6]}枚

2種類目
用途: ${data[i][9]}
用紙の種類: ${data[i][10]}
用紙サイズ: ${data[i][11]}
枚数: ${data[i][12]}枚

3種類目
用途: ${data[i][15]}
用紙の種類: ${data[i][16]}
用紙サイズ: ${data[i][17]}
枚数: ${data[i][18]}枚

用紙代金: ${data[i][24]}円
印刷料金: ${data[i][25]}円
合計: ${data[i][26]}円

-----------------------------------------------------------------------------------------
このメールは自動送信メールです。返信に対応いたしかねますのでご了承ください。
東北大学写真部
-----------------------------------------------------------------------------------------
  `;}

  else if (data[i][8] === `はい`){
    return `${data[i][1]} 様
東北大学写真部です。以下の内容で報告いただいた印刷代の支払いが未完了のままになっています。
早めのお支払いの上、以下のの編集リンクから支払い完了の報告をお願いします。
${data[i][27]}

報告内容
印刷日時：${data[i][0].toLocaleString('ja-JP', { timeZone: 'Asia/Tokyo' })}

1種類目
用途: ${data[i][3]}
用紙の種類: ${data[i][4]}
用紙サイズ: ${data[i][5]}
枚数: ${data[i][6]}枚

2種類目
用途: ${data[i][9]}
用紙の種類: ${data[i][10]}
用紙サイズ: ${data[i][11]}
枚数: ${data[i][12]}枚

用紙代金: ${data[i][24]}円
印刷料金: ${data[i][25]}円
合計: ${data[i][26]}円

-----------------------------------------------------------------------------------------
このメールは自動送信メールです。返信に対応いたしかねますのでご了承ください。
東北大学写真部
-----------------------------------------------------------------------------------------
  `;}

  else {
    return `${data[i][1]} 様
東北大学写真部です。以下の内容で報告いただいた印刷代の支払いが未完了のままになっています。
早めのお支払いの上、以下のの編集リンクから支払い完了の報告をお願いします。
${data[i][27]}

報告内容
印刷日時：${data[i][0].toLocaleString('ja-JP', { timeZone: 'Asia/Tokyo' })}

1種類目
用途: ${data[i][3]}
用紙の種類: ${data[i][4]}
用紙サイズ: ${data[i][5]}
枚数: ${data[i][6]}枚

用紙代金: ${data[i][24]}円
印刷料金: ${data[i][25]}円
合計: ${data[i][26]}円
-----------------------------------------------------------------------------------------
このメールは自動送信メールです。返信に対応いたしかねますのでご了承ください。
東北大学写真部
-----------------------------------------------------------------------------------------
  `;
  }
}

//リマインダ送信レポートを管理者へ送信
function sendReport(log_counter, log_user) {
  //管理者（会計）のメルアド、各自設定すること．
  const address = `hogehoge@dc.tohoku.ac.jp`
  const subject = `【東北大学写真部】リマインダ送信のレポート`
  const body = createReportBody(log_counter, log_user);
  const options = {name: `東北大学写真部（自動送信）`}
  MailApp.sendEmail(address, subject, body, options);
}

function createReportBody(log_counter, log_user) {
  return `リマインダーメールを${log_counter}件送信しました。
以下のユーザに送信しました。
${log_user.join(`\n`)}`
}
