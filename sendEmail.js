//以後、基本的に編集不要なのでわからない人は触らないように．
//フォームが送信されるとこのmain関数が実行される
function main(e){
  const row = e.range.getRow();//フォームの内容が反映されたシートの行を取得（eトリガー）
  const answer = sheet.getRange(row,1,1,27).getValues()[0];//行の内容を配列に格納
  let editURL = sheet.getRange(row, 28).getValue();

  //編集URLが存在していない場合、編集URLを割り当て
  if(editURL===``){
    allocate_Edit_URL(row);
    editURL = sheet.getRange(row, 28).getValue();
  }

  //メール送信関数の呼び出し
  send_Email(answer, editURL);
}

//スプレッドシートのAB列にGoogleフォームの編集リンクを割り当て
function allocate_Edit_URL(row){

  const form = FormApp.openById(formID);//フォームの取得
  const responses = form.getResponses();//回答の取得
  const latestresponse = responses[responses.length - 1];
  const editURL = latestresponse.getEditResponseUrl();//編集URLの取得

  //AB列＝28列目にリンクを貼り付け
  const abcolumn = 28;

  //URLをAB列に挿入
  sheet.getRange(row, abcolumn).setValue(editURL);
  
}

function send_Email(answer, editURL){
   
  //送信先を取得
  const send_address = answer[2];
  let subject;

  //メールの件名を状況に応じて変更
  if (answer[20] === '完了'){
    subject = '【東北大学写真部】印刷報告控え兼領収書';
  }
  else if (answer[20] === '未完了'){
    subject = '【東北大学写真部】印刷報告控え（支払い未完了）';
  }
  else {
    subject = '【東北大学写真部】印刷報告控え';
  }
  const options = {name: '東北大学写真部（送信専用）'};//差出人の設定
  
  //メール送信
  GmailApp.sendEmail(send_address, subject, createEmailBody(answer, editURL), options);
}

//メールの本文
function createEmailBody(answer, editURL){
  //バグ報告フォームのリンク，状況に応じて消すか変えてね．
  const reportURL = `https://forms.gle/3Zdt5SNgbmC9aQkJ8`;
  
  //if文で自動メールの内容を変更
  if (answer[14]==='はい'){
    return `${answer[1]} 様
印刷の報告ありがとうございます。以下の内容で受け付けました。

提出日時：${answer[0].toLocaleString('ja-JP', { timeZone: 'Asia/Tokyo' })}

1種類目
用途：${answer[3]}
用紙の種類：${answer[4]}
用紙サイズ：${answer[5]}
枚数：${answer[6]}枚

2種類目
用途：${answer[9]}
用紙の種類：${answer[10]}
用紙サイズ：${answer[11]}
枚数：${answer[12]}枚

3種類目
用途：${answer[15]}
用紙の種類：${answer[16]}
用紙サイズ：${answer[17]}
枚数：${answer[18]}枚

用紙代金：${answer[24]}円
印刷料金：${answer[25]}円
合計：${answer[26]}円

支払い状態：${answer[20]}

お支払いがまだ済んでない場合は上記の料金をお支払いの上、下記のリンクから支払い報告をお願いします。
また、報告内容に変更がある場合も下記リンクから変更をお願いします。
${editURL}

メール内容・挙動に不具合・不審な点がある場合は以下のフォームでバグ報告のご協力をお願いします。
${reportURL}

----------------------------------------------------------------------------------------
このメールは自動送信メールです。返信に対応いたしかねますのでご了承ください。
東北大学写真部
----------------------------------------------------------------------------------------
  `
  }

  else if(answer[8]==='はい'){
    return `${answer[1]} 様
印刷の報告ありがとうございます。以下の内容で受け付けました。

提出日時：${answer[0].toLocaleString('ja-JP', { timeZone: 'Asia/Tokyo' })}

1種類目
用途：${answer[3]}
用紙の種類：${answer[4]}
用紙サイズ：${answer[5]}
枚数：${answer[6]}枚

2種類目
用途：${answer[9]}
用紙の種類：${answer[10]}
用紙サイズ：${answer[11]}
枚数：${answer[12]}枚

用紙代金：${answer[24]}円
印刷料金：${answer[25]}円
合計：${answer[26]}円

支払い状態：${answer[20]}

お支払いがまだ済んでない場合は上記の料金をお支払いの上、下記のリンクから支払い報告をお願いします。
また、報告内容に変更がある場合も下記リンクから変更をお願いします。
${editURL}

メール内容・挙動に不具合・不審な点がある場合は以下のフォームでバグ報告のご協力をお願いします。
${reportURL}

----------------------------------------------------------------------------------------
このメールは自動送信メールです。返信に対応いたしかねますのでご了承ください。
東北大学写真部
----------------------------------------------------------------------------------------
  `
  }

  else{
    return `${answer[1]} 様
印刷の報告ありがとうございます。以下の内容で受け付けました。

提出日時：${answer[0].toLocaleString('ja-JP', { timeZone: 'Asia/Tokyo' })}

1種類目
用途：${answer[3]}
用紙の種類：${answer[4]}
用紙サイズ：${answer[5]}
枚数：${answer[6]}枚

用紙代金：${answer[24]}円
印刷料金：${answer[25]}円
合計：${answer[26]}円

支払い状態：${answer[20]}

お支払いがまだ済んでない場合は上記の料金をお支払いの上、下記のリンクから支払い報告をお願いします。
また、報告内容に変更がある場合も下記リンクから変更をお願いします。
${editURL}

メール内容・挙動に不具合・不審な点がある場合は以下のフォームでバグ報告のご協力をお願いします。
${reportURL}

----------------------------------------------------------------------------------------
このメールは自動送信メールです。返信に対応いたしかねますのでご了承ください。
東北大学写真部
----------------------------------------------------------------------------------------
`
  }
}