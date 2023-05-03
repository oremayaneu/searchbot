// LINE Messaging APIのチャネルアクセストークン
var LINE_ACCESS_TOKEN = "bQMtf+MIR2Px412iPNZ9aVhUi2V4rwavtcXW6+0qEOr7+7lgIhMgQkJngat6zPxfMG6nH6f70FdNLTo7yTPHnJUuojsT32ya79jRT0PtjSyer+p8SV/bElPpm030CPOYkih9AjJ5xcuJl9v0p5DYkwdB04t89/1O/w1cDnyilFU=";

// スプレッドシートID
var ss = SpreadsheetApp.openById("1L1P3QLbbTgd4gmRcNVE4HQ5kZyITeyrkEi3MAoGbAyo");

// シート名
var sh = ss.getSheetByName("フォームの回答1");




// LINE Messaging APIからPOST送信を受けたときに起動する
// e はJSON文字列
function doPost(e){
  if (typeof e === "undefined"){
    // 動作を終了する
    return;
  } else {
    // JSON文字列をパース(解析)し、変数jsonに格納する
    var json = JSON.parse(e.postData.contents);

    // 変数jsonを関数replyFromSheetに渡し、replyFromSheetを実行する
    replyFromSheet(json)
  }
}
 

 
// 返信用の関数replyFromSheet
// data には変数jsonが代入される
function replyFromSheet(data) {

  // 返信先URL
  var replyUrl = "https://api.line.me/v2/bot/message/reply";
  
  // 受信したメッセージ情報を変数に格納する
  var replyToken　= data.events[0].replyToken; // reply token
  var postMsg = data.events[0].message.text; // ユーザーが送信した語句 
  
  // ここまで基本設定


  
  // 値の抽出
  var sheet = SpreadsheetApp.getActiveSheet();
  
  var lastRow = sheet.getDataRange().getLastRow();
  var Row = lastRow - 1;
               
               // 2行7列から下端行まで4列分（4グループ分）のデータを取得
  var finder = sheet.getRange(2, 7, Row, 4).createTextFinder(postMsg).useRegularExpression(true); // 検索語を含んだセルを発見 // [正規表現を使用した検索]有効 
  var ranges = finder.findAll();  // そのセルの情報を取得
  var rows = [];
  var reply = [];  // bot返信
  // それらのセルの行の番号を配列に追加
  for ( var i = 0; i < ranges.length; i++ ) {
      rows.push(ranges[i].getRow()) ;
  }
  for ( var j = 0; j < ranges.length; j++ ) {  // 見つかった人数分forを回す
      var k = rows[j];  // ヒットした人のうち、j人目のデータの行目をkとする
      var range1 = sheet.getRange(k,2);  // ハンネは2行目
      var answer1 = range1.getValue();  // ハンネを取得
      var range2 = sheet.getRange(k,3);  // line名は3行目
      var answer2 = range2.getValue();  // line名を取得
      var range3 = sheet.getRange(k,12);  // 書きたいかどうかは12行目
      var answer3 = range3.getValue();  // 書きたいかどうかを取得
      reply = reply + "\n\n" + answer1 + "(LINE名:" + answer2 + ")" + "\n:" + answer3;  // 返信に追加していく
  }
  
  var replyText = '「' + postMsg + '」の検索結果です。'  // bot返信の外枠
  
  replyText = replyText + "\n________________________" + reply;  // 外枠に作成したreplyを追加
  
  
  var muse = "高坂穂乃果園田海未南ことり小泉花陽星空凛西木野真姫絢瀬絵里東條希矢澤にこ";
  var aqours = "高海千歌桜内梨子渡辺曜黒澤ルビィ国木田花丸津島善子小原鞠莉黒澤ダイヤ松浦果南";
  var niji = "高咲侑上原歩夢宮下愛優木せつ菜中須かすみ桜坂しずく天王寺璃奈エマヴェルデ近江彼方朝香果林三船栞子鐘嵐珠ミアテイラー";
  var liella = "澁谷かのん嵐千砂都平安名すみれ唐可可葉月恋桜小路きな子米女メイ若菜四季鬼塚夏美ウィーンマルガレーテ";
  
  // 入力に基づき、どのグループの箱推しを検出するか決定
  
  var group = 1;  // 初期値
  var check1=muse.indexOf(postMsg);  // 入力がmuseメンバーと一致
  if(check1 > -1){
    group = 7;  // μ’sは7行目と教える
    var group_n = "μ’s";  // group name = μ’s
  }
  var check2=aqours.indexOf(postMsg);  // 入力がaqoursメンバーと一致
  if(check2 > -1){
    group = 8;
    var group_n = "Aqours";  //group name = Aqours
  }
  var check3=niji.indexOf(postMsg);  //入力がnijiメンバーと一致
  if(check3 > -1){
    group = 9;
    var group_n = "虹ヶ咲";  // group name = 虹ヶ咲
  }
  var check4=liella.indexOf(postMsg);  // 入力がliellaメンバーと一致
  if(check4 > -1){
    group = 10;
    var group_n = "Liella!";  // group name = Liella!
  }
  
  if(group > 6){  // どれかのグループに属していれば7以上なのでスルー
    var finder_g = sheet.getRange(2, group, Row, 1).createTextFinder("箱推し").useRegularExpression(true);  // 教えてもらったグループ（列）内で箱推しの人を発見
    var ranges_g = finder_g.findAll();  //以下同様
    var rows_g = [];
    var reply_g = [];
    // 検索語を含む行番号を配列に入れる
    for ( var l = 0; l < ranges_g.length; l++ ) {
      rows_g.push(ranges_g[l].getRow()) ;
    }
    for ( var m = 0; m < ranges_g.length; m++ ) { 
      var n = rows_g[m];
      var range1_g = sheet.getRange(n,2);
      var answer1_g = range1_g.getValue();
      var range2_g = sheet.getRange(n,3);
      var answer2_g = range2_g.getValue();
      var range3_g = sheet.getRange(n,12);
      var answer3_g = range3_g.getValue();
      reply_g = reply_g + "\n\n" + answer1_g + "(LINE名:" + answer2_g + ")" + "\n:" + answer3_g;
    }
    replyText = replyText + "\n\n" + "【以下" + group_n + "箱推しの方】" + reply_g;  // 箱推しは念の為区別して返信に追加
  }else{  // どこにも属していないなら
    var ranges_g = [];  // ヒットしたデータはなし
  }
  

  if (ranges.length + ranges_g.length <= 0) {  // 推し検索で取得したデータ数 + 箱推し検索で取得したデータ数 <= 0 なら検索語に問題あり
    replyText = '入力に不備があるか、該当結果がありません。'
  }

  
  
  // 以下基本設定
  // LINE messaging apiにJSON形式でデータをPOST 
  // replyするメッセージの定義
  var postData = {
    "replyToken" : replyToken,
    "messages" : [
      {
        "type" : "text",
        "text" : replyText
      }
    ]
  };

  // LINE messaging apiにJSON形式でデータをPOST
  var headers = {
    "Content-Type": "application/json; charset=UTF-8",
    "Authorization": "Bearer " + LINE_ACCESS_TOKEN,
  };

  // POSTオプション作成
  var options = {
    "method" : "post",
    "headers" : headers,
    "payload" : JSON.stringify(postData)
  };
    
  // LINE Messaging APIにデータを送信する
  UrlFetchApp.fetch(replyUrl, options);
}