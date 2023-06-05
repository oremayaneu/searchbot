// LINE Messaging APIのチャネルアクセストークン
var LINE_ACCESS_TOKEN = " ";

// スプレッドシートID
var ss = SpreadsheetApp.openById(" ");

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
  var replyToken = data.events[0].replyToken; // reply token
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
  
  
  var muse = ["高坂穂乃果","園田海未","南ことり","小泉花陽","星空凛","西木野真姫","絢瀬絵里","東條希","矢澤にこ"];
  var aqours = ["高海千歌","桜内梨子","渡辺曜","黒澤ルビィ","国木田花丸","津島善子","小原鞠莉","黒澤ダイヤ","松浦果南"];
  var niji = ["高咲侑","上原歩夢","宮下愛","優木せつ菜","中須かすみ","桜坂しずく","天王寺璃奈","エマヴェルデ","近江彼方","朝香果林","三船栞子","鐘嵐珠","ミアテイラー"];
  var liella = ["澁谷かのん","嵐千砂都","平安名すみれ","唐可可","葉月恋","桜小路きな子","米女メイ","若菜四季","鬼塚夏美","ウィーンマルガレーテ"];
  
  // 入力に基づき、どのグループの箱推しを検出するか決定
  
  var group = 1;  // 初期値
  var group_n = "";


  muse.forEach(chara_name => {
    if (chara_name.indexOf(postMsg) > -1) { // 検索ワードを含む要素を発見したらμ’s
        group = 7;  // μ’sは7行目と教える
        group_n = "μ’s";  // group name = μ’s
    }
  });

  // group未定なら次はAqoursを探索
  if (group === 1) {
    aqours.forEach(chara_name => {
      if (chara_name.indexOf(postMsg) > -1) {
          group = 8;
          group_n = "Aqours";
      }
    });
  }

  // group未定なら次は虹ヶ咲を探索
  if (group === 1) {
    niji.forEach(chara_name => {
      if (chara_name.indexOf(postMsg) > -1) {
          group = 9;
          group_n = "虹ヶ咲";
      }
    });
  }

  // group未定なら次はLiellaを探索
  if (group === 1) {
    liella.forEach(chara_name => {
      if (chara_name.indexOf(postMsg) > -1) {
          group = 10;
          group_n = "Liella!";
      }
    });
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