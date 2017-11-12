// template.inc
// Excelで読み込ませるシート名および
// ヘッダー・フッター・1行ずつ処理するループの記述を定義する

// シート名の定義
sheetN='Sheet1';

// シートの何行目から読み込むか？
readFrom=2;

var tmpObj={};

// ------------------------------------------------------
// ヘッダーの定義
// 

tmpObj.h = (function() {/*

<!doctype html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<title>TITLE</title>
<meta name="keywords" content="TEST">
<meta name="description" content="TEST">
<meta http-equiv="X-UA-Compatible" content="IE=edge">

<style>
.red{color: #ff0000;}
.pink{color: #ffe0e0;}
.blue{color: #0000ff;}
.black{color: #000;}
.disp ul, .disp li{
  list-style-type: none;
  margin: 0;
  padding: 0
}
.disp li{
  display: block;
  border-bottom: 1px solid #ccc;
}
.disp li:last-child{
  border: none;
}

.tag{
  display: inline-block;
  width: 120px;
}
.item{
  display: inline-block;
 }
</style>

</head>
<body id="body">

コンテンツのヘッダーを記述      
<h1>商品一覧</h1>

*/}).toString().match(/\/\*([^]*)\*\//)[1];


// ------------------------------------------------------
//フッターの定義
tmpObj.f = (function() {/*

コンテンツのフッターを記述

</body>
</html>


*/}).toString().match(/\/\*([^]*)\*\//)[1];

// ------------------------------------------------------
// ループの記述
// 1列目の内容を入れたい箇所に@@1@@のように指定
tmpObj.l = (function() {/*


<h2>@@4@@</h2>
<ul class="disp">
<li><span class="tag">商品カテゴリー</span><span class="item">@@10@@</span></li>
<li><span class="tag">値段</span><span class="item @@7@@">@@6@@</span></li>
<li><span class="tag">コメント</span><span class="item">@@5@@</span></li>
<li><span class="tag">特記事項</span><span class="item">@@8@@<br>（担当：@@9@@）</span></li>

</ul>

*/}).toString().match(/\/\*([^]*)\*\//)[1];


