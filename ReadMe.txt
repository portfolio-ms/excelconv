excelconv.html

動作環境： PC版Chrome、FireFox（IEは不可）
必須：     - template.inc
           - js/jquery-3.2.1.min.js
           - xlsx.full.min.js
利用方法： 単純に動作を見たい場合は、このまま同梱のsample.xlsxファイルを
           2-3の手順に従ってドロップしてください。
           1. template.inc に以下の設定をする
　　　　　　　- SheetName : どのシートを処理の対象にするか
　　　　　　　- readFrom  : シートの何行目から処理を開始するか
　　　　　　　- tmpObj.h  : ヘッダ、/*～*/の間にHTMLタグを書くこと
　　　　　　　- tmpObj.f  : フッタ、/*～*/の間にHTMLタグを書くこと
　　　　　　　- tmpObj.l  : ループ、Excelファイルを一行ごとに読み込んで
　　　　　　　　　　　　　　埋め込む内容。
　　　　　　　　　　　　　　A列（一列目）の内容を埋め込みたい場合は
　　　　　　　　　　　　　　@@1@@ のように記述する。
　　　　　　　　　　　　　　/*～*/の間にHTMLタグを書くこと
　　　　　　　　　　　　　　
　　　　　 2. excelconv.html を ChromeかFireFoxのブラウザで開く
　　　　　 
　　　　　 3. 「ここに Excel ファイルをドロップ」の箇所に処理したい
　　　　　    Excel ファイルをドラッグ＆ドロップする
　　　　　    
　　　　　 4. 一つ下にあるiFrameにHTML、その下のiFrameにHTMLのタグ
　　　　　    が表示されるので、必要に応じて処理する。
　　　　　    
注意事項： 1. 文字化けを起こす場合は、template.incのファイルの
              文字コードをUTF-8にしてください。
           2. 改行などはExcelのシート中に<br>タグをいれるなどして
              対応してください。