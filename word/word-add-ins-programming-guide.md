# Word アドインのプログラミングの概要

_適用対象:Word 2016、Word for iPad、Word for Mac_

Word 2016 では、Word オブジェクトを操作するための新しいオブジェクト モデルが導入されています。このオブジェクト モデルは、Word のアドインを作成するために、Office.js で提供されている既存のオブジェクト モデルに追加されたものです。このオブジェクト モデルには、Web アプリケーションでホストされた JavaScript を介してアクセスします。

## マニフェスト

この新しい Word アドインの JavaScript API は、Office 2013 のアドイン モデルで使用されていたものと同じマニフェスト形式を使用します。マニフェストは、アドインがホストされている場所、表示方法、アクセス許可、その他の情報について説明します。[アドイン マニフェスト](https://msdn.microsoft.com/en-us/library/office/fp161044.aspx)をカスタマイズする方法の詳細について説明します。 

Word アドイン マニフェストを発行するには、いくつかのオプションがあります。ネットワーク共有、アプリ カタログ、または Office ストアに [Office アドインを発行する](https://msdn.microsoft.com/EN-US/library/office/fp123515.aspx)方法をご参照ください。

## JavaScript API for Word について

JavaScript API for Word は Office.js で読み込まれます。Word 文書の内容を操作する一連のコマンドをキューに入れるために使用する一連の JavaScript のプロキシ オブジェクトが用意されてす。これらのコマンドは、バッチで実行されます。バッチの結果として、コンテンツの挿入や Word オブジェクトと JavaScript プロキシ オブジェクトの同期などアクションが Word 文書に対して実行されます。 

### アドインの実行

アドインを実行する際に必要なものを見てみましょう。すべてのアドインには、Office.initialize イベント ハンドラーが必要です。アドインの初期化の詳細については、「[API について](https://msdn.microsoft.com/EN-US/library/fp160953.aspx)」をご参照ください。  

Word アドインは、Word.run() メソッドに関数を渡すことで実行されます。run メソッドに渡される関数には、context 引数が必要です。この[コンテキスト オブジェクト](word-add-ins-javascript-reference/requestcontext.md)は、Office オブジェクトから取得するコンテキスト オブジェクトとは異なりますが、どちらも Word ランタイム環境と対話するという同じ目的に使用されます。コンテキスト オブジェクトを使用して、Word JavaScript オブジェクト モデルにアクセスできます。基本的な Word アドインのコメントとコードを確認してみましょう。

**例 1.Word アドインの初期設定と実行**

```javascript
    (function () {
        "use strict";

        // The initialize event handler is run each time the page is loaded.
        Office.initialize = function (reason) {
            
            // Checks for the DOM to load using the jQuery ready function.
            $(document).ready(function () {
                // Set your initialization code. You can use the reason 
                // argument to determine how the add-in was loaded.
                // You can also load saved settings from the Office object.
            });
        };

        // Run a batch operation against the Word object model.
        // Use the context argument to get access to the Word document.
        Word.run(function (context) {

            // Create a proxy object for the document.
            var thisDocument = context.document;
        })
    })();
```

例 1. は Word アドインの作成に必要な基本的なコードを示しています。これを使用して Office.js を初期化します。またこれには Word 文書を操作するための run メソッドが含まれています。

### プロキシ オブジェクト

Word JavaScript オブジェクト モデルは、Word 内のオブジェクトと緩く結合されています。Word JavaScript オブジェクトは、Word 文書内の実際のオブジェクトのプロキシ オブジェクトです。プロキシ オブジェクトで実行されたすべてのアクションは、Word では認識されません。また、Word 文書の状態は、ドキュメントの状態が同期されるまでプロキシ オブジェクトで認識されません。ドキュメントの状態は、context.sync() の実行時に同期されます。sync() メソッドは原則的に各プロキシ オブジェクトに対してキューの一連のコマンドを実行します。例 2 は、本文のプロキシ オブジェクトと、その本文プロキシ オブジェクトにテキスト プロパティを読み込むためのキューに登録済みのコマンドの作成、さらに Word 文書内の本文と本文プロキシ オブジェクトとの同期を示しています。 

**例 2.ドキュメント本文と本文のプロキシオブジェクトを同期する。**

```javascript
    // Run a batch operation against the Word object model.
    Word.run(function (context) {

        // Create a proxy object for the document body.
        // The body object hasn't been set with any property values. 
        var body = context.document.body;

        // Queue a command to load the text property for the proxy document body object.
        context.load(body, 'text');

        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log("Body contents: " + body.text);
        });  
    })
```

### コマンド キュー

Word のプロキシ オブジェクトには、オブジェクト モデルにアクセスして更新するためのメソッドが用意されています。これらのメソッドは、バッチでキューに入れられた順序で順番に実行されます。コマンドのバッチは、context.sync() 呼び出しが行われる前に作成されます。そのコンテキストを使用するすべてのオブジェクトのキューに入れられたコマンドがすべて実行されます。  

例 3 では、コマンドのキューが機能する仕組みを示しています。context.sync() が呼び出されると、まず最初に、本文テキストを[読み込むコマンド](Word%20Add-ins%20JavaScript%20Reference/loadoption.md)が Word で実行されます。次に、Word の本文にテキストを挿入するコマンドが生成されます。その結果は本文のプロキシ オブジェクトに返されます。Word JavaScript の body.text プロパティの値は、テキストが Word 文書に挿入される<u>前</u>の Word 文書本文の値になります。  

**例 3.コマンドのバッチを実行する。**

```javascript
    // Run a batch operation against the Word object model.
    Word.run(function (context) {

        // Create a proxy object for the document body.
        var body = context.document.body;

        // Queue a command to load the text in the proxy body object.
        context.load(body, 'text');

        // Queue a command to insert text into the end of the Word document body.
        body.insertText('This is text inserted after loading the body.text property',
                        Word.InsertLocation.end);

        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log("Body contents: " + body.text);
        });  
    })
```

## フィードバックをお寄せください

お客様からのフィードバックを重視しています。 

* ドキュメントを確認していだだき、ドキュメントに関する質問や問題があれば、直接このリポジトリに[問題を送信](https://github.com/OfficeDev/office-js-docs/issues)してお知らせください。
* プログラミングの経験と、今後のバージョン、コード サンプルなどで希望されるものについてお知らせください。ご提案やアイデアの入力には、[このサイト](http://officespdev.uservoice.com/)をご使用ください。


## その他の技術情報

* [Word アドイン](word-add-ins.md)
* [Word アドインの JavaScript リファレンス](word-add-ins-javascript-reference.md)
* [Office アドイン](https://msdn.microsoft.com/en-us/library/office/jj220060.aspx)
* [Office アドインを使う](http://dev.office.com/getting-started/addins)
* &lt;a herf="https://github.com/OfficeDev?utf8=%E2%9C%93&amp;query=Word"&gt;GitHub の Word アドイン&lt;/a&gt;
* [Word のスニペット エクスプローラー](http://officesnippetexplorer.azurewebsites.net/#/snippets/word)

