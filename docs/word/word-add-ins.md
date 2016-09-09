# 最初の Word アドインをビルドする

_適用対象:Word 2016、Word for iPad、Word for Mac_

Word JavaScript API は、Office アプリケーションを拡張するための Office アドイン プログラミング モデルの一部です。このアドイン プログラミング モデルでは、Word の拡張機能をホストするために Web アプリケーションを使用します。どの Web プラットフォームや言語でも Word を拡張できるようになりました。

Word アドインは Word 内で実行され、Word 2016 で使用可能なWord JavaScript API を使用してドキュメントのコンテンツを操作することができます。実際のアドインの作成には、次の 2 つのパーツがあります。1) 任意の場所をホストできる Web アプリケーションと、2) Web アプリケーションがホストされている場所を検出するために Word で使用される[アドイン マニフェスト](../../docs/overview/add-in-manifests.md) (マニフェストが提供する事柄はこれだけではありません。詳細については、「[プログラミングの概要](word-add-ins-programming-overview.md)」をご参照ください) です。

>**Word アドイン = manifest.xml + Web アプリ**

### 設定する
このセクションでは、簡単な Web アプリとアプリ マニフェストを作成します。この Web アプリは、Word 文書に定型句を追加するためのアプリです。

1 - ローカル ドライブに BoilerplateAddin という名前のフォルダーを作成します (たとえば、C:\\BoilerplateAddin)。以下の手順で作成するファイルはすべてこのフォルダーに保存します。

2 - アドイン ビュー用に home.html という名前のファイルを作成します。このアドインには 3 つのボタン (選択されると定型句を追加するもの) を含めます。home.html に次のコードを貼り付けます。

```html
    <!DOCTYPE html>
    <html>
      <head>
        <meta charset="UTF-8" />
        <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
        <title>Boilerplate text app</title>
        <script src="https://ajax.aspnetcdn.com/ajax/jQuery/jquery-2.1.4.min.js"></script>
        <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
        <script src="home.js" type="text/javascript"></script>
        </head>
        <body>
            <div>
                    <h1>Welcome</h1>
            </div>
            <div>
                    <p>This sample shows how to add boilerplate text to a document by using the Word JavaScript API.</p>
                    <br />
                    <h3>Try it out</h3>
                    <button id="emerson">Add quote from Ralph Waldo Emerson</button>
                    <button id="checkhov">Add quote from Anton Chekhov</button>
                    <button id="proverb">Add Chinese proverb</button>
            </div>
            <h3><div id="supportedVersion"/></h3>
        </body>
    </html>
```

3 - home.js という名前のファイルを作成して、そのファイルに次のコードを貼り付けます。これには、初期化コードと、Word 文書を変更するためのアドイン コードのすべてが含まれています。このコードは、Word 文書のカーソル位置、または選択部分に基づいて、テキストを挿入します。

```javascript
    (function () {
        "use strict";

        // The initialize function is run each time the page is loaded.
        Office.initialize = function (reason) {
            $(document).ready(function () {

                // Use this to check whether the API is supported in the Word client.
                if (Office.context.requirements.isSetSupported('WordApi', 1.1)) {
                    // Do something that is only available via the new APIs
                    $('#emerson').click(insertEmersonQuoteAtSelection);
                    $('#checkhov').click(insertChekhovQuoteAtTheBeginning);
                    $('#proverb').click(insertChineseProverbAtTheEnd);
                    $('#supportedVersion').html('This code is using Word 2016 or greater.');
                }
                else {
                    // Just letting you know that this code will not work with your version of Word.
                    $('#supportedVersion').html('This code requires Word 2016 or greater.');
                }
            });
        };

        function insertEmersonQuoteAtSelection() {
            Word.run(function (context) {

                // Create a proxy object for the document.
                var thisDocument = context.document;

                // Queue a command to get the current selection.
                // Create a proxy range object for the selection.
                var range = thisDocument.getSelection();

                // Queue a command to replace the selected text.
                range.insertText('"Hitch your wagon to a star."\n', Word.InsertLocation.replace);

                // Synchronize the document state by executing the queued commands,
                // and return a promise to indicate task completion.
                return context.sync().then(function () {
                    console.log('Added a quote from Ralph Waldo Emerson.');
                });
            })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
        }

        function insertChekhovQuoteAtTheBeginning() {
            Word.run(function (context) {

                // Create a proxy object for the document body.
                var body = context.document.body;

                // Queue a command to insert text at the start of the document body.
                body.insertText('"Knowledge is of no value unless you put it into practice."\n', Word.InsertLocation.start);

                // Synchronize the document state by executing the queued commands,
                // and return a promise to indicate task completion.
                return context.sync().then(function () {
                    console.log('Added a quote from Anton Chekhov.');
                });
            })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
        }

        function insertChineseProverbAtTheEnd() {
            Word.run(function (context) {

                // Create a proxy object for the document body.
                var body = context.document.body;

                // Queue a command to insert text at the end of the document body.
                body.insertText('"To know the road ahead, ask those coming back."\n', Word.InsertLocation.end);

                // Synchronize the document state by executing the queued commands,
                // and return a promise to indicate task completion.
                return context.sync().then(function () {
                    console.log('Added a quote from a Chinese proverb.');
                });
            })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
        }
    })();
```

4 - BoilerplateManifest.xml という名前の XML ファイルを作成して、このファイルに次のコードを貼り付けます。これは、場所または表示名などのアドインに関する情報を検出するために Word が使用するマニフェスト ファイルです。
```xml
<?xml version="1.0" encoding="UTF-8"?>
    <OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
               xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
               xsi:type="TaskPaneApp">
        <Id>2b88100c-656e-4bab-9f1e-f6731d86e464</Id>
        <Version>1.0.0.0</Version>
        <ProviderName>Microsoft</ProviderName>
        <DefaultLocale>en-US</DefaultLocale>
        <DisplayName DefaultValue="Boilerplate content" />
        <Description DefaultValue="Insert boilerplate content into a Word document." />
        <Hosts>
            <Host Name="Document"/>
        </Hosts>
        <DefaultSettings>
            <SourceLocation DefaultValue="\\MyShare\boilerplate\home.html" />
        </DefaultSettings>
        <Permissions>ReadWriteDocument</Permissions>
    </OfficeApp>
```

5 - GUID を生成して、<code>OfficeApp/Id</code> 要素の値を GUID に置き換えます。

6 - すべてのファイルを保存します。 これで、最初の Word アドインが作成できました。

7- home.js、home.html、および BoilerplateManifest.xml を[ネットワーク上の共有フォルダーにコピーするか](https://technet.microsoft.com/en-us/library/cc770880.aspx) (Windows)、ローカル サーバーにホストします (Mac)。

8 - BoilerplateManifest.xml の[SourceLocation](../../reference/manifest/sourcelocation.md) を編集して、home.html の場所を指すようにします。

この時点で、初めてのアドインが配置されたことになります。 次に、Word がアドインを検索する場所を認識できるようにする必要があります。

#### Word 2016 for Windows で試してみる

1. Word を起動し、ドキュメントを開きます。
2. [**ファイル**] タブを選択し、[**オプション**] を選択します。
3. [**セキュリティ センター**] を選択し、[**セキュリティ センターの設定**] ボタンを選択します。
4. **[信頼されているアドイン カタログ]** を選択します。
5. **[カタログの URL]** ボックスに、BoilerplateManifest.xml があるフォルダー共有へのパスを入力して、**[カタログの追加]** を選択します。
6. [**メニューに表示する**] チェック ボックスをオンにし、[**OK**] を選択します。
7. これらの設定が Office を次回起動したときに適用されることを示すメッセージが表示されます。Word を終了して、再起動します。

次は、作成したアドインを実行します。次の手順を実行して、動作を確認します。

1. Word 文書を開きます。
2. Word 2016 の**[挿入]** タブで、**[マイ アドイン]** を選択します。
3. **[共有フォルダー]** タブを選択します。
4. **[定型コンテンツ]**、**[挿入]** の順に選択します。
5. アドインが作業ウィンドウに読み込まれます。読み込まれたときの状態については、図 1 を参照してください。
6. 定型句を Word 文書に入力するボタンを選択します。


### Word 2016 for Mac で試してみる

次は、作成したアドインを実行します。次の手順を実行して、動作を確認します。

1. Users/Library/Containers/com.microsoft.word/Data/Documents/ に「wef」というフォルダーを作成します。
2. マニフェスト BoilerplateManifest.xml を wef フォルダー (Users/Library/Containers/com.microsoft.word/Data/Documents/wef) に保存します。
3. Mac で Word 2016 を開き、[挿入] タブ > [マイ アドイン] ドロップ ダウンをクリックします。 ドロップ ダウンにアドインがリスト表示されるはずです。 選択すると、アドインが読み込まれます。

__図 1.Word に読み込まれた定型句のコンテンツ アドイン__
![定型句のアドインが読み込まれた Word アプリケーションのイメージ。](../../images/boilerplateAddin.png "定型句を入力するための単純な Word アドイン。")

## フィードバックをお寄せください

お客様からのフィードバックを重視しています。

* ドキュメントを確認していだだき、ドキュメントに関する質問や問題があれば、[問題を送信](https://github.com/OfficeDev/office-js-docs/issues)してお知らせください。
* プログラミングの経験と、今後のバージョンまたはコード サンプルなどで希望されるものについてお知らせください。ご提案やアイデアの入力には、[UserVoice サイト](http://officespdev.uservoice.com/)をご使用ください。

## その他のリソース

* [Office アドインを使う](https://dev.office.com/getting-started/addins?product=word)
* [Word add-ins on GitHub](https://github.com/OfficeDev?utf8=%E2%9C%93&query=Word)
