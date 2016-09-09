
# Office.initialize イベント
ランタイム環境が読み込まれ、アプリケーションやホストされたドキュメントを対話操作するアドインの準備ができたときに発生します。 

|||
|:-----|:-----|
|**ホスト:**|Access、Excel、Outlook、PowerPoint、Project、Word|
|**最終変更バージョン**|1.1|

```js
Office.initialize = function (reason) {/* initialization code */}
```


## 注釈

_initialize_ イベント リスナー関数の **reason** パラメーターは、初期化の実行方法を指定する [InitializationReason](../../reference/shared/initializationreason-enumeration.md) 列挙値を返します。作業ウィンドウ アドインまたはコンテンツ アドインは、次の 2 つの場合に初期化できます。


- ユーザーが Office ホスト アプリケーションのリボンの [ **挿入**] タブにある [ **アドイン**] ドロップダウン リストの [ **最近使用したアドイン**] セクションから、または [ **アドインの挿入**] ダイアログ ボックスからアドインを挿入した場合。
    
- 既にアドインが含まれているドキュメントをユーザーが開いた場合。
    

 >**メモ**: **initialize** イベント リスナー関数の reason パラメーターは、作業ウィンドウ アドインとコンテンツ アドインの **InitializationReason** 列挙値のみを返し、Outlook アドインの値は返しません。


## 例

**InitializationEnumeration** の値を使用すると、アドインが初めて挿入された場合と、アドインが既にドキュメントの一部になっている場合とで、異なるロジックを実装できます。次の例は、_reason_ パラメーターの値を使用して、作業ウィンドウ アドインまたはコンテンツアドインが初期化された方法を表示する単純なロジックを示しています。


```js
Office.initialize = function (reason) {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, code specific to the add-in can run.
    // Display initialization reason.
    if (reason == "inserted")
    write("The add-in was just inserted.");

    if (reason == "documentOpened")
    write("The add-in is already part of the document.");
    });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```




## サポートの詳細


次の表で、大文字 Y は、このイベントは、対応する Office ホスト アプリケーションでサポートされていることを示します。空のセルは、Office ホスト アプリケーションがこのイベントをサポートしないことを示します。

Office ホスト アプリケーションとサーバーの要件の詳細については、「[Office アドインを実行するための要件](../../docs/overview/requirements-for-running-office-add-ins.md)」をご覧ください。


||**Windows デスクトップ版 Office**|**Office Online (ブラウザー)**|**Office for iPad**|**デバイス用 OWA**|**Outlook for Mac**|
|:-----|:-----|:-----|:-----|:-----|:-----|
|**Access**||Y||||
|**Excel**|Y|Y|Y|||
|**Outlook**|Y|Y||Y|Y|
|**PowerPoint**|Y|Y|Y|||
|**Project**|Y|||||
|**Word**|Y|Y|Y|||

|||
|:-----|:-----|
|**最小限のアクセス許可レベル**|[Restricted](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**アドインの種類**|コンテンツ、Outlook、作業ウィンドウ|
|**ライブラリ**|Office.js|
|**名前空間**|Office|

## サポート履歴




|**変更内容**|**1.1**|
|:-----|:-----|
|1.1|PowerPoint Online のサポートが追加されました。|
|1.1|Office for iPad で Excel、PowerPoint、および Word のサポートが追加されました。|
|1.1|Access 用コンテンツ アドインの初期化のサポートが追加されました。|
|1.0|導入|
