
# Document.url プロパティ
ホスト アプリケーションが現在開いているドキュメントの URL を取得します。

|||
|:-----|:-----|
|**ホスト:**|Access、Excel、Project、Word|
|**最終変更バージョン**|1.1|

```
var docUrl = Office.context.document.url;
```


## 戻り値

ドキュメントの URL。URL が利用できない場合は  **null** を返します。


## 解説

 **重要:** **url** プロパティは、ドキュメントの名前や保存場所に個人を特定できる情報 (PII) が含まれている可能性がある情報を返します。この情報を保存するか送信する必要がある場合は、必ず暗号化された形式で実行してください。


## 例




```
function displayDocumentUrl() {
    write(Office.context.document.url);
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```




## サポートの詳細


次の表で、大文字 Y は、このプロパティは、対応する Office ホスト アプリケーションでサポートされていることを示します。空のセルは、Office ホスト アプリケーションでこのプロパティをサポートしないことを示します。

Office ホスト アプリケーションとサーバーの要件の詳細については、「[Office アドインを実行するための要件](../../docs/overview/requirements-for-running-office-add-ins.md)」をご覧ください。


**サポートされるホスト (プラットフォーム別)**


||**Windows デスクトップ版 Office**|**Office Online (ブラウザー)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||Y||
|**Excel**|Y|Y|Y|
|**Project**|Y|||
|**Word**|Y|Y|Y|

|||
|:-----|:-----|
|**最小限のアクセス許可レベル**|[Restricted](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**アプリの種類**|コンテンツ、作業ウィンドウ|
|**ライブラリ**|Office.js|
|**名前空間**|Office|

## サポート履歴





****


|**変更内容**|**1.1**|
|:-----|:-----|
|1.1|Word Online のサポートが追加されました。|
|1.1|Office for iPad で Excel、PowerPoint、および Word のサポートが追加されました。|
|1.1|Access 用コンテンツのアドインのサポートが追加されました。|
|1.0|導入|
