
# Error.name プロパティ
エラーの名前を取得します。

|||
|:-----|:-----|
|**ホスト:**|Access、Excel、Outlook、PowerPoint、Project、Word|
|**選択内容の最終変更**|1.1|

```
var errName = asyncResult.error.name;
```


## 戻り値

エラーの名前を示す  **string** 。


## 注釈

**Error** オブジェクトとそのプロパティには、[AsyncResult](../../reference/shared/asyncresult.md) オブジェクトからアクセスします。AsyncResult オブジェクトは、非同期データ操作の _callback_ 引数として渡される関数で返されます。


## 例

エラーをスローさせるため、テーブルまたはマトリックスを選択し、 `setText` 関数を呼び出します。


```js
function setText() {
    Office.context.document.setSelectedDataAsync("Hello World!",
        function (asyncResult) {
            if (asyncResult.status === "failed")
                var error = asyncResult.error;
            write(error.name + ": " + error.message);
        });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```




## サポートの詳細


次の表で、大文字 Y は、このメソッドは、対応する Office ホスト アプリケーションでサポートされていることを示します。空のセルは、Office ホスト アプリケーションでこのメソッドをサポートしないことを示します。

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
|**アドインの種類**|コンテンツ、作業ウィンドウ、Outlook|
|**ライブラリ**|Office.js|
|**名前空間**|Office|

## サポート履歴



****


|**変更内容**|**1.1**|
|:-----|:-----|
|1.1|PowerPoint Online のサポートが追加されました。|
|1.1|Office for iPad で Excel、PowerPoint、および Word のサポートが追加されました。|
|1.1|Access 用コンテンツのアドインのサポートが追加されました。|
|1.0|導入|
