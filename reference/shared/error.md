
# Error オブジェクト
非同期的なデータ操作中に発生したエラーに関する特定の情報を提供します。

|||
|:-----|:-----|
|**ホスト:**|Access、Excel、Outlook、PowerPoint、Project、Word|
|**最終変更バージョン**|1.1|

```
asyncResult.error
```


## メンバー


**プロパティ**


|**名前**|**説明**|
|:-----|:-----|
|[code](../../reference/shared/error.code.md)|エラーの数値コードを取得します。|
|[name](../../reference/shared/error.name.md)|エラーの名前を取得します。|
|[message](../../reference/shared/error.message.md)|エラーの詳細な説明を取得します。|

## 注釈

**Error** オブジェクトには [AsyncResult](../../reference/shared/asyncresult.md) オブジェクトからアクセスします。このオブジェクトは、_Document_ オブジェクトの [setSelectedDataAsync](../../reference/shared/document.setselecteddataasync.md) メソッドなど、非同期的なデータ操作の **callback** 引数として渡される関数で返されます。


## 例

次の例では、 **setSelectedDataAsync** メソッドを使用して、選択されたテキストに "Hello World!" と設定します。操作が失敗した場合は、 **Error** オブジェクトの **name** および **message** プロパティの値を表示します。


```js
function setText() {

    Office.context.document.setSelectedDataAsync("Hello World!", {},
        function (asyncResult) {
            if (asyncResult.status === "failed")
            var err = asyncResult.error; 
                write(err.name + ": " + err.message);
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
|1.1|Office for iPad で Excel、PowerPoint、および Word のサポートが追加されました。|
|1.1|Access 用コンテンツのアドインのサポートが追加されました。|
|1.0|導入|
