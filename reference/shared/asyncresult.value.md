
# AsyncResult.value プロパティ
この非同期操作のペイロードまたはコンテンツを取得します (ある場合)。

|||
|:-----|:-----|
|**ホスト:**|Access、Excel、Outlook、PowerPoint、Project、Word|
|**最終変更バージョン**|1.1|

```js
var dataValue = asyncResult.value;
```


## 戻り値

非同期呼び出しが実行された時点での要求の値を返します。 


 >**メモ**:  特定の "Async" メソッドに対して **value** プロパティが返す値は、そのメソッドの用途とコンテキストによって異なります。"Async" メソッドに対して **value** プロパティが返す値については、各メソッドについて説明するトピックの「コールバック値」を参照してください。"Async" メソッドの詳細なリストについては、[AsyncResult](../../reference/shared/asyncresult.md) オブジェクトについて説明するトピックの「解説」を参照してください。


## 解説

**AsyncResult** オブジェクトへのアクセスは、_Document_ オブジェクトの [getSelectedDataAsync](../../reference/shared/document.getselecteddataasync.md) および [setSelectedDataAsync](../../reference/shared/document.setselecteddataasync.md) メソッドなど、"Async" メソッドの **callback** パラメーターに引数として渡される関数で行います。


## 例




```js
function getData() {
    Office.context.document.getSelectedDataAsync(Office.CoercionType.Table, function(asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed) {
            write(asyncResult.error.message);
        }
        else {
            write(asyncResult.value);
        }
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

||**Windows デスクトップ版 Office**|**Office Online (ブラウザー)**|**Office for iPad**|**デバイス用 OWA**|**Office for Mac**|
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



|**変更内容**|**1.1**|
|:-----|:-----|
|1.1|PowerPoint Online のサポートが追加されました。|
|1.1|Office for iPad で Excel、PowerPoint、および Word のサポートが追加されました。|
|1.1|Access 用のアドインのサポートが追加されました。|
|1.0|導入|
