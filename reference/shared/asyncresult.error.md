
# <a name="asyncresult.error-property"></a>AsyncResult.error プロパティ
エラーの説明を提供する **Error** オブジェクトを取得します (エラーが発生した場合)。

|||
|:-----|:-----|
|**ホスト:**|Access、Excel、Outlook、PowerPoint、Project、Word|
|**最終変更バージョン**|1.1|

```js
var errorObj = asyncResult.error;
```


## <a name="return-value"></a>戻り値

**[Error](../../reference/shared/error.md)** オブジェクト。


## <a name="example"></a>例




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


## <a name="support-details"></a>サポートの詳細


次の表で、大文字 Y は、このメソッドは、対応する Office ホスト アプリケーションでサポートされていることを示します。空のセルは、Office ホスト アプリケーションでこのメソッドをサポートしないことを示します。

Office ホスト アプリケーションとサーバーの要件の詳細については、「[Office アドインを実行するための要件](../../docs/overview/requirements-for-running-office-add-ins.md)」をご覧ください。


| |**Windows デスクトップ版 Office**|**Office Online (ブラウザー)**|**Office for iPad**|**デバイス用 OWA**|**Outlook for Mac**|
|:-----|:-----|:-----|:-----|:-----|:-----|
|**Access**||Y||||
|**Excel**|Y|Y|Y|||
|**Outlook**|Y|Y||Y|Y|
|**PowerPoint**|Y|Y|Y|||
|**Project**|Y|||||
|**Word**|Y|Y|Y|||

|||
|:-----|:-----|
|**最小限のアクセス許可レベル**|[制限あり](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**アドインの種類**|コンテンツ、作業ウィンドウ、Outlook|
|**ライブラリ**|Office.js|
|**名前空間**|Office|

## <a name="support-history"></a>サポート履歴



|**バージョン**|**変更内容**|
|:-----|:-----|
|1.1|PowerPoint Online のサポートが追加されました。|
|1.1|Office for iPad で Excel、PowerPoint、および Word のサポートが追加されました。|
|1.1|Access 用のアドインのサポートが追加されました。|
|1.0|導入|
