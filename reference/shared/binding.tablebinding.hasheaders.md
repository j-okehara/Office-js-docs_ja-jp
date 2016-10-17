
# <a name="tablebinding.hasheaders-property"></a>TableBinding.hasHeaders プロパティ
テーブルにヘッダーがあるかどうかを取得します。

|||
|:-----|:-----|
|**ホスト:**|Access、Excel、PowerPoint、Project、Word|
|**[要件セット](../../docs/overview/specify-office-hosts-and-api-requirements.md)に指定できるもの**|TableBindings|
|**選択内容の最終変更**|1.1|

```
var colCount = bindingObj.hasHeaders;
```


## <a name="return-value"></a>戻り値

指定された [TableBinding](../../reference/shared/binding.tablebinding.md) にヘッダーがある場合は、**true** を返します。ヘッダーがない場合は、**false** を返します。


## <a name="example"></a>例




```js
function showBindingHasHeaders() {
    Office.context.document.bindings.getByIdAsync("myBinding", function (asyncResult) {
        write("Binding has headers: " + asyncResult.value.hasHeaders);
    });
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```




## <a name="support-details"></a>サポートの詳細


次の表で、大文字 Y は、このプロパティは、対応する Office ホスト アプリケーションでサポートされていることを示します。空のセルは、Office ホスト アプリケーションでこのプロパティをサポートしないことを示します。

Office ホスト アプリケーションとサーバーの要件の詳細については、「[Office アドインを実行するための要件](../../docs/overview/requirements-for-running-office-add-ins.md)」をご覧ください。


**サポートされるホスト (プラットフォーム別)**


||**Windows デスクトップ版 Office**|**Office Online (ブラウザー)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||Y||
|**Excel**|Y|Y|Y|
|**Word**|Y|Y|Y|

|||
|:-----|:-----|
|**要件セットに指定できるもの**|TableBindings|
|**最小限のアクセス許可レベル**|[ReadDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**アドインの種類**|コンテンツ、作業ウィンドウ|
|**ライブラリ**|Office.js|
|**名前空間**|Office|

## <a name="support-history"></a>サポート履歴





****


|**バージョン**|**変更内容**|
|:-----|:-----|
|1.1|Office for iPad における Excel と Word のサポートが追加されました。|
|1.1|
            Access 用アプリにこのイベントのサポートが追加されました。|
|1.0|導入|
