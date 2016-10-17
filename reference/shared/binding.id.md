
# <a name="binding.id-property"></a>Binding.id プロパティ
オブジェクトの識別子を取得します。

|||
|:-----|:-----|
|**ホスト:**|Access、Excel、Word|
|**最終変更バージョン**|1.1|

```js
var bindingID = bindingObj.id;
```


## <a name="return-value"></a>戻り値

同じ [Document](../../reference/shared/document.md) オブジェクト内にあるバインドの中から、このバインドを一意に識別する文字列。


## <a name="example"></a>例




```js
Office.context.document.bindings.getByIdAsync("myBinding", function (asyncResult) {
    write(asyncResult.value.id);
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```




## <a name="support-details"></a>サポートの詳細


次の表で、大文字 Y は、このメソッドは、対応する Office ホスト アプリケーションでサポートされていることを示します。空のセルは、Office ホスト アプリケーションでこのメソッドをサポートしないことを示します。

Office ホスト アプリケーションとサーバーの要件の詳細については、「[Office アドインを実行するための要件](../../docs/overview/requirements-for-running-office-add-ins.md)」をご覧ください。


**サポートされるホスト (プラットフォーム別)**


||**Windows デスクトップ版 Office**|**Office Online (ブラウザー)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||Y||
|**Excel**|Y|Y|Y|
|**Word**|Y|Y|Y|

|||
|:-----|:-----|
|**最小限のアクセス許可レベル**|[制限あり](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**アドインの種類**|コンテンツ、作業ウィンドウ|
|**ライブラリ**|Office.js|
|**名前空間**|Office|

## <a name="support-history"></a>サポート履歴





****


|**バージョン**|**変更内容**|
|:-----|:-----|
|1.1|Office for iPad における Excel と Word のサポートが追加されました。|
|1.1|
            Access 用アプリ のバインドのサポートが追加されました。|
|1.0|導入|
