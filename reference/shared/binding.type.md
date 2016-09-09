
# Binding.type プロパティ
バインドの種類を取得します。

|||
|:-----|:-----|
|**ホスト:**|Access、Excel、Word|
|**最終変更バージョン**|1.1|

```
var bindingType = bindingObj.type;
```


## 戻り値

[BindingType](../../reference/shared/bindingtype-enumeration.md)。


## 例




```js
Office.context.document.bindings.getByIdAsync("MyBinding", function (asyncResult) { 
    write(asyncResult.value.type); 
}) 

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
|**Word**|Y||Y|

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
|1.1|Office for iPad における Excel と Word のサポートが追加されました。|
|1.1|Access 用のアドインのサポートが追加されました。|
|1.0|導入|
