
# MatrixBinding.columnCount プロパティ
マトリックス データ構造内の列数を整数値で取得します。

|||
|:-----|:-----|
|**ホスト:**|Access、Excel、PowerPoint、Project、Word|
|**[要件セット](../../docs/overview/specify-office-hosts-and-api-requirements.md)に指定できるもの**|MatrixBindings|
|**選択内容の最終変更**|1.1|

```js
var colCount = bindingObj.columnCount;
```


## 戻り値

指定された [MatrixBinding](../../reference/shared/binding.matrixbinding.md) オブジェクト内の列数。


## 例




```js
function showBindingColumnCount() {
    Office.context.document.bindings.getByIdAsync("myBinding", function (asyncResult) {
        write("Column: " + asyncResult.value.columnCount);
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


**サポートされるホスト (プラットフォーム別)**


||**Windows デスクトップ版 Office**|**Office Online (ブラウザー)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|Y|Y|Y|
|**Word**|Y||Y|

|||
|:-----|:-----|
|**要件セットに指定できるもの**|MatrixBindings|
|**最小限のアクセス許可レベル**|[WriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**アプリの種類**|コンテンツ、作業ウィンドウ|
|**ライブラリ**|Office.js|
|**名前空間**|Office|

## サポート履歴

|**変更内容**|**1.1**|
|:-----|:-----|
|1.1|Office for iPad における Excel と Word のサポートが追加されました。|
|1.0|導入|
