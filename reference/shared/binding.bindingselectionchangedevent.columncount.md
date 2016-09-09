
# BindingSelectionChangedEventArgs.columnCount プロパティ
選択されている列の数を取得します。

|||
|:-----|:-----|
|**ホスト:**|Access、Excel、Word|
|**最終変更バージョン**|1.1|

```
var colCount = eventArgsObj.columnCount;
```


## 戻り値

選択されている列の数。セルが 1 つだけ選択されている場合は 1 が返されます。


## 注釈

ユーザーの選択範囲が連続していない場合は、バインド内で最後に連続している選択範囲の数が返されます。 

Word では、このプロパティは [BindingType](../../reference/shared/bindingtype-enumeration.md) が "table" のバインドでのみ機能します。バインドの種類が "matrix" の場合は **null** が返されます。また、テーブルに結合セルが含まれている場合は、呼び出しが失敗します。テーブルの構造が均一になっていないと、このプロパティは正しく機能しないからです。


## 例

次の例では、[SelectionChanged](../../reference/shared/binding.bindingselectionchangedevent.md) イベントのイベント ハンドラーを、`myTable` という [id](../../reference/shared/binding.id.md) を持つバインドに追加します。ユーザーが選択範囲を変更すると、ハンドラーは選択範囲内の最初のセルの座標と、選択された行および列の数を表示します。


```js
function addSelectionHandler() {
    Office.context.document.bindings.getByIdAsync("myTable", function (result) {
        result.value.addHandlerAsync("bindingSelectionChanged", myHandler);
    });
}

// Display selection start coordinates and row/column count.
function myHandler(bArgs) {
    write("Selection start row/col: " + bArgs.startRow + "," + bArgs.startColumn);
    write("Selection row count: " + bArgs.rowCount);
    write("Selection col count: " + bArgs.columnCount);
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


| |**Windows デスクトップ版 Office**|**Office Online (ブラウザー)**|**Office for iPad**|
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
|1.1|Access 用コンテンツ アドインの **SelectionChanged** イベントで、イベント ハンドラーの追加と削除ができるようになりました。|
|1.0|導入|