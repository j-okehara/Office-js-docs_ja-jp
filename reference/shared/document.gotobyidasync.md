
# Document.goToByIdAsync メソッド
ドキュメント内の指定されたオブジェクトまたは場所に移動します。

|||
|:-----|:-----|
|**ホスト:**|Excel、PowerPoint、Word|
|**要件セットに指定できるもの**|セットには指定できない|
|**で追加**|1.1|

```js
Office.context.document.goToByIdAsync(id, goToType, [,options], callback);
```


## パラメーター



|**名前**|**型**|**説明**|**サポートのメモ**|
|:-----|:-----|:-----|:-----|
| _id_|**string** または **number**|移動先のオブジェクトまたは場所の識別子です。必ず指定します。||
| _goToType_|[GoToType](../../reference/shared/gototype-enumeration.md)|移動先の場所の型です。必ず指定します。||
| _オプション_|**object**|次の[オプションのパラメーター](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)のいずれかを指定します||
| _selectionMode_|[SelectionMode](../../reference/shared/selectionmode-enumeration.md)|_id_ パラメーターで指定した場所が選択されている (強調表示されている) かどうかを指定します。|**Excel の場合:**<br/> **Office.SelectionMode.Selected** は、バインド内のすべてのコンテンツ、または名前付きアイテムを選択します。 <br/>**Office.SelectionMode.None** では、テキスト バインドの場合は、セルを選択します。マトリックス バインド、テーブル バインド、および名前付きアイテムの場合は、最初のデータ セルを選択します (テーブルの見出し行の最初のセルではありません)。<br/><br/> **PowerPoint の場合:**<br/> **Office.SelectionMode.Selected** は、スライド タイトルまたはスライドの最初のテキストボックスを選択します。<br/> **Office.SelectionMode.None** は何も選択しません。<br/><br/> **Word の場合:**<br/> **Office.SelectionMode.Selected** は、バインド内のすべてのコンテンツを選択します。 <br/>**Office.SelectionMode.None** では、テキスト バインドの場合はテキストの最初までカーソルを移動します。マトリックス バインドとテーブル バインドの場合は、最初のデータ セルを選択します (テーブルの見出し行の最初のセルではありません)。|
| _asyncContext_|**array**、 **boolean**、 **null**、 **number**、 **object** 、 **string**、または  **undefined**|変更されずに  **AsyncResult** オブジェクトで返される任意の型のユーザー定義項目。||
| _callback_|**object**|コールバックが戻るときに呼び出される関数で、唯一のパラメーターは  **AsyncResult** 型です。||

## コールバック値

_callback_ パラメーターに渡した関数が実行されると、[AsyncResult](../../reference/shared/asyncresult.md) オブジェクトを受け取ります。このオブジェクトには、コールバック関数の唯一のパラメーターからアクセスできます。

**goToByIdAsync** メソッドに渡されるコールバック関数で、**AsyncResult** オブジェクトのプロパティを使用して、次の情報を返すことができます。



|**プロパティ**|**使用目的**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|現在のビューを返します。|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|操作の成功または失敗を判断します。|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|操作が失敗した場合、エラーに関する情報を提供する [Error](../../reference/shared/error.md) オブジェクトにアクセスします。|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|ユーザー定義の  **object** または値を _asyncContext_ パラメーターとして渡した場合、そのオブジェクトまたは値にアクセスします。|

## 注釈

PowerPoint では、 **マスター表示**で  **goToByIdAsync** メソッドがサポートされません。


## 例

 **ID でバインドに移動する (Word と Excel)**

次の例は、以下の操作を行う方法を示しています。


-  **addFromSelectionAsync** メソッドを使用して、サンプル バインドとして使用する[テーブル バインドを作成します](../../reference/shared/bindings.addfromselectionasync.md)。
    
-  **そのバインドを**移動先のバインドとして指定します。
    
-  操作の状態を返す **匿名のコールバック関数を** _goToByIdAsync_ メソッドの **callback**パラメーターに渡します。
    
-  アドインのページに **値を表示します**。
    



```js
function gotoBinding() {
    //Create a new table binding for the selected table.
    Office.context.document.bindings.addFromSelectionAsync("table",{ id: "MyTableBinding" }, function (asyncResult) {
    if (asyncResult.status == "failed") {
              showMessage("Action failed with error: " + asyncResult.error.message);
           }
           else {
              showMessage("Added new binding with type: " + asyncResult.value.type +" and id: " + asyncResult.value.id);
           }
    });

    //Go to binding by id.
    Office.context.document.goToByIdAsync("MyTableBinding", Office.GoToType.Binding, function (asyncResult) {
        if (asyncResult.status == "failed") {
            showMessage("Action failed with error: " + asyncResult.error.message);
        }
        else {
            showMessage("Navigation successful");
        }
    });
}
```



 **スプレッドシート内のテーブルに移動する (Excel)**

次の例は、以下の操作を行う方法を示しています。


-  **テーブルの名前を**移動先のテーブルとして指定します。
    
-  操作の状態を返す **匿名のコールバック関数を** _goToByIdAsync_ メソッドの **callback**パラメーターに渡します。
    
-  アドインのページに **値を表示します**。
    



```js
function goToTable() {
    Office.context.document.goToByIdAsync("Table1", Office.GoToType.NamedItem, function (asyncResult) {
        if (asyncResult.status == "failed") {
            showMessage("Action failed with error: " + asyncResult.error.message);
        }
        else {
            showMessage("Navigation successful");
        }
    });
}
```



 **ID で現在選択されているスライドに移動する (PowerPoint)**

次の例は、以下の操作を行う方法を示しています。


-  **getSelectedDataAsync** メソッドを使用して、現在選択されているスライドの[ ID を取得します](../../reference/shared/document.getselecteddataasync.md)。
    
-  **返された ID** を移動先のスライドとして指定します。
    
-  操作の状態を返す **匿名のコールバック関数を** _goToByIdAsync_ メソッドの **callback**パラメーターに渡します。
    
-  アドインのページに  `asyncResult.value` から返された、文字列に変換された JSON オブジェクトの **値を表示します**。これには選択されたスライドに関する情報が含まれています。
    



```js
var firstSlideId = 0;
function gotoSelectedSlide() {
    //Get currently selected slide's id
    Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange, function (asyncResult) {
        if (asyncResult.status == "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
        else {
            firstSlideId = asyncResult.value.slides[0].id;
            app.showNotification(JSON.stringify(asyncResult.value));
        }
    });
    //Go to slide by id.
    Office.context.document.goToByIdAsync(firstSlideId, Office.GoToType.Slide, function (asyncResult) {
        if (asyncResult.status == "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
        else {
            app.showNotification("Navigation successful");
        }
    });
}
```



 **インデックスでスライドに移動する (PowerPoint)**

次の例は、以下の操作を行う方法を示しています。


-  移動先のスライド (最初、最後、前、または次のスライド) の **インデックスを指定します**。
    
-  操作の状態を返す **匿名のコールバック関数を** _goToByIdAsync_ メソッドの **callback**パラメーターに渡します。
    
-  アドインのページに **値を表示します**。
    



```js
function goToSlideByIndex() {
    var goToFirst = Office.Index.First;
    var goToLast = Office.Index.Last;
    var goToPrevious = Office.Index.Previous;
    var goToNext = Office.Index.Next;

    Office.context.document.goToByIdAsync(goToNext, Office.GoToType.Index, function (asyncResult) {
        if (asyncResult.status == "failed") {
            showMessage("Action failed with error: " + asyncResult.error.message);
        }
        else {
            showMessage("Navigation successful");
        }
    });
}
```




## サポートの詳細


次の表で、大文字 Y は、このメソッドは、対応する Office ホスト アプリケーションでサポートされていることを示します。空のセルは、Office ホスト アプリケーションでこのメソッドをサポートしないことを示します。

Office ホスト アプリケーションとサーバーの要件の詳細については、「[Office アドインを実行するための要件](../../docs/overview/requirements-for-running-office-add-ins.md)」をご覧ください。


**サポートされるホスト (プラットフォーム別)**


||**Windows デスクトップ版 Office**|**Office Online (ブラウザー)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|Y|Y|Y|
|**PowerPoint**|Y|Y|Y|
|**Word**|Y||Y|

|||
|:-----|:-----|
|**要件セットに指定できるもの**|セットには指定できない|
|**最小限のアクセス許可レベル**|[ReadDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**アプリの種類**|コンテンツ、作業ウィンドウ|
|**ライブラリ**|Office.js|
|**名前空間**|Office|

## サポート履歴



|**変更内容**|**1.1**|
|:-----|:-----|
|1.1|PowerPoint Online のサポートが追加されました。|
|1.1|Office for iPad で Excel、PowerPoint、および Word のサポートが追加されました。|
|1.1|導入|
