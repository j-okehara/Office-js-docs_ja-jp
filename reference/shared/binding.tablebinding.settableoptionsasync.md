
# <a name="tablebinding.settableoptionsasync-method"></a>TableBinding.setTableOptionsAsync メソッド
バインド テーブルにおけるテーブル書式設定オプションを更新します。

|||
|:-----|:-----|
|**ホスト:**|Excel|
|**[要件セット](../../docs/overview/specify-office-hosts-and-api-requirements.md)に指定できるもの**|セットには指定できない|
|**追加されたバージョン**|1.1|

```
bindingObj.setTableOptionsAsync(tableOptions [,options] , callback);
```


## <a name="parameters"></a>パラメーター



|**名前**|**型**|**説明**|**サポートのメモ**|
|:-----|:-----|:-----|:-----|
| _tableOptions_|**object**|オブジェクト リテラルは、適用するテーブル オプションを定義するプロパティ名と値のペアのリストです。必須。||
| _options_|**object**|次の[オプションのパラメーター](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)のいずれかを指定します||
| _asyncContext_|**array**、**boolean**、**null**、**number**、**object**、**string**、または **undefined**|変更されずに **AsyncResult** オブジェクトで返される任意の型のユーザー定義項目。||
| _callback_|**object**|コールバックが戻るときに呼び出される関数で、唯一のパラメーターは **AsyncResult** 型です。||

## <a name="callback-value"></a>コールバック値

_callback_ パラメーターに渡した関数が実行されると、[AsyncResult](../../reference/shared/asyncresult.md) オブジェクトを受け取ります。このオブジェクトには、コールバック関数の唯一のパラメーターからアクセスできます。

**goToByIdAsync** メソッドに渡されるコールバック関数で、**AsyncResult** オブジェクトのプロパティを使用して、次の情報を返すことができます。



|**プロパティ**|**使用目的**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|テーブル オプションの設定時には取得するデータやオブジェクトが存在しないため、常に **undefined** を返します。|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|操作の成功または失敗を判断します。|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|操作が失敗した場合、エラーに関する情報を提供する [Error](../../reference/shared/error.md) オブジェクトにアクセスします。|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|ユーザー定義の **object** または値を _asyncContext_ パラメーターとして渡した場合、そのオブジェクトまたは値にアクセスします。|

## <a name="example"></a>例

次の例は、以下の操作を行う方法を示しています。


-  バインド テーブルにおける更新対象の**テーブルの書式設定オプション**を指定する[オブジェクト リテラルを作成します](../../docs/excel/format-tables-in-add-ins-for-excel.md)。
    
-  書式設定のあるオブジェクトを  **tableOptions** パラメーターとして渡して、( `myBinding` の **id** を使用して) 以前バインドされたテーブルに対して _setTableOptions を呼び出します_。
    

```js
function updateTableFormatting(){
    var tableOptions = {bandedRows: true, filterButton: false, style: "TableStyleMedium3"}; 

    Office.select("bindings#myBinding").setTableOptionsAsync(tableOptions, function(asyncResult){});
}
```




## <a name="support-details"></a>サポートの詳細


次の表で、大文字 Y は、このメソッドは、対応する Office ホスト アプリケーションでサポートされていることを示します。空のセルは、Office ホスト アプリケーションでこのメソッドをサポートしないことを示します。

Office ホスト アプリケーションとサーバーの要件の詳細については、「[Office アドインを実行するための要件](../../docs/overview/requirements-for-running-office-add-ins.md)」をご覧ください。


**サポートされるホスト (プラットフォーム別)**


||**Windows デスクトップ版 Office**|**Office Online (ブラウザー)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|Y|Y|Y|

|||
|:-----|:-----|
|**要件セットに指定できるもの**|セットには指定できない。|
|**最小限のアクセス許可レベル**|[WriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**アドインの種類**|コンテンツ、作業ウィンドウ|
|**ライブラリ**|Office.js|
|**名前空間**|Office|

## <a name="support-history"></a>サポート履歴




|**バージョン**|**変更内容**|
|:-----|:-----|
|1.1|Office for iPad の Excel のサポートが追加されました。|
|1.1|導入|
