
# <a name="tablebinding.setformatsasync-method"></a>TableBinding.setFormatsAsync メソッド
バインド テーブル内の指定のアイテムとデータの書式を設定または更新します。

|||
|:-----|:-----|
|**ホスト:**|Excel|
|**[要件セット](../../docs/overview/specify-office-hosts-and-api-requirements.md)に指定できるもの**|セットには指定できない|
|**追加されたバージョン**|1.1|

```
bindingObj.setFormatsAsync(cellFormat [,options] , callback);
```


## <a name="parameters"></a>パラメーター



|**名前**|**型**|**説明**|**サポートのメモ**|
|:-----|:-----|:-----|:-----|
| _cellFormat_|**array**|ターゲットとなるセルと、対象セルに適用する書式設定を指定した 1 つ以上の JavaScript オブジェクトが含まれる配列。必須。||
| _options_|**object**|次の[オプションのパラメーター](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)のいずれかを指定します||
| _asyncContext_|**array**、**boolean**、**null**、**number**、**object**、**string**、または **undefined**|変更されずに **AsyncResult** オブジェクトで返される任意の型のユーザー定義項目。||
| _callback_|**object**|コールバックが戻るときに呼び出される関数で、唯一のパラメーターは **AsyncResult** 型です。||

## <a name="callback-value"></a>コールバック値

_callback_ パラメーターに渡した関数が実行されると、[AsyncResult](../../reference/shared/asyncresult.md) オブジェクトを受け取ります。このオブジェクトには、コールバック関数の唯一のパラメーターからアクセスできます。

**goToByIdAsync** メソッドに渡されるコールバック関数で、**AsyncResult** オブジェクトのプロパティを使用して、次の情報を返すことができます。



|**プロパティ**|**使用目的**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|書式を設定するときは、取得するデータやオブジェクトが存在しないため、常に **undefined** を返します。|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|操作の成功または失敗を判断します。|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|操作が失敗した場合、エラーに関する情報を提供する [Error](../../reference/shared/error.md) オブジェクトにアクセスします。|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|ユーザー定義の **object** または値を _asyncContext_ パラメーターとして渡した場合、そのオブジェクトまたは値にアクセスします。|

## <a name="remarks"></a>注釈

 **cellFormat パラメーターを指定する**

_cellFormat_ パラメーターは、幅、高さ、フォント、背景、配置などセルの書式設定値を設定または変更するために使用します。_cellFormat_ パラメーターとして渡す値は、対象となるセル ( `cells:`) とそれらに適用する書式設定 ( `format:`) を指定する 1 つ以上の JavaScript オブジェクトのリストを含む **array** です。

_cellFormat_ 配列内のそれぞれの JavaScript オブジェクトの形式は次のとおりです。

 `{cells:{`_cell_range_`}, format:{`_format_definition_`}}`

`cells:` プロパティは、以下のいずれかの値を使用して書式設定する範囲を指定します。


**cells プロパティでサポートされている範囲**


|**cells の範囲の設定**|**説明**|
|:-----|:-----|
| `{row: i}`|テーブル内の i データ行までの範囲を指定します。|
| `{column: i}`|テーブル内の i データ列までのセルの範囲を指定します。|
| `{row: i, column: j}`|テーブル内の i データ行から j データ列までのセルの範囲を指定します。|
| `Office.Table.All`|列見出し、データ、集計 (もしあれば) を含むテーブル全体を指定します。|
| `Office.Table.Data`|テーブル内のデータのみ (見出しと集計を含まない) を指定します。|
| `Office.Table.Headers`|見出し行のみを指定します。|


プロパティは、Excel の **[セルの書式設定]** ダイアログ ボックス (右クリック `format:` **[セルの書式設定]** または **[ホーム]**  >  **[書式設定]**  >  **[セルの書式設定]**) の設定のサブセットに対応する値を指定します。

`format:` プロパティの値には、JavaScript オブジェクト リテラルの 1 つ以上の _property name_ - _value_ ペアのリストを指定します。_property name_ では設定する書式設定プロパティの名前を指定し、_value_ ではプロパティの値を指定します。フォントの色とサイズの両方など、特定の書式の複数の値を指定できます。3 つの `format:` プロパティ値を指定する例を次に示します。




```
//Set cells: font color to green and size to 15 points.
format: {fontColor : "green", fontSize : 15}
```




```
//Set cells: border to dotted blue.
format: {borderStyle: "dotted", borderColor: "blue"}
```




```
//Set cells: background to red and alignment to centered.
format: {backgroundColor: "red", alignHorizontal: "center"}
```

数値の表示形式を指定するには、「code」文字列の数値の表示形式を  `numberFormat:` プロパティで指定できます。この文字列に指定できる数値の形式は、Excel の [ **セルの書式設定**] ダイアログ ボックスの [ **表示形式**] タブの [ **ユーザー定義**] 分類項目で設定できる形式に対応しています。次の例は、数値を小数点以下 2 桁を含むパーセントとして表示する方法を示しています。




```
format: {numberFormat:"0.00%"}
```

詳細については、「[ユーザー定義の表示形式を作成または削除する](http://office.microsoft.com/en-us/excel-help/create-or-delete-a-custom-number-format-HA102749035.aspx?CTT=1#BM1)」を参照してください。



 **1 つのターゲット指定する**

次の例は、見出し行のフォント色を赤に設定する _cellFormat_ 値を示しています。




```js
Office.select("bindings#myBinding).setFormatsAsync(
    [{cells: Office.Table.Headers, format: {fontColor: "red"}}], 
    function (asyncResult){});
```

 **複数のターゲットを指定する**

**setFormatsAsync** メソッドでは、1 つの関数呼び出しでバインド テーブル内の複数のターゲットを書式設定できます。これを行うには、書式設定するターゲットごとに _cellFormat_ 配列のオブジェクトの一覧を渡します。たとえば、次のコード行では、1 行名のフォントの色を黄色が設定され、3 行目の4 番目のセルに白い罫線と太字テキストが設定されます。




```js
Office.select("bindings#myBinding).setFormatsAsync(
    [{cells: {row: 1}, format: {fontColor: "yellow"}}, 
        {cells: {row: 3, column: 4}, format: {borderColor: "white", fontStyle: "bold"}}], 
    function (asyncResult){});
```

データの書き込み時にテーブルの書式を設定するには、_Document.setSelectedDataAsync_ または _TableBinding.setDataAsync_ メソッドのオプション パラメーター、[tableOptions](http://msdn.microsoft.com/library/4c1e13e9-b61a-47df-836c-3ca9aba4ca1c%28Office.15%29.aspx) と [cellFormat](http://msdn.microsoft.com/library/5b6ecf6f-c57f-4c0d-9605-59daee8fde13%28Office.15%29.aspx) を使用します。

**Document.setSelectedDataAsync** と **TableBinding.setDataAsync** メソッドのオプション パラメーターを使用して書式設定を行えるのは、初回のデータ書き込み時に書式を設定する場合のみです。データの書き込み後に書式設定を変更するには、次のメソッドを使用します。


- フォントの色やスタイルなど、セルの書式を更新するには、 **TableBinding.setFormatsAsync** メソッド (このメソッド) を使用します。
    
- 縞模様 (行) やフィルター ボタンなどのテーブル オプションを更新するには、[TableBinding.setTableOptions](../../reference/shared/binding.tablebinding.settableoptionsasync.md) メソッドを使用します。
    
- 書式設定をクリアするには、[TableBinding.clearFormats](../../reference/shared/binding.tablebinding.clearformatsasync.md) メソッドを使用します。
    
 **Excel Online の追加情報**

_cellFormat_ パラメーターに渡される _書式設定グループ_ の数が 100 を超えることはできません。1 つの書式設定グループは、指定のセル範囲に適用される書式設定のセットから構成されます。たとえば、次の呼び出しは、2 つの書式設定グループを _cellFormat_ に渡します。




```js
Office.select("bindings#myBinding).setFormatsAsync(
    [{cells: {row: 1}, format: {fontColor: "yellow"}}, 
        {cells: {row: 3, column: 4}, format: {borderColor: "white", fontStyle: "bold"}}], 
    function (asyncResult){});

```

詳細および例については、「[Excel 用アドインでテーブルの書式を設定する方法](../../docs/excel/format-tables-in-add-ins-for-excel.md)」を参照してください。


## <a name="support-details"></a>サポートの詳細


次の表で、大文字 Y は、このメソッドは、対応する Office ホスト アプリケーションでサポートされていることを示します。空のセルは、Office ホスト アプリケーションでこのメソッドをサポートしないことを示します。

Office ホスト アプリケーションとサーバーの要件の詳細については、「[Office アドインを実行するための要件](../../docs/overview/requirements-for-running-office-add-ins.md)」をご覧ください。


**サポートされるホスト (プラットフォーム別)**


||**Windows デスクトップ版 Office**||**Office Online (ブラウザー)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|:-----|
|**Excel**|Y||Y|Y|

|||
|:-----|:-----|
|**要件セットに指定できるもの**|セットには指定できない。|
|**最小限のアクセス許可レベル**|[WriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**アドインの種類**|コンテンツ、作業ウィンドウ|
|**ライブラリ**|Office.js|
|**名前空間**|Office|

## <a name="support-history"></a>サポート履歴



****


|**バージョン**|**変更内容**|
|:-----|:-----|
|1.1|Office for iPad の Excel のサポートが追加されました。|
|1.1|導入|
