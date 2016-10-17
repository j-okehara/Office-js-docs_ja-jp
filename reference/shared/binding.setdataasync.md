
# <a name="binding.setdataasync-method"></a>Binding.setDataAsync メソッド
指定されたバインド オブジェクトで表されるドキュメントのバインド セクションにデータを書き込みます。

|||
|:-----|:-----|
|**ホスト:**|Access、Excel、Word|
|**[要件セット](../../docs/overview/specify-office-hosts-and-api-requirements.md)に指定できるもの**|MatrixBindings, TableBindings, TextBindings|
|**TableBindings の最終変更**|1.1|

```js
bindingObj.setDataAsync(data [, options] ,callback);
```


## <a name="parameters"></a>パラメーター



|**名前**|**型**|**説明**|**サポートのメモ**|
|:-----|:-----|:-----|:-----|
| _data_|<table><tr><td><b>string</b></td><td>Excel、Excel Online、Word、Word Online のみ</td></tr><tr><td><b>配列</b> (配列の配列 – "matrix")</td><td>Excel および Word のみ</td></tr><tr><td>  <a href="https://msdn.microsoft.com/en-us/library/office/fp161002">  <b>TableData</b></a></td><td>Access、Excel、Word のみ</td></tr><tr><td><b>HTML</b></td><td>Word および Word Online のみ</td></tr><tr><td><b>Office Open XML</b></td><td>Word のみ</td></tr></table>|現在の選択範囲に設定するデータ。必須です。|**変更対象:**1.1。Access 用コンテンツ アドインのサポートには、**TableBinding** 要件セット 1.1 以降が必要です。|
| _options_|**object**|次の[オプションのパラメーター](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)のいずれかを指定します||
| _coercionType_|**[CoercionType](../../reference/shared/coerciontype-enumeration.md)**|設定されるデータを強制的に変換する方法を指定します。 ||
| _columns_|**文字列の配列**| 列名を指定します。|**追加対象:** v1.1。Access 用コンテンツ アドインのテーブル バインドに対してのみ。|
| _rows_|**Office.TableRange.ThisRow**|現在選択されている行にデータを設定するには、定義済みの文字列 "thisRow" を指定します。 |**追加対象:** v1.1。Access 用コンテンツ アドインのテーブル バインドに対してのみ。|
| _startColumn_|**number**|データのサブセットの開始列を指定します (列を指定する値は 0 から始まります)。 |テーブルまたはマトリックスのバインドに対してのみ。このパラメーターを省略すると、先頭列にデータの開始位置が設定されます。|
| _startRow_|**number**|バインディング内のデータのサブセットの開始行を指定します (行を指定する値は 0 から始まります)。 |テーブルまたはマトリックスのバインドに対してのみ。このパラメーターを省略すると、先頭行にデータの開始位置が設定されます。|
| _tableOptions_|**object**|挿入されたテーブルの場合に、[表書式オプション](../../docs/excel/format-tables-in-add-ins-for-excel.md) (ヘッダー行、集計行、縞模様行など) を指定するキー/値ペアのリスト。 |**追加:** v1.1。**サポート:**Excel。|
| _cellFormat_|**object**|挿入されたテーブルの場合に、一定範囲の列、行、またはセルを指定し、その範囲に適用する[セルの書式設定](../../docs/excel/format-tables-in-add-ins-for-excel.md)を指定するキー/値ペアのリスト。|**追加:** v1.1。**サポート:**Excel、Excel Online。|
| _asyncContext_|**array**、**boolean**、**null**、**number**、**object**、**string**、または **undefined**|変更されずに **AsyncResult** オブジェクトで返される任意の型のユーザー定義項目。||
| _callback_|**object**|コールバックが戻るときに呼び出される関数で、唯一のパラメーターは **AsyncResult** 型です。||

## <a name="callback-value"></a>コールバック値

_callback_ パラメーターに渡した関数が実行されると、[AsyncResult](../../reference/shared/asyncresult.md) オブジェクトを受け取ります。このオブジェクトには、コールバック関数の唯一のパラメーターからアクセスできます。

**setDataAsync** メソッドに渡されたコールバック関数で、**AsyncResult** オブジェクトのプロパティを使用して次の情報を戻せます。



|**プロパティ**|**使用目的**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|取得するオブジェクトまたはデータがないため、常に **undefined** を返します。|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|操作の成功または失敗を判断します。|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|操作が失敗した場合、エラーに関する情報を提供する [Error](../../reference/shared/error.md) オブジェクトにアクセスします。|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|ユーザー定義の **object** または値を _asyncContext_ パラメーターとして渡した場合、そのオブジェクトまたは値にアクセスします。|

## <a name="remarks"></a>注釈

_data_ に渡される値には、バインドに書き込まれるデータが含まれます。次の表に示されるように、渡された値の種類により、書き込まれる内容が決まります。



|**_data_ 値**|**書き込まれるデータ**|
|:-----|:-----|
|**string**|**string** に強制的に変換できるプレーンテキストまたはその他の値が書き込まれます。|
|配列の配列 ("matrix")|ヘッダーなしの表形式データが書き込まれます。たとえば、3 行 X 2 列のデータを書き込むには、` [["R1C1", "R1C2"], ["R2C1", "R2C2"], ["R3C1", "R3C2"]]` という配列を渡します。3 行 1 列のデータを書き込むには、`[["R1C1"], ["R2C1"], ["R3C1"]]` という配列を渡します。|
|[TableData](../../reference/shared/tabledata.md) オブジェクト|ヘッダー付きのテーブルが書き込まれます。|
また、バインドにデータを書き込むときに、次のアプリケーション固有の処理が適用されます。

 **Word では**、指定された _data_ は、次の規則に従ってバインドに書き込まれます。



|**_data_ 値**|**書き込まれるデータ**|
|:-----|:-----|
|**string**|指定されたテキストが書き込まれます。|
|配列の配列 ("matrix") または **TableData** オブジェクト|HTML|
|HTML|指定された HTML が書き込まれます。
 >**Important**  If any of the HTML you write is invalid, Word will not raise an error. Word will write as much of the HTML as it can and will omit any invalid data.

||Office Open XML ("Open XML")|指定された XML が書き込まれます。|  **Excel では**、指定された _data_ は、次の規則に従ってバインドに書き込まれます。



|**_data_ 値**|**書き込まれるデータ**|
|:-----|:-----|
|**string**|指定されたテキストが最初にバインドされたセルの値として挿入されます。また、バインドされたセルに追加する有効な数式を指定できます。たとえば、_data_ を `"=SUM(A1:A5)"` に設定すると、指定の範囲内の値が集計されます。ただし、バインドされたセルに数式を設定すると、その後、バインドされたセルからは追加された数式 (または既存の数式) を読み取れなくなります。バインドされたセルで [Binding.getDataAsync](../../reference/shared/binding.getdataasync.md) メソッドを呼び出してそのデータを読み取ると、このメソッドは、(数式の結果である) セルに表示されたデータのみを返すことができます。|
|配列の配列 (「matrix」)、形状が指定されたバインドの形状と完全に一致する場合|行と列のセットが書き込まれます。また、バインドされたセルに追加する有効な数式が含まれる配列の配列を指定できます。例えば、 _data_ を `[["=SUM(A1:A5)","=AVERAGE(A1:A5)"]]` と設定すると、2 つのセルが含まれるバインドに 2 つの数式が追加されます。1 つのバインドされたセルで数式を設定する場合と同じように、 **Binding.getDataAsync** メソッドを使用したバインドからは追加された数式 (または既存の数式) を読み取ることはできません。このメソッドは、バインドされたセルに表示されるデータのみを戻します。|
|**TableData** オブジェクト。テーブルの形状はバインドされたテーブルと一致する。|周囲のセルに含まれるデータが上書きされる場合を除いて、指定された行やヘッダーのセットが書き込まれます。**メモ:** **data** パラメーターに渡す _TableData_ オブジェクトに数式を指定する場合、予期したように結果を取得できないことがあります。これは、Excel の "集計列" が、列内の数式を自動的に複製するためです。バインドされたテーブルに数式が含まれる _data_ を作成する場合にこの問題を回避するには、データを (**TableData** オブジェクトとしてではなく) 配列の配列として指定し、_coercionType_ を **Microsoft.Office.Matrix** または "matrix" として指定してください。|
 **Excel Online の追加情報**


- このメソッドに対する単一の呼び出しで、_data_ パラメーターに渡される値に含まれるセルの総数が 20,000 を超えることはできません。
    
- _cellFormat_ パラメーターに渡される _書式設定グループ_ の数が 100 を超えることはできません。1 つの書式設定グループは、指定のセル範囲に適用される書式設定のセットから構成されます。たとえば、次の呼び出しは、2 つの書式設定グループを _cellFormat_ に渡します。
    
```js
  Office.select("bindings#myBinding").setDataAsync([['Berlin'],['Munich'],['Duisburg']],
    {cellFormat:[{cells: {row: 1}, format: {fontColor: "yellow"}}, 
        {cells: {row: 3, column: 4}, format: {borderColor: "white", fontStyle: "bold"}}]}, 
    function (asyncResult){});

```

上記以外の場合は、エラーが返されます。

**setDataAsync** メソッドは、省略可能な _startRow_ および _startColumn_ パラメーターに有効な範囲が指定されている場合は、テーブルまたはマトリックス バインドのサブセットにデータを書き込みます。


## <a name="example"></a>例




```js
function setBindingData() {
    Office.select("bindings#MyBinding").setDataAsync('Hello World!', function (asyncResult) { });
}
```

省略可能な _coercionType_ パラメーターを指定すると、バインドに書き込むデータの種類を指定できます。たとえば、Word で、テキスト バインドに HTML を書き込む場合は、次の例に示すように、_coercionType_ パラメーターを `"html"` に指定します。この例では、HTML の `<b>` タグを使用して "Hello" を太字に設定します。




```js
function writeHtmlData() {
    Office.select("bindings#myBinding").setDataAsync("<b>Hello</b> World!", {coercionType: "html"}, function (asyncResult) {
        if (asyncResult.status == "failed") {
            write('Error: ' + asyncResult.error.message);
        }
    });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

この例では、**setDataAsync** への呼び出しで、_data_ パラメーターを (1 列 X 3 行を書き込む) 配列の配列として渡し、_coercionType_ パラメーターでデータ構造を `"matrix"` に指定します。




```js
function writeBoundDataMatrix() {
    Office.select("bindings#myBinding").setDataAsync([['Berlin'],['Munich'],['Duisburg']],{ coercionType: "matrix" }, function (asyncResult) {
        if (asyncResult.status == "failed") {
            write('Error: ' + asyncResult.error.message);
        } else {
            write('Bound data: ' + asyncResult.value);
        }
    });
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

この例の `writeBoundDataTable` 関数では、**setDataAsync** への呼び出しで、_data_ パラメーターを (3 列 X 3 行を書き込む) **TableData** オブジェクトとして渡し、_coercionType_ でデータ構造を `"table"` に指定します。 

`updateTableData` 関数では、**setDataAsync** への呼び出しで、再度 _data_ パラメーターを **TableData** オブジェクトとして渡しますが、`writeBoundDataTable` 関数で作成された表の最後の列の値を新規ヘッダーと 1 列 X 3 行として更新します。オプションで、0 から始まる _startColumn_ パラメーターを 2 に指定すると、表の 3 列目の値が置換されます。




```js
function writeBoundDataTable() {
    // Create a TableData object.
    var myTable = new Office.TableData();
    myTable.headers = ['First Name', 'Last Name', 'Grade'];
    myTable.rows = [['Kim', 'Abercrombie', 'A'], ['Junmin','Hao', 'C'],['Toni','Poe','B']];

    // Set myTable in the binding.
    Office.select("bindings#myBinding").setDataAsync(myTable, { coercionType: "table" }, 
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                write('Error: '+ asyncResult.error.message);
        } else {
            write('Bound data: ' + asyncResult.value);
        }
    });
}

// Replace last column with different data.
function updateTableData() {
     var newTable = new Office.TableData();
     newTable.headers = ["Gender"];
     newTable.rows = [["M"],["M"],["F"]];
     Office.select("bindings#myBinding").setDataAsync(newTable, { coercionType: "table", startColumn:2 }, 
         function (asyncResult) {
             if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                 write('Error: '+ asyncResult.error.message);
         } else {
            write('Bound data: ' + asyncResult.value);
         }     
     });   
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
|**要件セットに指定できるもの**|MatrixBindings, TableBindings, TextBindings|
|**最小限のアクセス許可レベル**|[ReadWriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**アドインの種類**|コンテンツ、作業ウィンドウ|
|**ライブラリ**|Office.js|
|**名前空間**|Office|

## <a name="support-history"></a>サポート履歴



|**バージョン**|**変更内容**|
|:-----|:-----|
|1.1|Office for iPad における Excel と Word のサポートが追加されました。|
|1.1|<ul><li>Access 用アドインで、テーブル データの書き込みのサポートが追加されました。</li><li>Excel 用アドインで、オプション パラメーター <span class="parameter" sdata="paramReference">tableOptions</span> および <span class="parameter" sdata="paramReference">cellFormat</span> を使用して<a href="http://msdn.microsoft.com/library/46b05707-b350-41be-b6b8-311799c71a33(Office.15).aspx" target="_blank">テーブルのバインドにデータを書き込む際の書式設定</a>のサポートが追加されました。</li></ul>|
|1.0|導入|

## <a name="see-also"></a>関連項目



#### <a name="other-resources"></a>その他の技術情報


[ドキュメントまたはスプレッドシート内の領域へのバインド](../../docs/develop/bind-to-regions-in-a-document-or-spreadsheet.md)
