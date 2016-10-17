
# <a name="binding.getdataasync-method"></a>Binding.getDataAsync メソッド
バインド内に含まれるデータを返します。

|||
|:-----|:-----|
|**ホスト:**|Access、Excel、Word|
|**[要件セット](../../docs/overview/specify-office-hosts-and-api-requirements.md)に指定できるもの**|MatrixBindings, TableBindings, TextBindings|
|**TableBindings の最終変更**|1.1|

```
bindingObj.getDataAsync([, options] , callback );
```


## <a name="parameters"></a>パラメーター



|**名前**|**型**|**説明**|**サポートのメモ**|
|:-----|:-----|:-----|:-----|
| _options_|**object**|次の[オプションのパラメーター](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)のいずれかを指定します||
| _coercionType_|**[CoercionType](../../reference/shared/coerciontype-enumeration.md)**|設定されるデータを強制的に変換する方法を指定します。 ||
| _valueFormat_|[ValueFormat](../../reference/shared/valueformat-enumeration.md)|返される値 (数字、日付など) に書式を適用するかどうかを指定します。||
| _filterType_|[FilterType](../../reference/shared/filtertype-enumeration.md)|データを取得するときにフィルターを適用する必要があるかどうかを指定します。||
| _rows_|**Office.TableRange.ThisRow**| 現在選択されている行でデータを取得するには、定義済みの文字列 "thisRow" を指定します。|Access 用コンテンツ アドインのテーブル バインドに対してのみ。|
| _startRow_|**number**|テーブルまたはマトリックス バインドの場合、データのサブセットの開始行を指定します (行を指定する値は 0 から始まります)。 ||
| _startColumn_|**number**|テーブルまたはマトリックス バインドの場合、データのサブセットの開始列を指定します (列を指定する値は 0 から始まります)。 ||
| _rowCount_|**number**|テーブルまたはマトリックス バインドの場合、_startRow_ からの行のオフセット数を指定します。 ||
| _columnCount_|**number**|テーブルまたはマトリックス バインドの場合、_startColumn_ からの列のオフセット数を指定します。||
| _asyncContext_|**array**、**boolean**、**null**、**number**、**object**、**string**、または **undefined**|変更されずに **AsyncResult** オブジェクトで返される任意の型のユーザー定義項目。||
| _callback_|**object**|コールバックが戻るときに呼び出される関数で、唯一のパラメーターは **AsyncResult** 型です。||

## <a name="callback-value"></a>コールバック値

_callback_ パラメーターに渡した関数が実行されると、[AsyncResult](../../reference/shared/asyncresult.md) オブジェクトを受け取ります。このオブジェクトには、コールバック関数の唯一のパラメーターからアクセスできます。

**Binding.getDataAsync** メソッドに渡されるコールバック関数では、**AsyncResult** オブジェクトのプロパティを使用して次の情報を返すことができます。



|**プロパティ**|**使用目的**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|指定したバインドの値にアクセスします。_coercionType_ パラメーターを指定した場合 (呼び出しが成功すると)、データは [CoercionType](../../reference/shared/coerciontype-enumeration.md) 列挙体のトピックで説明する形式で返されます。|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|操作の成功または失敗を判断します。|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|操作が失敗した場合、エラーに関する情報を提供する [Error](../../reference/shared/error.md) オブジェクトにアクセスします。|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|ユーザー定義の **object** または値を _asyncContext_ パラメーターとして渡した場合、そのオブジェクトまたは値にアクセスします。|

## <a name="remarks"></a>注釈

省略可能なパラメーターを省略すると、次の既定値が使用されます (データの種類と形式に該当する場合)。



|**パラメーター**|**既定**|
|:-----|:-----|
| _coercionType_|元のバインドの種類 (強制的に変換される前の種類)。|
| _valueFormat_|書式なしデータ。|
| _filterType_|すべての値 (フィルターなし)。|
| _startRow_|先頭行。|
| _startColumn_|先頭列。|
| _rowCount_|すべての行。|
| _columnCount_|すべての列。|
[getDataAsync](../../reference/shared/binding.matrixbinding.md) メソッドを [MatrixBinding](../../reference/shared/binding.tablebinding.md) または **TableBinding** から呼び出す場合に、省略可能な _startRow_、_startColumn_、_rowCount_、_columnCount_ パラメーターが指定されていると (これらのパラメーターに連続する有効な範囲が指定されていると)、このメソッドはバインド値のサブセットを返します。


## <a name="example"></a>例




```
function showBindingData() {
    Office.select("bindings#MyBinding").getDataAsync(function (asyncResult) {
        write(asyncResult.value)
    });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```



_Binding.getDataAsync_ メソッドで `"table"` を使用する場合と `"matrix"`**coercionType** を使用する場合とでは、以下の 2 つの例からわかるように、ヘッダー行でのデータのフォーマットに関して動作に重要な違いが存在します。これらのコード例は、[Binding.SelectionChanged](../../reference/shared/binding.bindingselectionchangedevent.md) イベント用のイベント ハンドラー関数を示しています。

`"table"` _coercionType_ を指定した場合、[TableData.rows](../../reference/shared/tabledata.rows.md) プロパティ (以下のコード例中の `result.value.rows`) は、テーブルの本体行のみが含まれる配列を返します。このため、その 0 番目の行は、テーブル内で最初のヘッダー行以外の行となります。




```js
function selectionChanged(evtArgs) { 
    Office.select("bindings#TableTranslate").getDataAsync({ coercionType: 'table', startRow: evtArgs.startRow, startCol: 0, rowCount: 1, columnCount: 1 },  
        function (result) { 
            if (result.status == 'succeeded') { 
                write("Image to find: " + result.value.rows[0][0]); 
            } 
            else 
                write(result.error.message); 
    }); 
}     
// Function that writes to a div with id='message' on the page. 
function write(message){ 
    document.getElementById('message').innerText += message; 
}
```

ただし、 `"matrix"`_coercionType_ を指定した場合、以下のコード例の `result.value` は 0 番目の行のテーブル ヘッダーが含まれる配列を返します。テーブル ヘッダーに複数の行が含まれる場合は、これらのヘッダーがすべて個別の行として `result.value` マトリックスに取り込まれてから、テーブル本体の行が取り込まれます。




```js
function selectionChanged(evtArgs) { 
    Office.select("bindings#TableTranslate").getDataAsync({ coercionType: 'matrix', startRow: evtArgs.startRow, startCol: 0, rowCount: 1, columnCount: 1 },  
        function (result) { 
            if (result.status == 'succeeded') { 
                write("Image to find: " + result.value[1][0]); 
            } 
            else 
                write(result.error.message); 
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


**サポートされるホスト (プラットフォーム別)**


||**Windows デスクトップ版 Office**|**Office Online (ブラウザー)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||Y||
|**Excel**|Y|Y|Y|
|**Word**|Y|Y|Y|

|||
|:-----|:-----|
|**要件セットに指定できるもの**|MatrixBindings, TableBindings, TextBindings|
|**最小限のアクセス許可レベル**|[ReadDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**アドインの種類**|コンテンツ、作業ウィンドウ|
|**ライブラリ**|Office.js|
|**名前空間**|Office|

## <a name="support-history"></a>サポート履歴



****


|**バージョン**|**変更内容**|
|:-----|:-----|
|1.1|Office for iPad における Excel と Word のサポートが追加されました。|
|1.1|Access 用アドインのテーブル バインドのサポートが追加されました。|
|1.0|導入|

## <a name="see-also"></a>関連項目



#### <a name="other-resources"></a>その他の技術情報


[ドキュメントまたはスプレッドシート内の領域へのバインド](../../docs/develop/bind-to-regions-in-a-document-or-spreadsheet.md)
