
# <a name="tablebinding.addcolumnsasync-method"></a>TableBinding.addColumnsAsync メソッド
テーブルに列と値を追加します。

|||
|:-----|:-----|
|**ホスト:**|Excel、Word|
|**[要件セット](../../docs/overview/specify-office-hosts-and-api-requirements.md)に指定できるもの**|TableBindings|
|**最終変更バージョン**|1.0|

```
bindingObj.addColumnsAsync(data [, options], callback);
```


## <a name="parameters"></a>パラメーター



|**名前**|**型**|**説明**|**サポートのメモ**|
|:-----|:-----|:-----|:-----|
| _data_|**配列**または [TableData](../../reference/shared/tabledata.md)|テーブルに追加するデータの 1 つ以上の行を含む、配列の配列 ("matrix") または **TableData** オブジェクトです。必須。||
| _options_|**object**|次の[オプションのパラメーター](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)のいずれかを指定します。||
| _asyncContext_|**array**、**boolean**、**null**、**number**、**object**、**string**、または **undefined**|変更されずに **AsyncResult** オブジェクトで返される任意の型のユーザー定義項目。||
| _callback_|**object**|コールバックが戻るときに呼び出される関数で、唯一のパラメーターは **AsyncResult** 型です。||

## <a name="callback-value"></a>コールバック値

_callback_ パラメーターに渡した関数が実行されると、[AsyncResult](../../reference/shared/asyncresult.md) オブジェクトを受け取ります。このオブジェクトには、コールバック関数の唯一のパラメーターからアクセスできます。

**addColumnsAsync** メソッドに渡されたコールバック関数では、**AsyncResult** オブジェクトのパラメーターを利用し、次の情報を返します。



|**プロパティ**|**使用目的**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|取得するオブジェクトまたはデータがないため、常に **undefined** を返します。|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|操作の成功または失敗を判断します。|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|操作が失敗した場合、エラーに関する情報を提供する [Error](../../reference/shared/error.md) オブジェクトにアクセスします。|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|ユーザー定義の **object** または値を _asyncContext_ パラメーターとして渡した場合、そのオブジェクトまたは値にアクセスします。|

## <a name="remarks"></a>注釈

データおよびヘッダーの値を指定する 1 つ以上の列を追加するには、**TableData** オブジェクトを _data_ パラメーターとして渡します。データのみを指定する 1 つ以上の列を追加するには、配列の配列 ("matrix") を _data_ パラメーターとして渡します。

**addColumnAsync** 操作の成功または失敗はアトミックです。つまり、列を追加する操作はその全体が成功する必要があり、1 つでもエラーが発生すると、操作全体がロールバックされます (コールバックに返される **AsyncResult.status** プロパティもエラーを報告します)。


- _data_ 引数として渡す配列内の各行には、更新するテーブルと同数の行が必要です。そうでないと、操作全体が失敗します。
    
- 配列内の各行とセルは、その行またはセルをテーブル内の新しく追加される列に正常に追加する必要があります。何らかの理由によって、行またはセルを設定できなかった場合は、操作全体が失敗します。
    
- **TableData** オブジェクトを data 引数として渡す場合は、ヘッダー行の数が更新するテーブルのヘッダー行の数と同じである必要があります。
    
**Excel Online の追加情報**

このメソッドに対する単一の呼び出しで、**data** パラメーターに渡される _TableData_ オブジェクト内のセルの総数が 20,000 を超えることはできません。


## <a name="example"></a>例

次の例では、`"myTable"` という [id](../../reference/shared/binding.id.md) を持つバインド テーブルに 3 行 1 列を追加します。そのために、**TableData** オブジェクトを **addColumnsAsync** メソッドの _data_ 引数として渡します。この操作を正常に実行するには、更新するテーブルの行数が 3 行である必要があります。


```js
// Add a column to a binding of type table by passing a TableData object.
function addColumns() {
    var myTable = new Office.TableData();
    myTable.headers = [["Cities"]];
    myTable.rows = [["Berlin"], ["Roma"], ["Tokyo"]];

    Office.context.document.bindings.getByIdAsync("myTable", function (result) {
        result.value.addColumnsAsync(myTable);
    });
}
```

次の例では、[id](../../reference/shared/binding.id.md)`myTable` を持つバインド テーブルに 3 行 X 1 列を追加します。そのために、配列の配列 ("matrix") を _addColumnsAsync_ メソッドの **data** 引数として渡します。この操作を正常に実行するには、更新するテーブルの行数が 3 行である必要があります。




```js
// Add a column to a binding of type table by passing an array of arrays.
function addColumns() {
    var myTable = [["Berlin"], ["Roma"], ["Tokyo"]];

    Office.context.document.bindings.getByIdAsync("myTable", function (result) {
        result.value.addColumnsAsync(myTable);
    });
}
```


## <a name="support-details"></a>サポートの詳細


次の表で、大文字 Y は、このメソッドは、対応する Office ホスト アプリケーションでサポートされていることを示します。空のセルは、Office ホスト アプリケーションでこのメソッドをサポートしないことを示します。

Office ホスト アプリケーションとサーバーの要件の詳細については、「[Office アドインを実行するための要件](../../docs/overview/requirements-for-running-office-add-ins.md)」をご覧ください。


**サポートされるホスト (プラットフォーム別)**


||**Windows デスクトップ版 Office**|**Office Online (ブラウザー)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|Y|Y|Y|
|**Word**|Y|Y|Y|

|||
|:-----|:-----|
|**要件セットに指定できるもの**|TableBindings|
|**最小限のアクセス許可レベル**|[ReadWriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**アドインの種類**|コンテンツ、作業ウィンドウ|
|**ライブラリ**|Office.js|
|**名前空間**|Office|

## <a name="support-history"></a>サポート履歴




|**バージョン**|**変更内容**|
|:-----|:-----|
|1.1|Office for iPad における Excel と Word のサポートが追加されました。|
|1.0|導入|
