
# <a name="tablebinding.addrowsasync-method"></a>TableBinding.addRowsAsync メソッド
テーブルに行と値を追加します。

|||
|:-----|:-----|
|**ホスト:**|Access、Excel、Word|
|**[要件セット](../../docs/overview/specify-office-hosts-and-api-requirements.md)に指定できるもの**|TableBindings|
|**最終変更バージョン**|1.1|

```js
bindingObj.addRowsAsync(rows, [,options], callback);
```


## <a name="parameters"></a>パラメーター

_rows_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;型:**array**

&nbsp;&nbsp;&nbsp;&nbsp;テーブルに追加する 1 行以上のデータが含まれる配列の配列。必須。
    
_options_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;型: **object**

&nbsp;&nbsp;&nbsp;&nbsp;次の[オプションのパラメーター](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)を指定します。
    
&nbsp;&nbsp;&nbsp;&nbsp;_asyncContext_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;型: **array、boolean、null、number、object、string、undefined**<br/><br/>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;変更されずに [AsyncResult](../../reference/shared/asyncresult.md) オブジェクトで返される任意の型のユーザー定義項目。省略可能。<br/><br/>

_callback_<br />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;型: **object**
    
&nbsp;&nbsp;&nbsp;&nbsp;コールバックが戻るときに呼び出される関数。唯一のパラメーターは [AsyncResult](../../reference/shared/asyncresult.md) 型です。省略可能。



|**名前**|**型**|**説明**|**サポートのメモ**|
|:-----|:-----|:-----|:-----|
| _rows_|**array**|テーブルに追加する 1 行以上のデータが含まれる配列の配列。必須。||
| _options_|**object**|次の[オプションのパラメーター](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)のいずれかを指定します。||
| _asyncContext_|**array**、**boolean**、**null**、**number**、**object**、**string**、または **undefined**|変更されずに **AsyncResult** オブジェクトで返される任意の型のユーザー定義項目。||
| _callback_|**object**|コールバックが戻るときに呼び出される関数で、唯一のパラメーターは **AsyncResult** 型です。||

## <a name="callback-value"></a>コールバック値

_callback_ パラメーターに渡した関数が実行されると、[AsyncResult](../../reference/shared/asyncresult.md) オブジェクトを受け取ります。このオブジェクトには、コールバック関数の唯一のパラメーターからアクセスできます。

**addRowsAsync** メソッドに渡されるコールバック関数では、**AsyncResult** オブジェクトのプロパティを使用して次の情報を返すことができます。



|**プロパティ**|**使用目的**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|取得するオブジェクトまたはデータがないため、常に **undefined** を返します。|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|操作の成功または失敗を判断します。|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|操作が失敗した場合、エラーに関する情報を提供する [Error](../../reference/shared/error.md) オブジェクトにアクセスします。|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|ユーザー定義の **object** または値を _asyncContext_ パラメーターとして渡した場合、そのオブジェクトまたは値にアクセスします。|

## <a name="remarks"></a>注釈

The success or failure of an  **addRowsAsync** operation is atomic. That is, the entire add rows operation must succeed, or it will be completely rolled back (and the **AsyncResult.status** property returned to the callback will report failure):


- _data_ 引数として渡す配列内の各行には、更新するテーブルと同数の列が必要です。そうでないと、操作全体が失敗します。
    
- 配列内の各行とセルは、その行とセルをテーブル内の新しく追加される行に正常に追加する必要があります。何かの理由によって、行またはセルを設定できなかった場合は、操作全体が失敗します。
    
 **Excel Online の追加情報**

このメソッドに対する単一の呼び出しで、_rows_ パラメーターに渡される値に含まれるセルの総数が 20,000 を超えることはできません。


## <a name="example"></a>例




```js
function addRowsToTable() {
    Office.context.document.bindings.getByIdAsync("myBinding", function (asyncResult) {
        var binding = asyncResult.value;
        binding.addRowsAsync([["6", "k"], ["7", "j"]]);
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
|**要件セットに指定できるもの**|TableBindings|
|**最小限のアクセス許可レベル**|[ReadWriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**アドインの種類**|コンテンツ、作業ウィンドウ|
|**ライブラリ**|Office.js|
|**名前空間**|Office|

## <a name="support-history"></a>サポート履歴




|**バージョン**|**変更内容**|
|:-----|:-----|
|1.1|Office for iPad における Excel と Word のサポートが追加されました。|
|1.1|Access 用アドインでのテーブル データの書き込みのサポートが追加されました。|
|1.0|導入|
