

# <a name="office.select-method"></a>Office.select メソッド
渡されるセレクター文字列に基づくバインドを返す promise を作成します。

|||
|:-----|:-----|
|**ホスト:**|Access、Excel、Word|
|**[要件セット](../../docs/overview/specify-office-hosts-and-api-requirements.md)に指定できるもの**|MatrixBindings, PartialTableBindings, TableBindings, TextBindings|
|**最終変更バージョン**|1.1|

```js
Office.select(str, onError);
```


## <a name="parameters"></a>パラメーター


_str_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;型: **string**<br/>
&nbsp;&nbsp;&nbsp;&nbsp;promise を解析および作成するセレクター文字列。

_onError_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;型: **function**<br/>
&nbsp;&nbsp;&nbsp;&nbsp;コールバックが戻るときに呼び出される関数で、唯一のパラメーターは **AsyncResult** 型です。省略可能。
    

## <a name="callback-value"></a>コールバック値

_onError_ パラメーターに渡した関数が実行されると、[AsyncResult](../../reference/shared/asyncresult.md) オブジェクトを受け取ります。このオブジェクトには、コールバック関数の唯一のパラメーターからアクセスできます。操作が失敗した場合は、[AsyncResult.error](../../reference/shared/asyncresult.error.md) プロパティを使用して、エラーに関する情報を提供する [Error](../../reference/shared/error.md) オブジェクトにアクセスします。


## <a name="remarks"></a>注釈

**Office.select** メソッドは、任意の非同期メソッドが実行されたときに指定されたバインドを返すことを試行する、[Binding](../../reference/shared/binding.md) オブジェクトの promise へのアクセスを提供します。

サポートされている形式: "bindings# _bindingId_"。[id](../../reference/shared/binding.id.md) が `bindingId`.であるバインドの **Binding** オブジェクトを返します。詳細については、「[Office アドインにおける非同期プログラミング](../../docs/develop/asynchronous-programming-in-office-add-ins.md#asynchronous-programming-using-the-promises-pattern-to-access-data-in-bindings)」および「[ドキュメントまたはスプレッドシート内の領域へのバインド](../../docs/develop/bind-to-regions-in-a-document-or-spreadsheet.md)」を参照してください。


 >**注**:**select** メソッドの promise が正常に **Binding** オブジェクトを返す場合、このオブジェクトは [Binding](../../reference/shared/binding.md) オブジェクトの [getDataAsync](../../reference/shared/binding.getdataasync.md)、[setDataAsync](../../reference/shared/binding.setdataasync.md)、[addHandlerAsync](../../reference/shared/binding.addhandlerasync.md)、および [removeHandlerAsync](../../reference/shared/binding.removehandlerasync.md) の 4 つのメソッドのみを公開します。promise が **Binding** オブジェクトを返すことができない場合は、_onError_ コールバックを使用して [asyncResult.error](../../reference/shared/asyncresult.error.md) オブジェクトにアクセスし、詳細情報を取得できます。**select** メソッドによって返される **Binding** オブジェクトの promise によって公開される 4 つのメソッド以外の **Binding** オブジェクトのメンバーを呼び出す必要がある場合は、代わりに [getByIdAsync](../../reference/shared/bindings.getbyidasync.md) メソッドを使用します。[Document.bindings](../../reference/shared/document.bindings.md) プロパティと [Bindings.getByIdAsync](../../reference/shared/bindings.getbyidasync.md) メソッドを使用して **Binding** オブジェクトを取得します。


## <a name="example"></a>例

次のコード例では、**select** メソッドを使用して、"`cities`" という **id** を持つバインドを **Bindings** コレクションから取得します。その後、[addHandlerAsync](../../reference/shared/binding.addhandlerasync.md) メソッドを呼び出して、そのバインドの [dataChanged](../../reference/shared/binding.bindingdatachangedevent.md) イベントのイベント ハンドラーを追加します。


```js
function addBindingDataChangedEventHandler() {
    Office.select("bindings#cities", function onError(){}).addHandlerAsync(Office.EventType.BindingDataChanged,
    function (eventArgs) {
        doSomethingWithBinding(eventArgs.binding);
    });
}
```




## <a name="support-details"></a>サポートの詳細


次の表で、大文字 Y は、このメソッドは、対応する Office ホスト アプリケーションでサポートされていることを示します。空のセルは、Office ホスト アプリケーションでこのメソッドをサポートしないことを示します。

Office ホスト アプリケーションとサーバーの要件の詳細については、「[Office アドインを実行するための要件](../../docs/overview/requirements-for-running-office-add-ins.md)」をご覧ください。



||**Windows デスクトップ版 Office**|**Office Online (ブラウザー)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||Y||
|**Excel**|Y|Y|Y|
|**Word**|Y||Y|

|||
|:-----|:-----|
|**要件セットに指定できるもの**|MatrixBindings, PartialTableBindings, TableBindings, TextBindings|
|**最小限のアクセス許可レベル**|[ReadDocument (Open Office XML 用 ReadAllDocument)](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**アドインの種類**|コンテンツ、作業ウィンドウ|
|**ライブラリ**|Office.js|
|**名前空間**|Office|

## <a name="support-history"></a>サポート履歴



|**バージョン**|**変更内容**|
|:-----|:-----|
|1.1|Office for iPad における Excel と Word のサポートが追加されました。|
|1.1|Access 用コンテンツ アドインで作成されたテーブル バインドを返す **select** メソッドの使用が追加されました。|
|1.0|導入|
