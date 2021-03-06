
# <a name="document.addhandlerasync-method"></a>Document.addHandlerAsync メソッド
**Document** オブジェクト イベントのイベント ハンドラーを追加します。

|||
|:-----|:-----|
|**ホスト:**|Excel、PowerPoint、Project、Word|
|**[要件セット](../../docs/overview/specify-office-hosts-and-api-requirements.md)に指定できるもの**|DocumentEvents|
|**追加されたバージョン**|1.1|


```js
Office.context.document.addHandlerAsync(eventType, handler  [, options], callback);
```


## <a name="parameters"></a>パラメーター



|**名前**|**型**|**説明**|**サポートのメモ**|
|:-----|:-----|:-----|:-----|
| _eventType_|[EventType](../../reference/shared/eventtype-enumeration.md)|**Document** オブジェクト イベントの場合、_eventType_ パラメーターには、**Office.EventType.DocumentSelectionChanged** またはこの列挙型の対応するテキスト値を指定できます。||
| _handler_|**object**|追加するイベント ハンドラー関数で、唯一のパラメーターは [DocumentSelectionChangedEventArgs](../../reference/shared/document.selectionchangedeventargs.md) 型です。必須。||
| _options_|**object**|次の[オプションのパラメーター](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)のいずれかを指定します。||
| _asyncContext_|**array**、**boolean**、**null**、**number**、**object**、**string**、または **undefined**|変更されずに **AsyncResult** オブジェクトで返される任意の型のユーザー定義項目。||
| _callback_|**object**|コールバックが戻るときに呼び出される関数で、唯一のパラメーターは **AsyncResult** 型です。||

## <a name="callback-value"></a>コールバック値

_callback_ パラメーターに渡した関数が実行されると、[AsyncResult](../../reference/shared/asyncresult.md) オブジェクトを受け取ります。このオブジェクトには、コールバック関数の唯一のパラメーターからアクセスできます。

**addHandlerAsync** メソッドに渡されるコールバック関数の場合、**AsyncResult** オブジェクトのプロパティを使用して次の情報を返すことができます。



|**プロパティ**|**使用目的**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|イベント ハンドラーを追加する際に取得するデータまたはオブジェクトがないため、常に **undefined** が返されます。|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|操作の成功または失敗を判断します。|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|操作が失敗した場合、エラーに関する情報を提供する [Error](../../reference/shared/error.md) オブジェクトにアクセスします。|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|ユーザー定義の **object** または値を _asyncContext_ パラメーターとして渡した場合、そのオブジェクトまたは値にアクセスします。|

## <a name="remarks"></a>注釈

指定した  _eventType_ には、イベント ハンドラー関数の名前がそれぞれ異なるのであれば、複数のイベント ハンドラーを追加できます。


## <a name="example"></a>例




```js
function addSelectionChangedEventHandler() {
    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, MyHandler);
}

function MyHandler(eventArgs) {
    write('Event raised: ' + eventArgs.type);
    doSomethingWithDocument(eventArgs.document);
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
|**Excel**|Y|Y|Y|
|**PowerPoint**|Y|Y|Y|
|**Project**|Y|||
|**Word**|Y|Y|Y|

|||
|:-----|:-----|
|**要件セットに指定できるもの**|DocumentEvents|
|**最小限のアクセス許可レベル**|[ReadWriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**アドインの種類**|コンテンツ、作業ウィンドウ|
|**ライブラリ**|Office.js|
|**名前空間**|Office|

## <a name="support-history"></a>サポート履歴



|**バージョン**|**変更内容**|
|:-----|:-----|
|1.1|Office for iPad で Excel、PowerPoint、および Word のサポートが追加されました。|
|1.0|導入|
