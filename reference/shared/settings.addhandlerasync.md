

# Settings.addHandlerAsync メソッド
**settingsChanged** イベントのイベント ハンドラーを追加します。

|||
|:-----|:-----|
|**ホスト:**|Excel|
|**[要件セット](../../docs/overview/specify-office-hosts-and-api-requirements.md)に指定できるもの**|設定値|
|**最終変更バージョン**|1.0|

```js
Office.context.document.settings.addHandlerAsync(eventType, handler [, options], callback);
```


## パラメーター



|**名前**|**型**|**説明**|**サポートのメモ**|
|:-----|:-----|:-----|:-----|
| _eventType_|[EventType](../../reference/shared/eventtype-enumeration.md)|追加するイベントの種類を指定します。必須です。||
| _handler_|**object**|追加するイベント ハンドラー関数。必須です。||
| _オプション_|**object**|次の[オプションのパラメーター](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)のいずれかを指定します||
| _asyncContext_|**array**、 **boolean**、 **null**、 **number**、 **object** 、 **string**、または  **undefined**|変更されずに  **AsyncResult** オブジェクトで返される任意の型のユーザー定義項目。||
| _callback_|**object**|コールバックが戻るときに呼び出される関数で、唯一のパラメーターは  **AsyncResult** 型です。||

## コールバック値

_callback_ パラメーターに渡した関数が実行されると、[AsyncResult](../../reference/shared/asyncresult.md) オブジェクトを受け取ります。このオブジェクトには、コールバック関数の唯一のパラメーターからアクセスできます。

**addHandlerAsync** メソッドに渡されるコールバック関数の場合、**AsyncResult** オブジェクトのプロパティを使用して次の情報を返すことができます。



|**プロパティ**|**使用目的**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|イベント ハンドラーを追加する際に取得するデータまたはオブジェクトがないため、常に  **undefined** が返されます。|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|操作の成功または失敗を判断します。|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|操作が失敗した場合、エラーに関する情報を提供する [Error](../../reference/shared/error.md) オブジェクトにアクセスします。|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|ユーザー定義の  **object** または値を _asyncContext_ パラメーターとして渡した場合、そのオブジェクトまたは値にアクセスします。|

## 注釈

指定した  _eventType_ には、イベント ハンドラー関数の名前がそれぞれ異なるのであれば、複数のイベント ハンドラーを追加できます。


 >**重要**  アドインが Excel クライアントで実行されている場合、アドインのコードで **settingsChanged** イベントのハンドラーを登録できますが、このイベントが発生するのは、アドインが Excel Online で開かれているスプレッドシートと共に読み込まれ、_なおかつ_ 複数のユーザーがこのスプレッドシートで作業している (共同編集) 場合のみです。そのため、**settingsChanged** イベントが効率的にサポートされるのは、共同編集シナリオの Excel Online 内のみです。


## 例




```js
function addSelectionChangedEventHandler() {
    Office.context.document.settings.addHandlerAsync(Office.EventType.SettingsChanged, MyHandler);
}

function MyHandler(eventArgs) {
    write('Event raised: ' + eventArgs.type);
    doSomethingWithSettings(eventArgs.settings);
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```




## サポートの詳細


次の表で、大文字 Y は、このメソッドは、対応する Office ホスト アプリケーションでサポートされていることを示します。空のセルは、Office ホスト アプリケーションでこのメソッドをサポートしないことを示します。

Office ホスト アプリケーションとサーバーの要件の詳細については、「[Office アドインを実行するための要件](../../docs/overview/requirements-for-running-office-add-ins.md)」をご覧ください。


||**Windows デスクトップ版 Office**|**Office Online (ブラウザー)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**||Y||

|||
|:-----|:-----|
|**要件セットに指定できるもの**|設定値|
|**最小限のアクセス許可レベル**|[ReadWriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**アプリの種類**|コンテンツ、作業ウィンドウ|
|**ライブラリ**|Office.js|
|**名前空間**|Office|

## サポート履歴




|**変更内容**|**1.1**|
|:-----|:-----|
|1.0|導入|

