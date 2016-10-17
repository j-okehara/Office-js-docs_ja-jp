
# <a name="binding.bindingdatachanged-event"></a>Binding.bindingDataChanged イベント
バインド内でデータが変更されるときに発生します。

|||
|:-----|:-----|
|**ホスト:**|Access、Excel、Word|
|**BindingEvents の最終変更**|1.1|

```js
Office.EventType.BindingDataChanged
```


## <a name="remarks"></a>注釈

バインドの **BindingDataChanged** イベントのイベント ハンドラーを追加するには、[Binding](../../reference/shared/binding.addhandlerasync.md) オブジェクトの **addHandlerAsync** メソッドを使用します。このイベント ハンドラーは、[BindingDataChangedEventArgs](../../reference/shared/binding.bindingdatachangedeventargs.md) 型の引数を受け取ります。


## <a name="example"></a>例




```js
function addEventHandlerToBinding() {
    Office.select("bindings#MyBinding").addHandlerAsync(Office.EventType.BindingDataChanged, onBindingDataChanged);
}

function onBindingDataChanged(eventArgs) {
    write("Data has changed in binding: " + eventArgs.binding.id);
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
|**Word**|Y||Y|

|||
|:-----|:-----|
|**要件セットに指定できるもの**|BindingEvents|
|**アドインの種類**|コンテンツ、作業ウィンドウ|
|**ライブラリ**|Office.js|
|**名前空間**|Office|

## <a name="support-history"></a>サポート履歴

|**バージョン**|**変更内容**|
|:-----|:-----|
|1.1|Office for iPad における Excel と Word のサポートが追加されました。|
|1.1|
            Access 用アプリにこのイベントのサポートが追加されました。|
|1.0|導入|
