
# <a name="binding.bindingselectionchanged-event"></a>Binding.bindingSelectionChanged イベント
バインド内で選択が変更されるときに発生します。

|||
|:-----|:-----|
|**ホスト:**|Access、Excel、Word|
|**[要件セット](../../docs/overview/specify-office-hosts-and-api-requirements.md)に指定できるもの**|BindingEvents|
|**選択内容の最終変更**|1.1|

```
Office.EventType.BindingSelectionChanged
```

## <a name="remarks"></a>注釈

バインドの **BindingSelectionChanged** イベントのイベント ハンドラーを追加するには、[Binding](../../reference/shared/binding.addhandlerasync.md) オブジェクトの **addHandlerAsync** メソッドを使用します。このイベント ハンドラーは、[BindingSelectionChangedEventArgs](../../reference/shared/binding.bindingselectionchangedeventargs.md) 型の引数を受け取ります。


## <a name="example"></a>例




```
function addEventHandlerToBinding() {
 Office.select("bindings#MyBinding").addHandlerAsync(Office.EventType.BindingSelectionChanged, onBindingSelectionChanged);
}

function onBindingSelectionChanged(eventArgs) {
    write(eventArgs.binding.id + " has been selected.");
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## <a name="support-details"></a>サポートの詳細


次の表で、大文字 Y は、このイベントは、対応する Office ホスト アプリケーションでサポートされていることを示します。空のセルは、Office ホスト アプリケーションがこのイベントをサポートしないことを示します。

Office ホスト アプリケーションとサーバーの要件の詳細については、「[Office アドインを実行するための要件](../../docs/overview/requirements-for-running-office-add-ins.md)」をご覧ください。


**サポートされるホスト (プラットフォーム別)**


||**Windows デスクトップ版 Office**|**Office Online (ブラウザー)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||Y||
|**Excel**|Y|Y|Y|
|**Word**|Y|Y|Y|

|||
|:-----|:-----|
|**要件セットに指定できるもの**|BindingEvents|
|**アドインの種類**|コンテンツ、作業ウィンドウ|
|**ライブラリ**|Office.js|
|**名前空間**|Office|

## <a name="support-history"></a>サポート履歴





****


|**バージョン**|**変更内容**|
|:-----|:-----|
|1.1|Office for iPad における Excel と Word のサポートが追加されました。|
|1.1|
            Access 用アプリにこのイベントのサポートが追加されました。|
|1.0|導入|
