
# BindingDataChangedEventArgs オブジェクト
[DataChanged](../../reference/shared/binding.bindingdatachangedevent.md) イベントが発生したバインドに関する情報を提供します。

|||
|:-----|:-----|
|**ホスト:**|Access、Excel、Word|
|**BindingEvents の最終変更**|1.1|

```js
Office.EventType.BindingDataChanged
```


## メンバー


**プロパティ**


|**名前**|**説明**|
|:-----|:-----|
|[binding](../../reference/shared/binding.bindingdatachangedeventargs.binding.md)|[DataChanged](../../reference/shared/binding.md) イベントが発生したバインドを表す**Binding** オブジェクトを取得します。|
|[type](../../reference/shared/binding.bindingdatachangedeventargs.type.md)|発生したイベントの種類を識別する [EventType](../../reference/shared/eventtype-enumeration.md) 列挙値を取得します。|

## サポートの詳細


次の表で、大文字 Y は、このオブジェクトは、対応する Office ホスト アプリケーションでサポートされていることを示します。空のセルは、Office ホスト アプリケーションでこのオブジェクトをサポートしないことを示します。

Office ホスト アプリケーションとサーバーの要件の詳細については、「[Office アドインを実行するための要件](../../docs/overview/requirements-for-running-office-add-ins.md)」をご覧ください。


**サポートされるホスト (プラットフォーム別)**


||**Windows デスクトップ版 Office**|**Office Online (ブラウザー)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||Y||
|**Excel**|Y|Y|Y|
|**Word**|Y||Y|

|||
|:-----|:-----|
|**アプリの種類**|コンテンツ、作業ウィンドウ|
|**ライブラリ**|Office.js|
|**名前空間**|Office|

## サポート履歴




|**変更内容**|**1.1**|
|:-----|:-----|
|1.1|Office for iPad における Excel と Word のサポートが追加されました。|
|1.1|
            Access 用アプリにこのイベントのサポートが追加されました。|
|1.0|導入|
