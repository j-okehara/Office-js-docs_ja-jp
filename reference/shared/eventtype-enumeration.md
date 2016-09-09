
# EventType 列挙型
発生したイベントの種類を指定します。" **イベント名** " _EventArgs_ オブジェクトの **type** プロパティから返されます。

|||
|:-----|:-----|
|**ホスト:**|Access、Excel、PowerPoint、Project、Word|
|**選択内容の最終変更**|1.1|

```js
Office.EventType
```


## メンバー


**値**


|列挙体|値|説明|
|:-----|:-----|:-----|
|Office.EventType.ActiveViewChanged|"documentActiveViewChanged"|[Document.ActiveViewChanged](../../reference/shared/document.activeviewchanged.md) イベントが発生しました。|
|Office.EventType.DocumentSelectionChanged|"documentSelectionChanged"|[Document.SelectionChanged](../../reference/shared/document.selectionchanged.event.md) イベントが発生しました。|
|Office.EventType.BindingSelectionChanged|"bindingSelectionChanged"|[Binding.BindingSelectionChanged](../../reference/shared/binding.bindingselectionchangedevent.md) イベントが発生しました。|
|Office.EventType.BindingDataChanged|"bindingDataChanged"|[Binding.BindingDataChanged](../../reference/shared/binding.bindingdatachangedevent.md) イベントが発生しました。|
|Office.EventType.DataNodeDeleted|"nodeDeleted"|[CustomXmlPart.dataNodeDeleted](../../reference/shared/customxmlpart.datanodedeleted.event.md) イベントが発生しました。|
|Office.EventType.DataNodeInserted|"nodeInserted"|[CustomXmlPart.dataNodeInserted](../../reference/shared/customxmlpart.datanodeinserted.event.md) イベントが発生しました。|
|Office.EventType.DataNodeReplaced|"nodeReplaced"|[CustomXmlPart.dataNodeReplaced](../../reference/shared/customxmlpart.datanodereplaced.event.md) イベントが発生しました。|
|Office.EventType.SettingsChanged|"settingsChanged"|[Settings.settingsChanged](../../reference/shared/settings.settingschangedevent.md) イベントが発生しました。|

## 解説


 >**メモ**  Project 用アドインは、イベントの種類 **Office.EventType.ResourceSelectionChanged**、**Office.EventType.TaskSelectionChanged**、および **Office.EventType.ViewSelectionChanged** をサポートしています。


## サポートの詳細


次の表で、大文字 Y は、この列挙は、対応する Office ホスト アプリケーションでサポートされていることを示します。空のセルは、Office ホスト アプリケーションがこの列挙をサポートしないことを示します。

Office ホスト アプリケーションとサーバーの要件の詳細については、「[Office アドインを実行するための要件](../../docs/overview/requirements-for-running-office-add-ins.md)」をご覧ください。


**サポートされるホスト (プラットフォーム別)**


||**Windows デスクトップ版 Office**|**Office Online (ブラウザー)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|Y|Y|Y|
|**PowerPoint**|Y|Y||
|**Project**|Y|||
|**Word**|Y||Y|

|||
|:-----|:-----|
|**アプリの種類**|コンテンツ、作業ウィンドウ|
|**ライブラリ**|Office.js|
|**名前空間**|Office|

## サポート履歴



|**変更内容**|**1.1**|
|:-----|:-----|
|1.1| 列挙 Office.EventType.ActiveViewChanged が新しい **Document.ActiveViewChanged** イベントに追加されました。|
|1.0|導入|
