
# DocumentSelectionChangedEventArgs.document プロパティ
**Document** イベントが発生したドキュメントを表す **SelectionChanged** オブジェクトを取得します。

|||
|:-----|:-----|
|**ホスト:**|Excel、Word|
|**で追加**|1.1|




```js
var myDoc = eventArgsObj.document;
```


## 戻り値

[SelectionChanged](../../reference/shared/document.md) イベントが発生したドキュメントを表す [Document](../../reference/shared/document.selectionchanged.event.md) オブジェクト。


## サポートの詳細


次の表で、大文字 Y は、このメソッドは、対応する Office ホスト アプリケーションでサポートされていることを示します。空のセルは、Office ホスト アプリケーションでこのメソッドをサポートしないことを示します。

Office ホスト アプリケーションとサーバーの要件の詳細については、「[Office アドインを実行するための要件](../../docs/overview/requirements-for-running-office-add-ins.md)」をご覧ください。


**サポートされるホスト (プラットフォーム別)**


||**Windows デスクトップ版 Office**|**Office Online (ブラウザー)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|Y|Y|Y|
|**Word**|Y||Y|

|||
|:-----|:-----|
|**最小限のアクセス許可レベル**|[Restricted](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**アプリの種類**|コンテンツ、作業ウィンドウ|
|**ライブラリ**|Office.js|
|**名前空間**|Office|

## サポート履歴



****


|**変更内容**|**1.1**|
|:-----|:-----|
|1.1|Office for iPad で Excel、PowerPoint、および Word のサポートが追加されました。|
|1.0|導入|
