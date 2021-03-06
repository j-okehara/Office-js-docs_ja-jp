
# <a name="documentselectionchangedeventargs.document-property"></a>DocumentSelectionChangedEventArgs.document プロパティ
**SelectionChanged** イベントが発生したドキュメントを表す **Document**オブジェクトを取得します。

|||
|:-----|:-----|
|**ホスト:**|Excel、Word|
|**追加されたバージョン**|1.1|




```js
var myDoc = eventArgsObj.document;
```


## <a name="return-value"></a>戻り値

[SelectionChanged](../../reference/shared/document.md) イベントが発生したドキュメントを表す [Document](../../reference/shared/document.selectionchanged.event.md) オブジェクト。


## <a name="support-details"></a>サポートの詳細


次の表で、大文字 Y は、このメソッドは、対応する Office ホスト アプリケーションでサポートされていることを示します。空のセルは、Office ホスト アプリケーションでこのメソッドをサポートしないことを示します。

Office ホスト アプリケーションとサーバーの要件の詳細については、「[Office アドインを実行するための要件](../../docs/overview/requirements-for-running-office-add-ins.md)」をご覧ください。


**サポートされるホスト (プラットフォーム別)**


||**Windows デスクトップ版 Office**|**Office Online (ブラウザー)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|Y|Y|Y|
|**Word**|Y||Y|

|||
|:-----|:-----|
|**最小限のアクセス許可レベル**|[制限あり](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**アドインの種類**|コンテンツ、作業ウィンドウ|
|**ライブラリ**|Office.js|
|**名前空間**|Office|

## <a name="support-history"></a>サポート履歴



****


|**バージョン**|**変更内容**|
|:-----|:-----|
|1.1|Office for iPad で Excel、PowerPoint、および Word のサポートが追加されました。|
|1.0|導入|
