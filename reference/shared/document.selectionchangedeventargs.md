
# <a name="documentselectionchangedeventargs-object"></a>DocumentSelectionChangedEventArgs オブジェクト
[SelectionChanged](../../reference/shared/document.selectionchanged.event.md) イベントが発生したドキュメントに関する情報を提供します。

|||
|:-----|:-----|
|**ホスト:**|Excel、PowerPoint、Word|
|**追加されたバージョン**|1.1|

```

```


## <a name="members"></a>メンバー


**プロパティ**


|**名前**|**説明**|
|:-----|:-----|
|[document](../../reference/shared/document.selectionchangedeventargs.document.md)|**SelectionChanged** イベントが発生したドキュメントを表す **Document**オブジェクトを取得します。|
|[type](../../reference/shared/document.selectionchangedeventargs.type.md)|発生したイベントの種類を特定する **EventType** 列挙値を取得します。|

## <a name="support-details"></a>サポートの詳細


次の表で、大文字 Y は、このメソッドは、対応する Office ホスト アプリケーションでサポートされていることを示します。空のセルは、Office ホスト アプリケーションでこのメソッドをサポートしないことを示します。

Office ホスト アプリケーションとサーバーの要件の詳細については、「[Office アドインを実行するための要件](../../docs/overview/requirements-for-running-office-add-ins.md)」をご覧ください。


**サポートされるホスト (プラットフォーム別)**


||**Windows デスクトップ版 Office**|**Office Online (ブラウザー)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|Y|Y|Y|
|**PowerPoint**|Y|Y|Y|
|**Word**|Y||Y|

|||
|:-----|:-----|
|**アドインの種類**|コンテンツ、作業ウィンドウ|
|**ライブラリ**|Office.js|
|**名前空間**|Office|

## <a name="support-history"></a>サポート履歴



****


|**バージョン**|**変更内容**|
|:-----|:-----|
|1.1|Office for iPad で Excel、PowerPoint、および Word のサポートが追加されました。|
|1.0|導入|
