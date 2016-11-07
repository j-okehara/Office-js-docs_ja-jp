
# <a name="documentactiveviewchanged-event"></a>Document.ActiveViewChanged イベント
ユーザーがドキュメントの現在のビューを変更したときに発生します。

|||
|:-----|:-----|
|**ホスト:**|PowerPoint|
|**導入バージョン**|1.1|

```
Office.EventType.ActiveViewChanged
```


## <a name="remarks"></a>注釈

ドキュメントの **ActiveViewChanged** イベントのイベント ハンドラーを追加するには、[Document](../../reference/shared/document.addhandlerasync.md) オブジェクトの **addHandlerAsync** メソッドを使用します。このイベント ハンドラーは、[ActiveViewChangedEventArgs](../../reference/shared/document.activeviewchangedeventargs.md) 型の引数を受け取ります。


## <a name="support-details"></a>サポートの詳細


次の表で、大文字 Y は、このメソッドは、対応する Office ホスト アプリケーションでサポートされていることを示します。空のセルは、Office ホスト アプリケーションでこのメソッドをサポートしないことを示します。

Office ホスト アプリケーションとサーバーの要件の詳細については、「[Office アドインを実行するための要件](../../docs/overview/requirements-for-running-office-add-ins.md)」をご覧ください。


**サポートされるホスト (プラットフォーム別)**


||**Windows デスクトップ版 Office**|**Office Online (ブラウザー)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**PowerPoint**|Y||Y|

|||
|:-----|:-----|
|**導入バージョン**|1.1|
|**アドインの種類**|コンテンツ、作業ウィンドウ|
|**ライブラリ**|Office.js|
|**名前空間**|Office|
