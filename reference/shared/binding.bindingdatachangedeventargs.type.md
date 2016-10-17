
# <a name="bindingdatachangedeventargs.type-property"></a>BindingDataChangedEventArgs.type プロパティ
発生したイベントの種類を識別する **EventType** 列挙値を取得します。

|||
|:-----|:-----|
|**ホスト:**|Access、Excel、Word|
|**BindingEvents の最終変更**|1.1|

```js
var myEvent = eventArgsObj.type;
```


## <a name="return-value"></a>戻り値

発生したイベントの [EventType](../../reference/shared/eventtype-enumeration.md)。


## <a name="support-details"></a>サポートの詳細


次の表で、大文字 Y は、このプロパティは、対応する Office ホスト アプリケーションでサポートされていることを示します。空のセルは、Office ホスト アプリケーションでこのプロパティをサポートしないことを示します。

Office ホスト アプリケーションとサーバーの要件の詳細については、「[Office アドインを実行するための要件](../../docs/overview/requirements-for-running-office-add-ins.md)」をご覧ください。


**サポートされるホスト (プラットフォーム別)**


||**Windows デスクトップ版 Office**|**Office Online (ブラウザー)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||Y||
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
|1.1|Office for iPad における Excel と Word のサポートが追加されました。|
|1.1|Access 用のアドインのサポートが追加されました。|
|1.0|導入|
