

# SettingsChangedEventArgs.settings プロパティ
**Settings** イベントが発生した設定を表す **settingsChanged** オブジェクトを取得します。

|||
|:-----|:-----|
|**ホスト:**|Excel|
|**[要件セット](../../docs/overview/specify-office-hosts-and-api-requirements.md)に指定できるもの**|設定値|
|**最終変更バージョン**|1.0|

```js
var mySettings = eventArgsObj.settings;
```


## 戻り値

[settingsChanged](../../reference/shared/document.settings.md) イベントが発生した設定を表す [Settings](../../reference/shared/settings.settingschangedevent.md) オブジェクト。


## サポートの詳細


次の表で、大文字 Y は、このプロパティは、対応する Office ホスト アプリケーションでサポートされていることを示します。空のセルは、Office ホスト アプリケーションでこのプロパティをサポートしないことを示します。

Office ホスト アプリケーションとサーバーの要件の詳細については、「[Office アドインを実行するための要件](../../docs/overview/requirements-for-running-office-add-ins.md)」をご覧ください。



||**Windows デスクトップ版 Office**|**Office Online (ブラウザー)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**||Y||

|||
|:-----|:-----|
|**要件セットに指定できるもの**|設定値|
|**最小限のアクセス許可レベル**|[Restricted](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**アプリの種類**|コンテンツ、作業ウィンドウ|
|**ライブラリ**|Office.js|
|**名前空間**|Office|

## サポート履歴




|**変更内容**|**1.1**|
|:-----|:-----|
|1.0|導入|
