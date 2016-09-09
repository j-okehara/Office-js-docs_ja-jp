
# BindingSelectionChangedEventArgs.binding プロパティ
**SelectionChanged** イベントが発生したバインドを表す**Binding** オブジェクトを取得します。

|||
|:-----|:-----|
|**ホスト:**|Access、Excel、Word|
|**最終変更バージョン**|1.1|

```
var myBinding = eventArgsObj.binding;
```


## 戻り値

[SelectionChanged](../../reference/shared/binding.md) イベントが発生したバインドを表す [Binding](../../reference/shared/binding.bindingselectionchangedevent.md) オブジェクト。


## サポートの詳細


次の表で、大文字 Y は、このプロパティは、対応する Office ホスト アプリケーションでサポートされていることを示します。空のセルは、Office ホスト アプリケーションでこのプロパティをサポートしないことを示します。

Office ホスト アプリケーションとサーバーの要件の詳細については、「[ Office 用アプリを実行するための要件](../../docs/overview/requirements-for-running-office-add-ins.md)」をご覧ください。


**サポートされるホスト (プラットフォーム別)**


||**Windows デスクトップ版 Office**|**Office Online (ブラウザー)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||Y||
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
|1.1|Office for iPad における Excel と Word のサポートが追加されました。|
|1.1|Access 用コンテンツ アドインの **SelectionChanged** イベントで、イベント ハンドラーの追加と削除ができるようになりました。|
|1.0|導入|
