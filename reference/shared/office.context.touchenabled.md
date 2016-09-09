
# Context.touchEnabled プロパティ
タッチ対応 Office ホスト アプリケーションで アドイン が実行されているかどうかを取得します。

|||
|:-----|:-----|
|**ホスト:**|Excel、Word|
|**最終変更バージョン**|1.1|

```
var isTouchEnabled = Office.context.touchEnabled;
```


## 戻り値

アドインが iPad などのタッチ デバイスで実行されている場合は、**True** を返します。それ以外の場合は、**False** を返します。


## 注釈

**touchEnabled** プロパティを使って、アドインがいつタッチ デバイスで実行されているかを識別し、必要であれば、アドインの UI のコントロールの種類や、要素のサイズと間隔を調整して、タッチ操作に対応します。


## サポートの詳細


次の表で、大文字 Y は、このメソッドは、対応する Office ホスト アプリケーションでサポートされていることを示します。空のセルは、Office ホスト アプリケーションでこのメソッドをサポートしないことを示します。

Office ホスト アプリケーションとサーバーの要件の詳細については、「[Office アドインを実行するための要件](../../docs/overview/requirements-for-running-office-add-ins.md)」をご覧ください。

||**Office for iPad**|
|:-----|:-----|
|**Excel**|Y|
|**PowerPoint**|Y|
|**Word**|Y|

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
|1.1|導入。|
