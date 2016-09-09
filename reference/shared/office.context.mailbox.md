
# Context.mailbox プロパティ
特に Outlook アドイン向けに API メンバーへのアクセスを提供する  **mailbox** オブジェクトを取得します。

|||
|:-----|:-----|
|**ホスト:**|Outlook|
|**[要件セット](../../docs/overview/specify-office-hosts-and-api-requirements.md)に指定できるもの**|メールボックス|
|**最終変更バージョン**|1.0|

```js
var outlookOm = Office.context.mailbox;
```


## 戻り値


  [mailbox](http://msdn.microsoft.com/library/a3880d3b-8a09-4cf9-9274-f2682cb3b769%28Office.15%29.aspx) オブジェクト。


## 例

次のコード行は、JavaScript API for Office の [item](http://msdn.microsoft.com/library/ad288df1-3ca2-474c-bea4-c51f46e6fc43%28Office.15%29.aspx) オブジェクトにアクセスします。


```js
// Access the Item object.
var item = Office.context.mailbox.item;

```




## サポートの詳細


次の表で、大文字 Y は、このメソッドは、対応する Office ホスト アプリケーションでサポートされていることを示します。空のセルは、Office ホスト アプリケーションでこのメソッドをサポートしないことを示します。

Office ホスト アプリケーションとサーバーの要件の詳細については、「[Office アドインを実行するための要件](../../docs/overview/requirements-for-running-office-add-ins.md)」をご覧ください。


||**Windows デスクトップ版 Office**|**Office Online (ブラウザー)**|**Outlook for Mac**|
|:-----|:-----|:-----|:-----|
|**Outlook**|Y|Y|Y|

|||
|:-----|:-----|
|**要件セットに指定できるもの**|メールボックス|
|**最小限のアクセス許可レベル**|[Restricted](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**アドインの種類**|Outlook|
|**ライブラリ**|Office.js|
|**名前空間**|Office|

## サポート履歴


|**変更内容**|**1.1**|
|:-----|:-----|
|1.0|導入|
