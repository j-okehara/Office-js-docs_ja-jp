
# Context.commerceAllowed プロパティ
外部の支払システムにリンクできるプラットフォーム上で アドイン が実行されているかどうかを取得します。

|||
|:-----|:-----|
|**ホスト:**|Excel、Word|
|**最終変更バージョン**|1.1|

```
var allowCommerce = Office.context.commerceAllowed;
```


## 戻り値

開発者がセルを表示するか、プラットフォーム上のアドインの UI を更新できる場合は **True** を返します。それ以外の場合は、**False** を返します。


## 注釈

iOS アプリ ストアでは、追加の支払いシステムへのリンクを提供するアドインを含むアプリをサポートしません。ただし、Windows デスクトップで実行されている Office アドイン、またはブラウザーの Office Online に対しては、それらのリンクが許可されます。アドインの UI が iOS 以外のプラットフォーム上の外部支払いシステムへのリンクを提供するようにする場合は、 **commerceAllowed** プロパティを使って、そのリンクが表示されるタイミングを制御できます。


## サポートの詳細


次の表で、大文字 Y は、このメソッドは、対応する Office ホスト アプリケーションでサポートされていることを示します。空のセルは、Office ホスト アプリケーションでこのメソッドをサポートしないことを示します。

Office ホスト アプリケーションとサーバーの要件の詳細については、「[Office アドインを実行するための要件](../../docs/overview/requirements-for-running-office-add-ins.md)」をご覧ください。


||**Office for iPad**|
|:-----|:-----|
|**Excel**|Y|
|**PowerPoint**||
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
