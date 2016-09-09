

# event.source.id
この関数の呼び出しをトリガーしたコントロールの ID を取得します。

****

|||
|:-----|:-----|
|**ホスト:**Outlook|**アドインの種類:** Outlook|
|**[要件セット](../../docs/overview/specify-office-hosts-and-api-requirements.md)に指定できるもの**|メールボックス|
|**メールボックスの最終変更**|1.3|
|**適用可能な Outlook のモード**|読み取りおよび作成|



```js
event.source.id;
```


## 戻り値

この関数の呼び出しをトリガーしたコントロールの ID。この ID はマニフェストに基づきます。


## サポートの詳細


以下の表の大文字 Y は、対象プロパティが対応する Outlook ホスト アプリケーションでサポートされていることを示します。セルが空の場合、Outlook ホスト アプリケーションは対象プロパティをサポートしません。

Office ホスト アプリケーションとサーバーの要件の詳細については、「[Office アドインを実行するための要件](../../docs/overview/requirements-for-running-office-add-ins.md)」をご覧ください。

 **重要:** 現在アドイン コマンドおよびアドイン コマンドに関連付けられている API は、Windows デスクトップ上の [Office 2016 Preview](https://products.office.com/en-us/office-2016-preview) の Outlook でしか動作しません。


**サポートされるホスト (プラットフォーム別)**

| |**Windows デスクトップ版 Office**|**Office Online (ブラウザー)**|**デバイス用 OWA**|
|:-----|:-----|:-----|:-----|
|**Outlook**|Y|||

|||
|:-----|:-----|
|**要件セットに指定できるもの**|メールボックス|
|**最小限のアクセス許可レベル**|[ReadWriteItem](../../docs/outlook/understanding-outlook-add-in-permissions.md)|
|**アドインの種類**|Outlook|
|**ライブラリ**|Office.js|
|**名前空間**|Office|

## サポート履歴




|**変更内容**|**1.1**|
|:-----|:-----|
|1.3|導入|
