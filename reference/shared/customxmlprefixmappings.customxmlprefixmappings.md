
# CustomXmlPrefixMappings オブジェクト
カスタム名前空間プレフィックス マッピングのコレクションを表します。

|||
|:-----|:-----|
|**ホスト:**|Word|
|**[要件セット](../../docs/overview/specify-office-hosts-and-api-requirements.md)に指定できるもの**|CustomXmlParts|
|**最終変更バージョン**|1.1|

```
CustomXmlPrefixMappings
```


## メンバー


**メソッド**


|**名前**|**説明**|
|:-----|:-----|
|[addNamespaceAsync](../../reference/shared/customxmlprefixmappings.addnamespaceasync.md)|アイテムのクエリを実行するときに使用するプレフィックスを名前空間マッピングに非同期で追加します。|
|[getNamespaceAsync](../../reference/shared/customxmlprefixmappings.getnamespaceasync.md)|指定したプレフィックスにマップされた名前空間を非同期で取得します。|
|[getPrefixAsync](../../reference/shared/customxmlprefixmappings.getprefixasync.md)|指定した名前空間のプレフィックスを非同期で取得します。|

## サポートの詳細


次の表で、大文字 Y は、このメソッドは、対応する Office ホスト アプリケーションでサポートされていることを示します。空のセルは、Office ホスト アプリケーションでこのメソッドをサポートしないことを示します。

Office ホスト アプリケーションとサーバーの要件の詳細については、「[Office アドインを実行するための要件](../../docs/overview/requirements-for-running-office-add-ins.md)」をご覧ください。


||**Windows デスクトップ版 Office**|**Office Online (ブラウザー)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Word**|Y||Y|

|||
|:-----|:-----|
|**要件セットに指定できるもの**|CustomXmlParts|
|**最小限のアクセス許可レベル**|[ReadWriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**アプリの種類**|作業ウィンドウ|
|**ライブラリ**|Office.js|
|**名前空間**|Office|

## サポート履歴



****


|**変更内容**|**1.1**|
|:-----|:-----|
|1.1|Office for iPad における Word のサポートが追加されました。|
|1.0|導入|
