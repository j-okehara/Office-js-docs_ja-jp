
# NodeDeletedEventArgs オブジェクト
[dataNodeDeleted](../../reference/shared/customxmlpart.datanodedeleted.event.md) イベントの発生元の削除されたノードに関する情報を提供します。

|||
|:-----|:-----|
|**ホスト:**|Word|
|**[要件セット](../../docs/overview/specify-office-hosts-and-api-requirements.md)に指定できるもの**|CustomXmlParts|
|**で追加**|1.1|

```
NodeDeletedEventArgs
```


## メンバー


**プロパティ**


|**名前**|**説明**|
|:-----|:-----|
|[isUndoRedo](../../reference/shared/customxmlpart.isundoredo.md)|ノードが、ユーザーによる元に戻す/やり直し操作の一部として削除されたかどうかを取得します。|
|[oldNextSibling](../../reference/shared/customxmlpart.oldnextsibling.md)|**CustomXMLPart** オブジェクトから削除されたノードの以前の次の兄弟を取得します。|
|[oldNode](../../reference/shared/customxmlpart.oldnode.md)|**CustomXmlPart** オブジェクトから削除されたノードを取得します。|

## サポートの詳細


次の表で、大文字 Y は、このオブジェクトは、対応する Office ホスト アプリケーションでサポートされていることを示します。空のセルは、Office ホスト アプリケーションでこのオブジェクトをサポートしないことを示します。

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




|**変更内容**|**1.1**|
|:-----|:-----|
|1.1|Office for iPad における Word のサポートが追加されました。|
|1.0|導入|
