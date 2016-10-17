
# <a name="customxmlpart-object"></a>CustomXmlPart オブジェクト
**CustomXMLParts** コレクション内の単一の [CustomXMLPart](../../reference/shared/customxmlparts.customxmlparts.md) を表します。

|||
|:-----|:-----|
|**ホスト:**|Word|
|**[要件セット](../../docs/overview/specify-office-hosts-and-api-requirements.md)に指定できるもの**|CustomXmlParts|
|**最終変更バージョン**|1.1|

```
Office.context.document.customXmlParts.getByIdAsync(id);
```


## <a name="members"></a>メンバー


**プロパティ**


|**名前**|**説明**|
|:-----|:-----|
|[builtIn](../../reference/shared/customxmlpart.builtin.md)|CustomXMLPart が組み込みであるかどうかを示す値を取得します。|
|[id](../../reference/shared/customxmlpart.id.md)|CustomXMLPart の GUID を取得します。|
|[namespaceManager](../../reference/shared/customxmlpart.namespacemanager.md)|現在の CustomXMLPart に対して使用される名前空間プレフィックス マッピング (CustomXMLPrefixMappings) のセットを取得します。|

**メソッド**


|**名前**|**説明**|
|:-----|:-----|
|[addHandlerAsync](../../reference/shared/customxmlpart.addhandlerasync.md)|**CustomXmlPart** オブジェクト イベントのイベント ハンドラーを非同期的に追加します。|
|[deleteAsync](../../reference/shared/customxmlpart.deleteasync.md)|このカスタム XML パーツをコレクションから非同期的に削除します。|
|[getNodesAsync](../../reference/shared/customxmlpart.getnodesasync.md)|指定された XPath に一致するこのカスタム XML パーツ内の CustomXmlNodes を非同期に取得します。|
|[getXmlAsync](../../reference/shared/customxmlpart.getxmlasync.md)|このカスタム XML パーツ内の XML を非同期的に取得します。|
|[removeHandlerAsync](../../reference/shared/customxmlpart.removehandlerasync.md)|**CustomXmlPart** オブジェクト イベントのイベント ハンドラーを削除します。|

**イベント**


|**名前**|**説明**|
|:-----|:-----|
|[dataNodeDeleted](../../reference/shared/customxmlpart.datanodedeleted.event.md)|ノードが削除されるときに発生します。|
|[dataNodeInserted](../../reference/shared/customxmlpart.datanodeinserted.event.md)|ノードが挿入されるときに発生します。|
|[dataNodeReplaced](../../reference/shared/customxmlpart.datanodereplaced.event.md)|ノードが置換されるときに発生します。|

## <a name="support-details"></a>サポートの詳細


次の表で、大文字 Y は、このメソッドは、対応する Office ホスト アプリケーションでサポートされていることを示します。空のセルは、Office ホスト アプリケーションでこのメソッドをサポートしないことを示します。

Office ホスト アプリケーションとサーバーの要件の詳細については、「[Office アドインを実行するための要件](../../docs/overview/requirements-for-running-office-add-ins.md)」をご覧ください。


||**Windows デスクトップ版 Office**|**Office Online (ブラウザー)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Word**|Y|Y|Y|

|||
|:-----|:-----|
|**要件セットに指定できるもの**|CustomXmlParts|
|**最小限のアクセス許可レベル**|[ReadWriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**アドインの種類**|作業ウィンドウ|
|**ライブラリ**|Office.js|
|**名前空間**|Office|

## <a name="support-history"></a>サポート履歴



****


|**バージョン**|**変更内容**|
|:-----|:-----|
|1.1|Office for iPad における Word のサポートが追加されました。|
|1.0|導入|
