
# <a name="customxmlnode-object"></a>CustomXmlNode オブジェクト
ドキュメント内のツリーの XML ノードを表します。

|||
|:-----|:-----|
|**ホスト:**|Word|
|**[要件セット](../../docs/overview/specify-office-hosts-and-api-requirements.md)に指定できるもの**|CustomXmlParts|
|**最終変更バージョン**|1.1|

```js
CustomXmlNode
```


## <a name="members"></a>メンバー


**プロパティ**


|**名前**|**説明**|
|:-----|:-----|
|[baseName](../../reference/shared/customxmlnode.basename.md)|名前空間プレフィックスを持たないノードがある場合、そのベース名を取得します。|
|[nodeType](../../reference/shared/customxmlnode.nodetype.md)|**CustomXMLNode** の種類を取得します。|
|[namespaceUri](../../reference/shared/customxmlnode.namespaceuri.md)|**CustomXMLPart** の GUID を文字列で取得します。|

**メソッド**


|**名前**|**説明**|
|:-----|:-----|
|[getNodesAsync](../../reference/shared/customxmlnode.getnodesasync.md)|相対 XPath 式と一致する **CustomXMLNode** オブジェクトの配列としてノードを非同期的に取得します。|
|[getNodeValueAsync](../../reference/shared/customxmlnode.getnodevalueasync.md)|ノードの値を非同期的に取得します。|
|[getTextAsync](customxmlnode.gettextasync.md)|カスタム XML パーツ内の XML ノードのテキストを非同期的に取得します。|
|[getXmlAsync](../../reference/shared/customxmlnode.getxmlasync.md)|ノードの XML を非同期的に取得します。|
|[setNodeValueAsync](../../reference/shared/customxmlnode.setnodevalueasync.md)|ノードの値を非同期に設定します。|
|[setTextAsync](customxmlnode.settextasync.md)|カスタム XML パーツ内の XML ノードのテキストを非同期的に設定します。|
|[setXmlAsync](../../reference/shared/customxmlnode.setxmlasync.md)|ノードの XML を非同期的に設定します。|

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
