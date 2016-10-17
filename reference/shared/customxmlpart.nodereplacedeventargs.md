
# <a name="nodereplacedeventargs-object"></a>NodeReplacedEventArgs オブジェクト
[dataNodeReplaced](../../reference/shared/customxmlpart.datanodereplaced.event.md) イベントの発生元の置き換えられたノードに関する情報を提供します。

|||
|:-----|:-----|
|**ホスト:**|Word|
|**[要件セット](../../docs/overview/specify-office-hosts-and-api-requirements.md)に指定できるもの**|CustomXmlParts|
|**最終変更バージョン**|1.1|

```
NodeReplacedEventArgs
```


## <a name="members"></a>メンバー


**プロパティ**


|**名前**|**説明**|
|:-----|:-----|
|[isUndoRedo](../../reference/shared/customxmlpart.isundoredo.md)|置き換えられたノードが、ユーザーによる元に戻すまたはやり直し操作の一部として挿入されたかどうかを取得します。|
|[newNode](../../reference/shared/customxmlpart.newnode.md)|新しいノードを取得します。|
|[oldNode](../../reference/shared/customxmlpart.oldnode.md)|前の (置き換えられた) ノードを取得します。|

## <a name="support-details"></a>サポートの詳細


次の表で、大文字 Y は、このオブジェクトは、対応する Office ホスト アプリケーションでサポートされていることを示します。空のセルは、Office ホスト アプリケーションでこのオブジェクトをサポートしないことを示します。

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
