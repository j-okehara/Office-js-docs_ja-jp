
# <a name="customxmlparts-object"></a>CustomXmlParts オブジェクト
[CustomXMLPart](../../reference/shared/customxmlpart.customxmlpart.md) オブジェクトのコレクションを表します。

|||
|:-----|:-----|
|**ホスト:**|Word|
|**[要件セット](../../docs/overview/specify-office-hosts-and-api-requirements.md)に指定できるもの**|CustomXmlParts|
|**最終変更バージョン**|1.1|

```
Office.context.document.customXmlParts
```


## <a name="members"></a>メンバー


**メソッド**


|**名前**|**説明**|
|:-----|:-----|
|[addAsync](../../reference/shared/customxmlparts.addasync.md)|新しいカスタム XML パーツをファイルに非同期に追加します。|
|[getByIdAsync](../../reference/shared/customxmlparts.getbyidasync.md)|ID を使用してカスタム XML パーツを非同期に取得します。|
|[getByNamespaceAsync](../../reference/shared/customxmlparts.getbynamespaceasync.md)|指定された名前空間に一致するカスタム XML パーツの配列を非同期的に取得します。|

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
