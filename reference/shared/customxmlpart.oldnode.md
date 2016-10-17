
# <a name="nodedeletedeventargs.oldnode-property"></a>NodeDeletedEventArgs.oldNode プロパティ
**CustomXmlPart** オブジェクトから削除されたノードを取得します。

|||
|:-----|:-----|
|**ホスト:**|Word|
|**[要件セット](../../docs/overview/specify-office-hosts-and-api-requirements.md)に指定できるもの**|CustomXmlParts|
|**最終変更バージョン**|1.1|

```
var myNode = eventArgsObj.oldNode;
```


## <a name="return-value"></a>戻り値

削除されたノードを表す [CustomXmlNode](../../reference/shared/customxmlnode.customxmlnode.md)。


## <a name="remarks"></a>注釈

ドキュメントからサブツリーを削除する場合は、このノードに子が含まれている可能性があるので注意してください。また、このノードよりも下位のレベルにはクエリを実行できますが、ツリーの上位のレベルにはクエリを実行できません。そういう意味で、このノードは "切り離された" ノードです。つまり、このノードは存在しているだけのように見えます。


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




|**バージョン**|**変更内容**|
|:-----|:-----|
|1.1|Office for iPad における Word のサポートが追加されました。|
|1.0|導入|
