
# <a name="textbinding-object"></a>TextBinding オブジェクト
ドキュメント内のバインドされているテキスト選択を表します。

|||
|:-----|:-----|
|**ホスト:**|Access、Excel、PowerPoint、Project、Word|
|**[要件セット](../../docs/overview/specify-office-hosts-and-api-requirements.md)に指定できるもの**|TextBindings|
|**追加されたバージョン**|1.0|

```
TextBinding
```


## <a name="remarks"></a>注釈

**TextBinding** オブジェクトは、[id](../../reference/shared/binding.id.md) プロパティ、[type](../../reference/shared/binding.type.md) プロパティ、[getDataAsync](../../reference/shared/binding.getdataasync.md) メソッド、および [setDataAsync](../../reference/shared/binding.setdataasync.md) メソッドを [Binding](../../reference/shared/binding.md) オブジェクトから継承します。これ以外に、このオブジェクトが独自に実装するプロパティやメソッドはありません。


## <a name="support-details"></a>サポートの詳細


次の表で、大文字 Y は、このオブジェクトは、対応する Office ホスト アプリケーションでサポートされていることを示します。空のセルは、Office ホスト アプリケーションでこのオブジェクトをサポートしないことを示します。

Office ホスト アプリケーションとサーバーの要件の詳細については、「[Office アドインを実行するための要件](../../docs/overview/requirements-for-running-office-add-ins.md)」をご覧ください。


**サポートされるホスト (プラットフォーム別)**


||**Windows デスクトップ版 Office**|**Office Online (ブラウザー)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|Y|Y|Y|
|**Word**|Y|Y|Y|

|||
|:-----|:-----|
|**要件セットに指定できるもの**|TextBindings|
|**最小限のアクセス許可レベル**|[WriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**アドインの種類**|コンテンツ、作業ウィンドウ|
|**ライブラリ**|Office.js|
|**名前空間**|Office|

## <a name="support-history"></a>サポート履歴



****


|**バージョン**|**変更内容**|
|:-----|:-----|
|1.1|Office for iPad における Excel と Word のサポートが追加されました。|
|1.0|導入|
