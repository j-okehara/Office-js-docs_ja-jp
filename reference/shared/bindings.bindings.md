
# <a name="bindings-object"></a>Bindings オブジェクト
アドインがドキュメント内に持つバインドを表します。

|||
|:-----|:-----|
|**ホスト:**|Access、Excel、Word|
|の **最終変更**|1.1|

```js
Office.context.document.bindings
```


**プロパティ**

|||
|:-----|:-----|
|名前|説明|
|[document](../../reference/shared/bindings.document.md)|このバインドのセットに関連付けられたドキュメントを表す **Document** オブジェクトを取得します。|

**メソッド**

|||
|:-----|:-----|
|名前|説明|
|[addFromNamedItemAsync](../../reference/shared/bindings.addfromnameditemasync.md)|ドキュメント内の名前付きの項目にバインドを追加します。|
|[addFromPromptAsync](../../reference/shared/bindings.addfrompromptasync.md)|ユーザーがバインド先の選択範囲を指定するための UI を表示します。|
|[addFromSelectionAsync](../../reference/shared/bindings.addfromselectionasync.md)|種類が指定されたバインドのバインド オブジェクトをドキュメント内の現在の選択範囲に追加します。|
|[getAllAsync](../../reference/shared/bindings.getallasync.md)|以前に作成されたバインドをすべて取得します。|
|[getByIdAsync](../../reference/shared/bindings.getbyidasync.md)|指定したバインドを ID によって取得します。|
|[releaseByIdAsync](../../reference/shared/bindings.releasebyidasync.md)|指定したバインドを削除します。|

## <a name="support-details"></a>サポートの詳細


次の表で、大文字 Y は、このメソッドは、対応する Office ホスト アプリケーションでサポートされていることを示します。空のセルは、Office ホスト アプリケーションでこのメソッドをサポートしないことを示します。

Office ホスト アプリケーションとサーバーの要件の詳細については、「[Office アドインを実行するための要件](../../docs/overview/requirements-for-running-office-add-ins.md)」をご覧ください。


|||||
|:-----|:-----|:-----|:-----|
||Windows デスクトップ版 Office|Office Online (ブラウザー)|Office for iPad|
|**Access**||Y||
|**Excel**|Y|Y|Y|
|**Word**|Y|Y|Y|

|||
|:-----|:-----|
|**アドインの種類**|コンテンツ、作業ウィンドウ|
|**ライブラリ**|Office.js|
|**名前空間**|Office|

## <a name="support-history"></a>サポート履歴



|**バージョン**|**変更内容**|
|:-----|:-----|
|1.1|Office for iPad における Excel と Word のサポートが追加されました。|
|1.1|[addFromNamedItemAsync](../../reference/shared/bindings.addfromnameditemasync.md)、[addFromPromptAsync](../../reference/shared/bindings.addfrompromptasync.md)、および [addFromSelectionAsync](../../reference/shared/bindings.addfromselectionasync.md) で、Excel 用アプリのテーブル バインドとしてマトリックス データへのバインドのサポートが追加されました。|
|1.1|<ul><li><a href="8fa0cb4a-fad1-4f2e-9a7e-5f7aa7789eca.htm">document</a> プロパティで、Access 用コンテンツ アドインの現在の Access データベースを表す <span class="keyword">Document</span> オブジェクトへのアクセスが追加されました。</li><li>すべてのメソッドで、Access 用コンテンツ アドインにおけるテーブル バインドのサポートが追加されました。 </li></ul>|
|1.0|導入|
