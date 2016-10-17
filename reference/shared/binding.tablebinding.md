
# <a name="tablebinding-object"></a>TableBinding オブジェクト
バインドを行と列の 2 次元で、必要に応じてヘッダーと共に表します。

|||
|:-----|:-----|
|**ホスト:**|Access、Excel、PowerPoint、Project、Word|
|**[要件セット](../../docs/overview/specify-office-hosts-and-api-requirements.md)に指定できるもの**|TableBindings|
|**選択内容の最終変更**|1.1|

```
TableBinding
```


## <a name="members"></a>メンバー


**プロパティ**


|**名前**|**説明**|**Office.js v1.1 での更新**|
|:-----|:-----|:-----|
|[columnCount](../../reference/shared/binding.tablebinding.columncount.md)|指定された **TableBinding** オブジェクト内の列数を取得します。|Access 用コンテンツ アドインにおけるテーブル バインドのサポートが追加されました。|
|[hasHeaders](../../reference/shared/binding.tablebinding.hasheaders.md)|指定された **TableBinding** にヘッダーがある場合は true を返します。ヘッダーがない場合は false を返します。|Access 用コンテンツ アドインにおけるテーブル バインドのサポートが追加されました。|
|[rowCount](../../reference/shared/binding.tablebinding.rowcount.md)|指定された **TableBinding** オブジェクト内の行数。|パフォーマンス上の理由から、Access 用コンテンツ アプリでは常に -1 を返します。|

**メソッド**


|**名前**|**説明**|**Office.js v1.1 での更新**|
|:-----|:-----|:-----|
|[addColumnsAsync](../../reference/shared/binding.tablebinding.addcolumnsasync.md)|テーブルに列と値を追加します。||
|[addRowsAsync](../../reference/shared/binding.tablebinding.addrowsasync.md)|テーブルに行と値を追加します。|Access 用コンテンツ アドインにおけるテーブル バインドのサポートが追加されました。|
|[clearFormatsAsync](../../reference/shared/binding.tablebinding.clearformatsasync.md)|バインド テーブルの書式設定をクリアします。|Excel 用アプリの Office.js v1.1 における新機能です。|
|[deleteAllDataValuesAsync](../../reference/shared/binding.tablebinding.deletealldatavaluesasync.md)|テーブル内のヘッダー行以外の行と値をすべて削除し、ホスト アプリケーションに応じて適切にシフトします。|Access 用コンテンツ アドインにおけるテーブル バインドのサポートが追加されました。|
|[setDataAsync](../../reference/shared/binding.setdataasync.md)|指定されたバインド オブジェクトで表されるドキュメントのバインド セクションにデータを書き込みます。|<ul><li>
            Access 用コンテンツ アプリにおけるテーブルのバインドのサポートが追加されました。</li><li>Excel 用アプリでバインド テーブルにデータを書き込むときに書式設定を行うためのサポートが追加されました。</li></ul>|
|[setFormatsAsync](../../reference/shared/binding.tablebinding.setformatsasync.md)|バインド テーブル内の指定アイテムとデータに、セルとテーブルの書式設定を行います。|Excel 用アドインでテーブル書式を設定できます。|
|[setTableOptionsAsync](../../reference/shared/binding.tablebinding.settableoptionsasync.md)|バインド テーブルにおけるテーブル書式設定オプションを更新します。|Excel 用アプリで表書式を設定できます。|

## <a name="remarks"></a>注釈

**TableBinding** オブジェクトは、[id](../../reference/shared/binding.id.md) プロパティ、[type](../../reference/shared/binding.type.md) プロパティ、[getDataAsync](../../reference/shared/binding.getdataasync.md) メソッド、および [setDataAsync](../../reference/shared/binding.setdataasync.md) メソッドを [Binding](../../reference/shared/binding.md) 抽象オブジェクトから継承します。

Excel 内でテーブル バインドを構築すると、ユーザーがテーブルに追加する新しい各行は、それぞれ自動的にバインドに追加されます ( **rowCount** が増加します)。


## <a name="support-details"></a>サポートの詳細


次の表で、大文字 Y は、このオブジェクトは、対応する Office ホスト アプリケーションでサポートされていることを示します。空のセルは、Office ホスト アプリケーションでこのオブジェクトをサポートしないことを示します。

Office ホスト アプリケーションとサーバーの要件の詳細については、「[Office アドインを実行するための要件](../../docs/overview/requirements-for-running-office-add-ins.md)」をご覧ください。


**サポートされるホスト (プラットフォーム別)**


||**Windows デスクトップ版 Office**|**Office Online (ブラウザー)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||Y||
|**Excel**|Y|Y|Y|
|**Word**|Y|Y|Y|

|||
|:-----|:-----|
|**要件セットに指定できるもの**|TableBindings|
|**最小限のアクセス許可レベル**|[WriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**アドインの種類**|コンテンツ、作業ウィンドウ|
|**ライブラリ**|Office.js|
|**名前空間**|Office|

## <a name="support-history"></a>サポート履歴




|**バージョン**|**変更内容**|
|:-----|:-----|
|1.1|Office for iPad における Excel と Word のサポートが追加されました。|
|1.1|Excel における [テーブルを挿入する際の書式設定](../../docs/excel/format-tables-in-add-ins-for-excel.md)のサポートが追加されました。|
|1.1|Access 用のアドインのサポートが追加されました。|
|1.0|導入|
