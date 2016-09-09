
# MatrixBinding オブジェクト
行と列の 2 次元でバインドを表現します。 

|||
|:-----|:-----|
|**ホスト:**|Excel、Word|
|**[要件セット](../../docs/overview/specify-office-hosts-and-api-requirements.md)に指定できるもの**|MatrixBindings|
|**選択内容の最終変更**|1.1|

```
MatrixBinding
```


**プロパティ**


|**名前**|**説明**|
|:-----|:-----|
|[columnCount](../../reference/shared/binding.matrixbinding.columncount.md)|マトリックス データ構造内の列数を整数値で取得します。|
|[rowCount](../../reference/shared/binding.matrixbinding.rowcount.md)|マトリックス データ構造内の行数を整数値で取得します。|

## 注釈

**MatrixBinding** オブジェクトは、[id](../../reference/shared/binding.id.md) プロパティ、[type](../../reference/shared/binding.type.md) プロパティ、[getDataAsync](../../reference/shared/binding.getdataasync.md) メソッド、および [setDataAsync](../../reference/shared/binding.setdataasync.md) メソッドを [Binding](../../reference/shared/binding.md) オブジェクトから継承します。


## サポートの詳細


次の表で、大文字 Y は、このメソッドは、対応する Office ホスト アプリケーションでサポートされていることを示します。空のセルは、Office ホスト アプリケーションでこのメソッドをサポートしないことを示します。

Office ホスト アプリケーションとサーバーの要件の詳細については、「[Office アドインを実行するための要件](../../docs/overview/requirements-for-running-office-add-ins.md)」をご覧ください。


**サポートされるホスト (プラットフォーム別)**


||**Windows デスクトップ版 Office**|**Office Online (ブラウザー)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|Y|Y|Y|
|**Word**|Y||Y|

|||
|:-----|:-----|
|**要件セットに指定できるもの**|MatrixBindings|
|**アプリの種類**|コンテンツ、作業ウィンドウ|
|**ライブラリ**|Office.js|
|**名前空間**|Office|

## サポート履歴



****


|**変更内容**|**1.1**|
|:-----|:-----|
|1.1|Office for iPad における Excel と Word のサポートが追加されました。|
|1.0|導入|
