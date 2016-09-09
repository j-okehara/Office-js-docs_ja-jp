
# TableData オブジェクト
テーブルまたは [TableBinding](../../reference/shared/binding.tablebinding.md) 内のデータを表します。

|||
|:-----|:-----|
|**ホスト:**|Excel、Word|
|**[要件セット](../../docs/overview/specify-office-hosts-and-api-requirements.md)に指定できるもの**|TableBindings|
|**で追加**|1.1|

```
TableData
```

## メンバー


**プロパティ**


|**名前**|**説明**|
|:-----|:-----|
|[headers](../../reference/shared/tabledata.headers.md)|テーブル内のヘッダーを取得または設定します。|
|[rows](../../reference/shared/tabledata.rows.md)|テーブル内の行を取得または設定します。|

## サポートの詳細


次の表で、大文字 Y は、このオブジェクトは、対応する Office ホスト アプリケーションでサポートされていることを示します。空のセルは、Office ホスト アプリケーションでこのオブジェクトをサポートしないことを示します。

Office ホスト アプリケーションとサーバーの要件の詳細については、「[Office アドインを実行するための要件](../../docs/overview/requirements-for-running-office-add-ins.md)」をご覧ください。


||**Windows デスクトップ版 Office**|**Office Online (ブラウザー)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|Y|Y|Y|
|**Word**|Y|Y|Y|

|||
|:-----|:-----|
|**要件セットに指定できるもの**|TableBindings|
|**アプリの種類**|コンテンツ、作業ウィンドウ|
|**ライブラリ**|Office.js|
|**名前空間**|Office|

## サポート履歴




|**変更内容**|**1.1**|
|:-----|:-----|
|1.1|Word Online のサポートが追加されました。|
|1.1|Office for iPad における Excel と Word のサポートが追加されました。|
|1.0|導入|
