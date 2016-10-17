
# <a name="table-enumeration"></a>Table 列挙型
[テーブルの書式設定メソッド](../../docs/excel/format-tables-in-add-ins-for-excel.md)の _cellFormat_ パラメーターの `cells:` プロパティに列挙値を指定します。

|||
|:-----|:-----|
|**ホスト:**|Excel|
|**追加**|1.1|

```
Office.Table
```

## <a name="members"></a>メンバー


**値**


|**列挙**|**値**|**説明**|
|:-----|:-----|:-----|
|Office.Table.All|"all"|Office.Table.Data|
|Office.Table.Data|"data"|Office.Table.Headers|
|Office.Table.Headers|"headers"|見出し行のみ。|

## <a name="support-details"></a>サポートの詳細


次の表で、大文字 Y は、対応する Office ホスト アプリケーションでサポートされている列挙を示します。空のセルは、Office ホスト アプリケーションがこの列挙をサポートしないことを示します。

Office ホスト アプリケーションとサーバーの要件の詳細については、「[Office アドインを実行するための要件](../../docs/overview/requirements-for-running-office-add-ins.md)」をご覧ください。


**サポートされるホスト (プラットフォーム別)**


||**Windows デスクトップ版 Office**|**Office Online (ブラウザー)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|Y|Y|Y|

|||
|:-----|:-----|
|**アドインの種類**|コンテンツ、作業ウィンドウ|
|**ライブラリ**|Office.js|
|**名前空間**|Office|

## <a name="support-history"></a>サポート履歴




|**バージョン**|**変更内容**|
|:-----|:-----|
|1.1|Office for iPad の Excel のサポートが追加されました。|
|1.1|導入|
