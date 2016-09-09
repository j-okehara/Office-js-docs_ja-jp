
# ValueFormat 列挙型
呼び出されたメソッドが返す値 (数字、日付など) を、書式設定して返すかどうかを指定します。

|||
|:-----|:-----|
|**ホスト:**|Excel、Project、Word|
|**で追加**|1.0|

```
Office.ValueFormat
```


## メンバー


**値**


|**列挙体**|**値**|**説明**|
|:-----|:-----|:-----|
|Office.ValueFormat.Formatted|"formatted"|書式設定されたデータを返します。|
|Office.ValueFormat.Unformatted|"unformatted"|書式設定されていないデータを返します。|

## 注釈

たとえば、 _valueFormat_ パラメーターが `"formatted"` に指定されている場合、ホスト アプリケーションで通貨の書式が設定されている数字や、mm/dd/yy の書式が設定されている日付は、その書式を保持します。 _valueFormat_ パラメーターが `"unformatted"` に指定されている場合、日付は、基本的な順次シリアル番号形式で返されます。


## サポートの詳細


次の表で、大文字 Y は、この列挙は、対応する Office ホスト アプリケーションでサポートされていることを示します。空のセルは、Office ホスト アプリケーションがこの列挙をサポートしないことを示します。

Office ホスト アプリケーションとサーバーの要件の詳細については、「[Office アドインを実行するための要件](../../docs/overview/requirements-for-running-office-add-ins.md)」をご覧ください。


**サポートされるホスト (プラットフォーム別)**


||**Windows デスクトップ版 Office**|**Office Online (ブラウザー)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|Y|Y|Y|
|**Project**|Y|||
|**Word**|Y||Y|

|||
|:-----|:-----|
|**アプリの種類**|コンテンツ、作業ウィンドウ|
|**ライブラリ**|Office.js|
|**名前空間**|Office|

## サポート履歴



****


|**変更内容**|**1.1**|
|:-----|:-----|
|1.1|Office for iPad における Excel と Word のサポートが追加されました。|
|1.0|導入|
