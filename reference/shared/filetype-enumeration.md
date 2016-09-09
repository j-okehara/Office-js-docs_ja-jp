
# FileType 列挙型
ドキュメントを返す形式を指定します。

|||
|:-----|:-----|
|**ホスト:**|PowerPoint、Word|
|**最終変更バージョン**|1.1|

```js
Office.FileType
```


## メンバー


**値**


|**列挙体**|**値**|**Office.FileType.Compressed**|
|:-----|:-----|:-----|
|"compressed"|"compressed"|ドキュメント全体 (.pptx または .docx) を Office Open XML (OOXML) 形式でバイト配列として返します。|
|Office.FileType.Pdf|PDF 形式のドキュメント全体をバイト配列として返します。|Office.FileType.Text|
|Office.FileType.Text|"text"|ドキュメントのテキストのみを  **string** として返します。(Word のみ)|

## サポートの詳細


次の表で、大文字 Y は、この列挙は、対応する Office ホスト アプリケーションでサポートされていることを示します。空のセルは、Office ホスト アプリケーションがこの列挙をサポートしないことを示します。

Office ホスト アプリケーションとサーバーの要件の詳細については、「[Office アドインを実行するための要件](../../docs/overview/requirements-for-running-office-add-ins.md)」をご覧ください。


**サポートされるホスト (プラットフォーム別)**


||**Windows デスクトップ版 Office**|**Office Online (ブラウザー)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**PowerPoint**|Y|Y|Y|
|**Word**|Y||Y|

|||
|:-----|:-----|
|**アプリの種類**|コンテンツ、作業ウィンドウ|
|**ライブラリ**|Office.js|
|**名前空間**|Office|

## サポート履歴


|**変更内容**|**1.1**|
|:-----|:-----|
|1.1|Office for iPad で PowerPoint と Word のサポートが追加されました。|
|1.1|PDF として保存するためのサポートが追加されました。|
|1.0|導入|
