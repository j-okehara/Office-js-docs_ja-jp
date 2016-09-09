
# Slice.data プロパティ
ファイル スライスの生データを取得します。

|||
|:-----|:-----|
|**ホスト:**|PowerPoint、Word|
|**[要件セット](../../docs/overview/specify-office-hosts-and-api-requirements.md)に指定できるもの**|ファイル|
|**最終変更バージョン**|1.1|

```
var sliceData = slice.data;
```


## 戻り値

**Document.getFileAsync** メソッドの呼び出しの **fileType** パラメーターに指定された _Office.FileType.Text_ ("text") または [Office.FileType.Compressed](../../reference/shared/document.getfileasync.md) ("compressed") 形式のファイル スライスの生データ。


## 注釈

"compressed" 形式のファイルは、必要に応じて Base64 エンコード文字列に変換できるバイト配列を返します。


## サポートの詳細


次の表で、大文字 Y は、このプロパティは、対応する Office ホスト アプリケーションでサポートされていることを示します。空のセルは、Office ホスト アプリケーションでこのプロパティをサポートしないことを示します。

Office ホスト アプリケーションとサーバーの要件の詳細については、「[Office アドインを実行するための要件](../../docs/overview/requirements-for-running-office-add-ins.md)」をご覧ください。


||**Windows デスクトップ版 Office**|**Office Online (ブラウザー)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**PowerPoint**|Y|Y|Y|
|**Word**|Y|Y|Y|


|||
|:-----|:-----|
|**要件セットに指定できるもの**|ファイル|
|**最小限のアクセス許可レベル**|[ReadDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**アプリの種類**|コンテンツ、作業ウィンドウ|
|**ライブラリ**|Office.js|
|**名前空間**|Office|

## サポート履歴



****


|**変更内容**|**1.1**|
|:-----|:-----|
|1.1|Office for iPad で PowerPoint と Word のサポートが追加されました。|
|1.0|導入|
