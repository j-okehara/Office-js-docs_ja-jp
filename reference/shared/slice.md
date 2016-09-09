
# Slice オブジェクト
ドキュメント ファイルのスライスを表します。

|||
|:-----|:-----|
|**ホスト:**|PowerPoint、Word|
|**[要件セット](../../docs/overview/specify-office-hosts-and-api-requirements.md)に指定できるもの**|ファイル|
|**最終変更バージョン**|1.1|

```
slice
```


## メンバー


**プロパティ**


|**名前**|**説明**|
|:-----|:-----|
|**[data](../../reference/shared/slice.data.md)**|ファイル スライスの生データを取得します。|
|**[index](../../reference/shared/slice.index.md)**|ファイル スライスのインデックスを取得します。|
|**[size](../../reference/shared/slice.size.md)**|スライスのサイズをバイト単位で取得します。|

## 注釈

**Slice** オブジェクトには、[File.getSliceAsync](../../reference/shared/file.getsliceasync.md) メソッドを使用してアクセスします。


## サポートの詳細


次の表で、大文字 Y は、このオブジェクトは、対応する Office ホスト アプリケーションでサポートされていることを示します。空のセルは、Office ホスト アプリケーションでこのオブジェクトをサポートしないことを示します。

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




|**変更内容**|**1.1**|
|:-----|:-----|
|1.1|Office for iPad で PowerPoint と Word のサポートが追加されました。|
|1.0|導入|
