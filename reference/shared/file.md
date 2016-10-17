
# <a name="file-object"></a>File オブジェクト
Office アドインに関連付けられているドキュメント ファイルを表します。

|||
|:-----|:-----|
|**ホスト:**|PowerPoint、Word|
|**[要件セット](../../docs/overview/specify-office-hosts-and-api-requirements.md)に指定できるもの**|ファイル|
|**最終変更バージョン**|1.1|

```
file
```


## <a name="members"></a>メンバー


**プロパティ**


|**名前**|**説明**|
|:-----|:-----|
|**[size](../../reference/shared/file.size.md)**|ドキュメントのファイル サイズをバイト単位で取得します。|
|**[sliceCount](../../reference/shared/file.slicecount.md)**|ファイルが分割されるスライス数を取得します。|

**メソッド**


|**名前**|**説明**|
|:-----|:-----|
|**[closeAsync](../../reference/shared/file.closeasync.md)**|ドキュメント ファイルを閉じます。|
|**[getSliceAsync](../../reference/shared/file.getsliceasync.md)**|指定したスライスを返します。|

## <a name="remarks"></a>注釈

**File** オブジェクトには、[Document.getFileAsync](../../reference/shared/asyncresult.value.md) メソッドに渡されるコールバック関数の [AsyncResult.value](../../reference/shared/document.getfileasync.md) プロパティを使用してアクセスします。


## <a name="support-details"></a>サポートの詳細


次の表で、大文字 Y は、このオブジェクトは、対応する Office ホスト アプリケーションでサポートされていることを示します。空のセルは、Office ホスト アプリケーションでこのオブジェクトをサポートしないことを示します。

Office ホスト アプリケーションとサーバーの要件の詳細については、「[Office アドインを実行するための要件](../../docs/overview/requirements-for-running-office-add-ins.md)」をご覧ください。


|||||
|:-----|:-----|:-----|:-----|
||Windows デスクトップ版 Office|Office Online (ブラウザー)|Office for iPad|
|**PowerPoint**|Y|Y|Y|
|**Word**|Y||Y|

|||
|:-----|:-----|
|**要件セットに指定できるもの**|ファイル|
|**アドインの種類**|コンテンツ、作業ウィンドウ|
|**ライブラリ**|Office.js|
|**名前空間**|Office|

## <a name="support-history"></a>サポート履歴



****


|**バージョン**|**変更内容**|
|:-----|:-----|
|1.1|Office for iPad で PowerPoint と Word のサポートが追加されました。|
|1.0|導入|
