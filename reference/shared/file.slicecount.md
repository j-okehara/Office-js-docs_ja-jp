
# <a name="file.slicecount-property"></a>File.sliceCount プロパティ
ファイルが分割されるスライス数を取得します。

|||
|:-----|:-----|
|**ホスト:**|PowerPoint、Word|
|**追加されたバージョン**|1.1|

```
var slices = file.sliceCount;
```


## <a name="return-value"></a>戻り値

スライス数。


## <a name="support-details"></a>サポートの詳細


次の表で、大文字 Y は、このメソッドは、対応する Office ホスト アプリケーションでサポートされていることを示します。空のセルは、Office ホスト アプリケーションでこのメソッドをサポートしないことを示します。

Office ホスト アプリケーションとサーバーの要件の詳細については、「[Office アドインを実行するための要件](../../docs/overview/requirements-for-running-office-add-ins.md)」をご覧ください。


|||||
|:-----|:-----|:-----|:-----|
||Windows デスクトップ版 Office|Office Online (ブラウザー)|Office for iPad|
|**PowerPoint**|Y|Y|Y|
|**Word**|Y|Y|Y|

|||
|:-----|:-----|
|**最小限のアクセス許可レベル**|[制限あり](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**アドインの種類**|コンテンツ、作業ウィンドウ|
|**ライブラリ**|Office.js|
|**名前空間**|Office|

## <a name="support-history"></a>サポート履歴



****


|**バージョン**|**変更内容**|
|:-----|:-----|
|1.1|Office for iPad で PowerPoint と Word のサポートが追加されました。|
|1.0|導入|
