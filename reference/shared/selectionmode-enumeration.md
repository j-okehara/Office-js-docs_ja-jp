
# <a name="selectionmode-enumeration"></a>SelectionMode 列挙体
移動先の場所を選択 (強調表示) するかどうかを指定します ([Document.goToByIdAsync](../../reference/shared/document.gotobyidasync.md) メソッドを使用する場合)。

|||
|:-----|:-----|
|**導入された Office.js バージョン**|1.1|

|||
|:-----|:-----|
|**ホスト:**|Excel、PowerPoint、Word|
|**追加されたバージョン**|1.1|



```
Office.SelectionMode
```


## <a name="members"></a>メンバー


**値**


|**列挙**|**値**|**説明**|
|:-----|:-----|:-----|
|Office.SelectionMode.Selected|"selected"|場所を選択 (強調表示) します。|
|Office.SelectionMode.None|"none"|カーソルが場所の先頭に移動します。|

## <a name="support-details"></a>サポートの詳細


次の表で、大文字 Y は、このメソッドは、対応する Office ホスト アプリケーションでサポートされていることを示します。空のセルは、Office ホスト アプリケーションでこのメソッドをサポートしないことを示します。

Office ホスト アプリケーションとサーバーの要件の詳細については、「[Office アドインを実行するための要件](../../docs/overview/requirements-for-running-office-add-ins.md)」をご覧ください。


**サポートされるホスト (プラットフォーム別)**


||**Windows デスクトップ版 Office**|**Office Online (ブラウザー)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|Y|Y|Y|
|**PowerPoint**|Y|||
|**Word**|Y||Y|

|||
|:-----|:-----|
|**アドインの種類**|コンテンツ、作業ウィンドウ|
|**ライブラリ**|Office.js|
|**名前空間**|Office|

## <a name="support-history"></a>サポート履歴



****


|**バージョン**|**変更内容**|
|:-----|:-----|
|1.1|導入|
