
# <a name="documentmode-enumeration"></a>DocumentMode 列挙型
関連付けられているアプリケーションのドキュメントを読み取り専用または読み取り/書き込みのどちらかに指定します。 

|||
|:-----|:-----|
|**ホスト:**|Excel、PowerPoint、Project、Word|
|**追加されたバージョン**|1.1|

```
Office.DocumentMode
```


## <a name="members"></a>メンバー


**値**


|**列挙**|**値**|**説明**|
|:-----|:-----|:-----|
|Office.DocumentMode.ReadOnly|"readOnly"|ドキュメントは、読み取り専用です。|
|Office.DocumentMode.ReadWrite|"readWrite"|ドキュメントは、読み取り/書き込み可能です。|

## <a name="remarks"></a>注釈

**Document** オブジェクトの [mode](../../reference/shared/document.md) プロパティによって返されます。


## <a name="support-details"></a>サポートの詳細


次の表で、大文字 Y は、この列挙は、対応する Office ホスト アプリケーションでサポートされていることを示します。空のセルは、Office ホスト アプリケーションがこの列挙をサポートしないことを示します。

Office ホスト アプリケーションとサーバーの要件の詳細については、「[Office アドインを実行するための要件](../../docs/overview/requirements-for-running-office-add-ins.md)」をご覧ください。


**サポートされるホスト (プラットフォーム別)**


||**Windows デスクトップ版 Office**|**Office Online (ブラウザー)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|Y|Y|Y|
|**PowerPoint**|Y|Y|Y|
|**Project**|Y|||
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
|1.1|Office for iPad で Excel、PowerPoint、および Word のサポートが追加されました。|
|1.0|導入|
