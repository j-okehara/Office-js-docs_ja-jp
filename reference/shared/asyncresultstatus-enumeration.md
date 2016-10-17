
# <a name="asyncresultstatus-enumeration"></a>AsyncResultStatus 列挙型
非同期呼び出しの結果を指定します。 

|||
|:-----|:-----|
|**ホスト:**|Access、Excel、Outlook、PowerPoint、Project、Word|
|**最終変更バージョン**|1.1|

```
Office.AsyncResultStatus
```


## <a name="members"></a>メンバー


**値**


|**列挙**|**値**|**説明**|
|:-----|:-----|:-----|
|Office.AsyncResultStatus.Succeeded|"succeeded"|呼び出しが成功しました。|
|Office.AsyncResultStatus.Failed|"failed"|呼び出しが失敗しました。|

## <a name="remarks"></a>注釈

[AsyncResult](../../reference/shared/asyncresult.status.md) オブジェクトの [status](../../reference/shared/asyncresult.md) プロパティによって返されます。


## <a name="support-details"></a>サポートの詳細


次の表で、大文字 Y は、この列挙は、対応する Office ホスト アプリケーションでサポートされていることを示します。空のセルは、Office ホスト アプリケーションがこの列挙をサポートしないことを示します。


Office ホスト アプリケーションとサーバーの要件の詳細については、「[Office アドインを実行するための要件](../../docs/overview/requirements-for-running-office-add-ins.md)」をご覧ください。


**サポートされるホスト (プラットフォーム別)**


||**Windows デスクトップ版 Office**|**Office Online (ブラウザー)**|**Office for iPad**|**デバイス用 OWA**|**Office for Mac**|
|:-----|:-----|:-----|:-----|:-----|:-----|
|**Access**|Y|||||
|**Excel**|Y|Y|Y|||
|**Outlook**|Y|Y||Y|Y|
|**PowerPoint**|Y|Y|Y|||
|**Project**|Y|||||
|**Word**|Y||Y|||

|||
|:-----|:-----|
|**アドインの種類**|コンテンツ、作業ウィンドウ、Outlook|
|**ライブラリ**|Office.js|
|**名前空間**|Office|

## <a name="support-history"></a>サポート履歴


|**バージョン**|**変更内容**|
|:-----|:-----|
|1.1|Office for iPad で Excel、PowerPoint、および Word のサポートが追加されました。|
|1.1|Access 用のアドインのサポートが追加されました。|
|1.0|導入|
