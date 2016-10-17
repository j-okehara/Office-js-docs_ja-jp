
# <a name="initializationreason-enumeration"></a>InitializationReason 列挙型
ドキュメントにアドインが挿入されたばかりであるか、既に含まれていたかを指定します。 

|||
|:-----|:-----|
|**ホスト:**|Excel、Project、Word|
|**追加されたバージョン**|1.0|

```
Office.InitializationReason
```


## <a name="members"></a>メンバー


**値**


|**列挙**|**値**|**説明**|
|:-----|:-----|:-----|
|Office.InitializationReason.Inserted|"inserted"|アドイン は、ドキュメントに挿入されたばかりです。|
|Office.InitializationReason.DocumentOpened|"documentOpened"|アドイン は、開かれたドキュメントの一部として既に含まれています。|

## <a name="support-details"></a>サポートの詳細


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
|**アドインの種類**|コンテンツ、作業ウィンドウ|
|**ライブラリ**|Office.js|
|**名前空間**|Office|

## <a name="support-history"></a>サポート履歴




|**バージョン**|**変更内容**|
|:-----|:-----|
|1.1|Office for iPad における Excel と Word のサポートが追加されました。|
|1.0|導入|
