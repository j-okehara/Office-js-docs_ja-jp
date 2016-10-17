
# <a name="filtertype-enumeration"></a>FilterType 列挙
データを取得するときにホスト アプリケーションからのフィルタリングを適用するかどうかを指定します。

|||
|:-----|:-----|
|**ホスト:**|Excel、Project、Word|
|**最終変更バージョン**|1.1|

```js
Office.FilterType
```


## <a name="members"></a>メンバー


**値**


|**列挙**|**値**|**説明**|
|:-----|:-----|:-----|
|Office.FilterType.All|"all"|(ホスト アプリケーションによってフィルターが適用されていない) すべてのデータを返します。|
|Office.FilterType.OnlyVisible|"onlyVisible"|(ホスト アプリケーションによってフィルターが適用された) 表示可能なデータのみを返します。|

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
