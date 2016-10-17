
# <a name="bindingtype-enumeration"></a>BindingType 列挙型
 返されるバインド オブジェクトの種類を指定します。

|||
|:-----|:-----|
|**ホスト:**|Access、Excel、Word|
|**最終変更**|1.1|

```
Office.BindingType
```


## <a name="members"></a>メンバー


**値**


|**列挙**|**値**|**説明**|
|:-----|:-----|:-----|
|Office.BindingType.Matrix|"matrix"|ヘッダー行なしの表形式データ。データは、次のような配列の配列として返されます。` [[row1column1, row1column2],[row2column1, row2column2]]`|
|Office.BindingType.Table|"table"|ヘッダー行ありの表形式データ。データは、[TableData](../../reference/shared/tabledata.md) オブジェクトとして返されます。|
|Office.BindingType.Text|"text"|プレーンテキスト。データは、一連の文字として返されます。|

## <a name="support-details"></a>サポートの詳細


次の表で、大文字 Y は、この列挙は、対応する Office ホスト アプリケーションでサポートされていることを示します。空のセルは、Office ホスト アプリケーションがこの列挙をサポートしないことを示します。

Office ホスト アプリケーションとサーバーの要件の詳細については、「[Office アドインを実行するための要件](../../docs/overview/requirements-for-running-office-add-ins.md)」をご覧ください。


**サポートされるホスト (プラットフォーム別)**


||**Windows デスクトップ版 Office**|**Office Online (ブラウザー)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Access**|Y|||
|**Excel**|Y|Y|Y|
|**Word**|Y|Y|Y|

|||
|:-----|:-----|
|**アドインの種類**|コンテンツ、作業ウィンドウ|
|**ライブラリ**|Office.js|
|**名前空間**|Office|

## <a name="support-history"></a>サポート履歴



|**バージョン**|**変更内容**|
|:-----|:-----|
|1.1|Office for iPad における Excel と Word のサポートが追加されました。|
|1.1|
            Access 用アプリにおけるデータのバインドのサポートが追加されました。|
|1.0|導入。|
