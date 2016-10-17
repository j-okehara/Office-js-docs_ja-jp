
# <a name="coerciontype-enumeration"></a>CoercionType 列挙型
呼び出されたメソッドによって返される、または設定されるデータを強制的に変換する方法を指定します。

|||
|:-----|:-----|
|**ホスト:**|Access、Excel、Outlook、PowerPoint、Project、Word|
|**メールボックスの最終変更**|1.1|

```js
Office.CoercionType
```

## <a name="members"></a>メンバー


**値**


|**列挙**|**値**|**説明**|
|:-----|:-----|:-----|
|Office.CoercionType.Html|"html"|データを HTML として取得または設定します。<br/><br/> **注** Word 用のアドインと Outlook 用の Outlook アドイン (新規作成モード) のデータにのみ適用されます。|
|Office.CoercionType.Matrix|"matrix"|データをヘッダーなしの表形式データとして取得または設定します。データは、1 次元の一連の文字を含む配列の配列として取得または設定されます。たとえば、3 行 2 列の構成に含まれる **string** 値は [` [["R1C1", "R1C2"], ["R2C1", "R2C2"], ["R3C1", "R3C2"]]`] と表します。<br/><br/> **注** Excel および Word のデータにのみ適用されます。|
|Office.CoercionType.Ooxml|"ooxml"|データを Office Open XML として取得または設定します。<br/><br/> **注** Word のデータにのみ適用されます。|
|Office.CoercionType.SlideRange|"slideRange"|選ばれているスライドの ID、タイトル、インデックスの配列が含まれる JSON オブジェクトを返します。たとえば、2 つのスライドが選ばれている場合は、[`{"slides":[{"id":257,"title":"Slide 2","index":2},{"id":256,"title":"Slide 1","index":1}]}`] です。<br/><br/> **注** 現在のスライドまたは選択範囲のスライドを取得する [Document.getSelectedData](../../reference/shared/document.getselecteddataasync.md) メソッドを呼び出す場合、PowerPoint のデータにのみ適用されます。|
|Office.CoercionType.Table|"table"|データをヘッダー付き (省略可能) の表形式データとして取得または設定します。データをヘッダー付き (省略可能) の表形式データとして取得または設定します。<br/><br/> **注** Access、Excel、および Word のデータにのみ適用されます。|
|Office.CoercionType.Text|"text"|データをテキスト (**string**) として取得または設定します。データは、1 次元の一連の文字として取得または設定されます。|
|Office.CoercionType.Image|"image"|データは返されるか、イメージ ストリームとして設定されます。<br/><br/> **注** Excel、Word、PowerPoint のデータにのみ適用されます。|
PowerPoint では、**Office.CoercionType.Text**、**Office.CoercionType.Image**、**Office.CoercionType.SlideRange** のみがサポートされます。

Project では、 **Office.CoercionType.Text** のみがサポートされます。


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
|**Word**|Y|Y|Y|||

|||
|:-----|:-----|
|**アドインの種類**|コンテンツ、Outlook (新規作成モード)、作業ウィンドウ|
|**ライブラリ**|Office.js|
|**名前空間**|Office|

## <a name="support-history"></a>サポート履歴


|**バージョン**|**変更内容**|
|:-----|:-----|
|1.1|Word Online のサポートが追加されました。|
|1.1|Office for iPad で Excel、PowerPoint、および Word のサポートが追加されました。|
|1.1|Access 用のアドインのサポートが追加されました。|
|1.1|[新規作成モードの Outlook アドイン](../../docs/outlook/compose-scenario.md)のサポートが追加されました。|
|1.0|導入|
