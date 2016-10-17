
# <a name="gototype-enumeration"></a>GoToType 列挙型
ナビゲートする場所またはオブジェクトの種類を指定します。

|||
|:-----|:-----|
|**ホスト:**|Excel、PowerPoint、Word|
|**追加されたバージョン**|1.1|

```js
Office.GoToType
```


## <a name="members"></a>メンバー


**値**


|**列挙**|**値**|**説明**|**サポートされるクライアント**|
|:-----|:-----|:-----|:-----|
|Office.GoToType.Binding|"binding"|指定されたバインド ID を使用して Binding オブジェクトに移動します。|Excel</br>Word|
|Office.GoToType.NamedItem|"namedItem"|テーブルまたは範囲に割り当てられた名前など、アイテムの名前を使用してアイテムに移動します。Excel で、名前付き範囲またはテーブルの構造化参照を使用できます (例: "Worksheet2!Table1")。|Excel|
|Office.GoToType.Slide|"slide"|指定された ID を使用してスライドに移動します。|PowerPoint|
|Office.GoToType.Index|"index"|スライド番号または次の列挙型を使用して指定されたインデックスに移動します。</br>**Office.Index.First**</br>**Office.Index.Last**</br>**Office.Index.Next**</br>**Office.Index.Previous**|PowerPoint|

## <a name="support-details"></a>サポートの詳細


次の表で、大文字 Y は、この列挙は、対応する Office ホスト アプリケーションでサポートされていることを示します。空のセルは、Office ホスト アプリケーションがこの列挙をサポートしないことを示します。


Office ホスト アプリケーションとサーバーの要件の詳細については、「[Office アドインを実行するための要件](../../docs/overview/requirements-for-running-office-add-ins.md)」をご覧ください。


**サポートされるホスト (プラットフォーム別)**


||**Windows デスクトップ版 Office**|**Office Online (ブラウザー)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|Y|Y|Y|
|**PowerPoint**|Y|Y|Y|
|**Word**|Y||Y|

|||
|:-----|:-----|
|**アドインの種類**|コンテンツ、作業ウィンドウ|
|**ライブラリ**|Office.js|
|**名前空間**|Office|

## <a name="support-history"></a>サポート履歴




|**バージョン**|**変更内容**|
|:-----|:-----|
|1.1|Office for iPad で Excel、PowerPoint、および Word のサポートが追加されました。|
|1.1|導入|
