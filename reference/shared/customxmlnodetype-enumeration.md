
# CustomXMLNodeType 列挙型
ノードの種類を指定します。



|||
|:-----|:-----|
|**ホスト:**|Word|
|**最終変更バージョン**|1.1|



```js
Office.CustomXMLNodeType
```


## メンバー


**値**


|**列挙体**|**値**|**説明**|
|:-----|:-----|:-----|
|Office.CustomXMLNodeType.Attribute|"attribute"|ノードは属性です。|
|Office.CustomXMLNodeType.CData|"CData"|ノードは CData 型です。|
|Office.CustomXMLNodeType.NodeComment|"comment"|ノードはコメントです。|
|Office.CustomXMLNodeType.Element|"element"|ノードは要素です。|
|Office.CustomXMLNodeType.NodeDocument|"nodeDocument"|ノードはドキュメント要素です。|
|Office.CustomXMLNodeType.ProcessingInstruction|"processingInstruction"|ノードは処理命令です。|
|Office.CustomXMLNodeType.Text|"text"|ノードはテキスト ノードです。|

## サポートの詳細


次の表で、大文字 Y は、この列挙は、対応する Office ホスト アプリケーションでサポートされていることを示します。空のセルは、Office ホスト アプリケーションがこの列挙をサポートしないことを示します。

Office ホスト アプリケーションとサーバーの要件の詳細については、「[Office アドインを実行するための要件](../../docs/overview/requirements-for-running-office-add-ins.md)」をご覧ください。


**サポートされるホスト (プラットフォーム別)**


||**Windows デスクトップ版 Office**|**Office Online (ブラウザー)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Word**|Y||Y|



|||
|:-----|:-----|
|**アプリの種類**|作業ウィンドウ|
|**ライブラリ**|Office.js|
|**名前空間**|Office|

## サポート履歴


|**変更内容**|**1.1**|
|:-----|:-----|
|1.1|Office for iPad における Word のサポートが追加されました。|
|1.0|導入|
