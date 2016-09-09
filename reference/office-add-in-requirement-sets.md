
# Office アドインの要件セット

要件セットは、API メンバーの名前付きグループです。 Office アドインは、マニフェストで指定されている要件セットを使用するか、実行時チェックを使用して、Office ホストがアドインに必要な API をサポートしているかどうかを判別します。 詳細については、「[Office ホストと API 要件を指定する](../docs/overview/specify-office-hosts-and-api-requirements.md)」をご覧ください。

Office ホストでのアドインのサポート状況の概要については、「[Office アドインを使用できるホストおよびプラットフォーム](https://dev.office.com/add-in-availability)」ページを参照してください。

## 要件セット


次の表に、要件セットの名前、各セットのメソッド、要件セットをサポートする Office ホスト アプリケーション、API のバージョン番号を示します。

Outlook の要件セットについては、「[Outlook API 要件セットについて](./outlook/tutorial-api-requirement-sets.md)」を参照してください。

|  セットの名前  |  バージョン  |  Office ホスト  |  セット内のメソッド  |
|:-----|-----|:-----|:-----|
| ExcelApi   | 1.2 | Excel 2016<br>Excel Online<br>Excel for iPad<br>|ワークシート保護<br>ワークシート関数<br>並べ替え<br>フィルター<br>R1C1 参照スタイル<br>セルの結合<br>行の高さと列の幅の調整<br>Chart.getImage()<br>Range.getUsedRange(valuesOnly)|
| ExcelApi   | 1.1 | Excel 2016<br>Excel Online<br>Excel for iPad<br>|Excel 名前空間内のすべての要素|
| WordApi    | 1.2 | Word 2016<br>Word 2016 for Mac<br>Word for iPad<br>Word Online (プレビュー) | Word 名前空間内のすべての要素。 WordApi のこのバージョンには次のメソッドが追加されました。<br>Body.select(selectionMode)<br>Body.insertInlinePictureFromBase64(base64EncodedImage, insertLocation)<br>contentControl.select(selectionMode)<br>contentControl.insertInlinePictureFromBase64(base64EncodedImage, insertLocation)<br>inlinePicture.paragraph<br>inlinePicture.delete<br>inlinePicture.insertBreak(breakType, insertLocation)<br>inlinePicture.insertFileFromBase64(base64file, insertLocation)<br>inlinePicture.insertHtml(html, insertLocation)<br>inlinePicture.insertInlinePictureFromBase64(base64file, insertLocation)<br>inlinePicture.insertOoxml(ooxml, insertLocation)<br>inlinePicture.insertParagraph(paragraphText, insertLocation)<br>inlinePicture.insertText(text, insertLocation)<br>inlinePicture.select(selectionMode)<br>paragraph.select(selectionMode)<br>range.inlinePictures<br>range.select(selectionMode)<br>range.insertInlinePictureFomBase64(base64EcodedImage, insertLocation)|
| WordApi    | 1.1 | Word 2016<br>Word 2016 for Mac<br>Word for iPad<br>|WordApi 1.2 以降に追加された API メンバー (上に記載) 以外のすべての Word 名前空間の要素。|
| ActiveView | 1.1 | PowerPoint<br>PowerPoint Online|Document.getActiveViewAsync|
| BindingEvents  | 1.1 | Access Web アプリ<br>Excel<br>Excel Online<br>Word 2013 以降<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|Binding.addHanderAsync<br>Binding.removeHanderAsync|
| CompressedFile    | 1.1 |PowerPoint<br>Word 2013 以降<br>Word 2016 for Mac<br>Word Online<br>Word for iPad<br/>Excel Online<br/>PowerPoint Online|Document.getFileAsync メソッドを使用するときの、<br>バイト配列 (Office.FileType.Compressed) としての Office Open XML (OOXML) 形式への出力をサポートします。|
| CustomXmlParts    | 1.1 |Word 2013 以降<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|CustomXmlNode.getNodesAsync<br>CustomXmlNode.getNodeValueAsync<br>CustomXmlNode.getXmlAsync<br>CustomXmlNode.setNodeValueAsync<br>CustomXmlNode.setXmlAsync<br>CustomXmlPart.addHandlerAsync<br>CustomXmlPart.deleteAsync<br>CustomXmlPart.getNodesAsync<br>CustomXmlPart.getXmlAsync<br>CustomXmlPart.removeHandlerAsync<br>CustomXmlParts.addAsync<br>CustomXmlParts.getByIdAsync<br>CustomXmlParts.getByNamespaceAsync<br>CustomXmlPrefixMappings.addNamespaceAsync<br>CustomXmlPrefixMappings.getNamespaceAsync<br>CustomXmlPrefixMappings.getPrefixAsync|
| DialogAPI | 1.1 | Excel<br>PowerPoint<br>Word 2016<br>Outlook|Office.context.ui.displayDialogAsync()<br>Office.context.ui.messageParent()<br>Office.context.ui.close()|
| DocumentEvents    | 1.1 | Excel<br>Excel Online<br>PowerPoint<br>Word 2013 以降<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|Document.addHandlerAsync<br>Document.removeHandlerAsync|
| File  | 1.1 | PowerPoint<br>Word 2013 以降<br>Word 2016 for Mac<br>Word Online<br>Word for iPad<br>PowerPoint Online|Document.getFileAsync<br>File.closeAsync<br>File.getSliceAsync|
| HtmlCoercion  | 1.1 | Word 2013 以降<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|Document.getSelectedDataAsync、Document.setSelectedDataAsync、Binding.getDataAsync、または Binding.setDataAsync メソッドを使用してデータを読み書きするときの、<br>HTML (Office.CoercionType.Html) への強制型変換をサポートします。|
| ImageCoercion | 1.1 | Word 2013 以降<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|Document.setSelectedDataAsync メソッドを使用してデータを書き込むときに、画像 (Office.CoercionType.Image) への変換をサポートしています。|
| メールボックス   |   | Outlook for Windows<br>Outlook for web<br>Outlook for Mac<br>Outlook Web App |「[Outlook API 要件セットについて](./outlook/tutorial-api-requirement-sets.md)」を参照|
| MatrixBindings    | 1.1 | Excel<br>Excel Online<br>Word|Bindings.addFromNamedItemAsync<br>Bindings.addFromSelectionAsync<br>Bindings.getAllAsync<br>Bindings.getByIdAsync<br>Bindings.releaseByIdAsyncMatrix<br>Binding.getDataAsyncMatrix<br>Binding.setDataAsync|
| MatrixCoercion    | 1.1 | Excel<br>Excel Online<br>Word 2013 以降<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|Document.getSelectedDataAsync、Document.setSelectedDataAsync、Binding.getDataAsync、または Binding.setDataAsync の各メソッドを使用してデータを読み書きするとき、"matrix" (配列の配列) データ構造への強制型変換 (Office.CoercionType.Matrix) をサポートします。|
| OoxmlCoercion | 1.1 | Word 2013 以降<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|Document.getSelectedDataAsync、Document.setSelectedDataAsync、Binding.getDataAsync、または Binding.setDataAsync の各メソッドを使用してデータを読み書きするとき、Open Office XML (OOXML) 形式への強制型変換 (Office.CoercionType.Ooxml) をサポートします。|
| PartialTableBindings  | 1.1 | Access Web アプリ||
| PdfFile   | 1.1 | PowerPoint<br/>PowerPoint Online<br/>Word 2013 以降<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|Document.getFileAsync メソッドを使用するときの、<br>PDF 形式 (Office.FileType.Pdf) への出力をサポートします。|
| Selection | 1.1 | Excel<br>Excel Online<br>PowerPoint<br>Project<br>Word 2013 以降<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|Document.getSelectedDataAsync<br>Document.setSelectedDataAsync|
| Settings  | 1.1 | Access Web アプリ<br>Excel<br>Excel Online<br>PowerPoint<br>PowerPoint Online<br>Word 2013 以降<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|Settings.get<br>Settings.remove<br>Settings.saveAsync<br>Settings.set|
| TableBindings | 1.1 | Access Web アプリ<br>Excel<br>Excel Online<br>Word 2013 以降<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|Bindings.addFromNamedItemAsync<br>Bindings.addFromSelectionAsync<br>Bindings.getAllAsync<br>Bindings.getByIdAsync<br>Bindings.releaseByIdAsyncTable<br>Binding.addColumnsAsyncTable<br>Binding.addRowsAsyncTable<br>Binding.deleteAllDataValuesAsyncTable<br>Binding.getDataAsyncTable<br>Binding.setDataAsync|
| TableCoercion | 1.1 | Access Web アプリ<br>Excel<br>Excel Online<br>Word 2013 以降<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|Document.getSelectedDataAsync、Document.setSelectedDataAsync、Binding.getDataAsync、または Binding.setDataAsync の各メソッドを使用してデータを読み書きするとき、"table" データ構造への強制型変換 (Office.CoercionType.Table) をサポートします。|
| TextBindings  | 1.1 | Excel<br>Excel Online<br>Word 2013 以降<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|Bindings.addFromNamedItemAsync<br>Bindings.addFromSelectionAsync<br>Bindings.getAllAsync<br>Bindings.getByIdAsync<br>Bindings.releaseByIdAsyncText<br>Binding.getDataAsyncText<br>Binding.setDataAsync|
| TextCoercion  | 1.1 | Excel<br>Excel Online<br>PowerPoint<br>Project<br>Word 2013 以降<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|Document.getSelectedDataAsync、Document.setSelectedDataAsync、Binding.getDataAsync、または Binding.setDataAsync の各メソッドを使用してデータを読み書きするとき、テキスト形式への強制型変換 (Office.CoercionType.Text) をサポートします。|
| TextFile  | 1.1 | Word 2013 以降<br>Word 2016 for Mac<br>Word Online<br>Word for iPad<br/>|Document.getFileAsync メソッドを使用するとき、テキスト形式 (Office.FileType.Text) への出力をサポートします。|

## 要件セットの一部ではないメソッド


JavaScript API for Office の以下のメソッドは、要件セットの一部ではありません。 アドインでこれらのメソッドが必要な場合は、アドインのマニフェストで **Methods** 要素と **Method** 要素を使用してメソッドが必要であると宣言するか、または if ステートメントを使用してランタイム チェックを実行します。 詳細については、「[Office ホストと API 要件を指定する](../docs/overview/specify-office-hosts-and-api-requirements.md)」をご覧ください。



|**メソッド名**|**サポートされる Office のホスト**|
|:-----|:-----|
|Bindings.addFromPromptAsync|Access web アプリ、Excel、Excel Online|
|Document.getFilePropertiesAsync|Excel、Excel Online、Word、PowerPoint|
|Document.getProjectFieldAsync|Project Standard 2013、Project Professional 2013|
|Document.getResourceFieldAsync|Project Standard 2013、Project Professional 2013|
|Document.getSelectedResourceAsync|Project Standard 2013、Project Professional 2013|
|Document.getSelectedTaskAsync|Project Standard 2013、Project Professional 2013|
|Document.getSelectedViewAsync|PowerPoint、PowerPoint Online|
|Document.getTaskAsync|Project Standard 2013、Project Professional 2013|
|Document.getTaskFieldAsync|Project Standard 2013、Project Professional 2013|
|Document.goToByIdAsync|Excel、Excel Online、Word、PowerPoint|
|Settings.addHandlerAsync|Access web アプリ、Excel、Excel Online、Word、PowerPoint|
|Settings.refreshAsync|Access web アプリ、Excel、Excel Online、Word、PowerPoint、PowerPoint Online|
|Settings.removeHandlerAsync|Access web アプリ、Excel、Excel Online、Word、PowerPoint|
|TableBinding.clearFormatsAsync|Excel、Excel Online|
|TableBinding.setFormatsAsync|Excel、Excel Online|
|TableBinding.setTableOptionsAsync|Excel、Excel Online|

## その他のリソース



- [Office のホストと API の要件を指定する](../docs/overview/specify-office-hosts-and-api-requirements.md)

