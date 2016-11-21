# <a name="office-common-api-requirement-sets"></a>Office 共通 API の要件セット

要件セットは、API メンバーの名前付きグループです。Office アドインは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office ホストがアドインに必要な API をサポートしているかどうかを判別します。詳しくは、「[Office のホストと API の要件を指定する](../docs/overview/specify-office-hosts-and-api-requirements.md)」をご覧ください。

Office ホストでのアドインのサポート状況について詳しくは、「[Office アドインを使用できるホストおよびプラットフォーム](https://dev.office.com/add-in-availability)」をご覧ください。

## <a name="hostspecific-api-requirement-sets"></a>ホスト固有の API 要件セットについて

Excel、Word、OneNote、Outlook、ダイアログ API の要件セットについて詳しくは、次をご覧ください。

- [Excel JavaScript API の要件セット](excel-api-requirement-sets.md)
- [Word JavaScript API の要件セット](word-api-requirement-sets.md)
- [OneNote JavaScript API の要件セット](onenote-api-requirement-sets.md)
- [Outlook API 要件セットについて](../outlook/tutorial-api-requirement-sets.md)
[ダイアログ API の要件セット](dialog-api-requirement-sets.md)

## <a name="common-api-requirement-sets"></a>共通 API の要件セット

次の表は、共通 API の要件セット、各セットのメソッド、その要件セットをサポートする Office ホスト アプリケーションの一覧です。これらの API 要件セットのバージョンはすべて 1.1 です。


|  要件セット  |  Office ホスト  |  セット内のメソッド  |
|:-----|-----|:-----|:-----|
| ActiveView | PowerPoint<br>PowerPoint&nbsp;Online|Document.getActiveViewAsync|
| BindingEvents  | Access Web App<br>Excel<br>Excel Online<br>Word 2013 以降<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|Binding.addHanderAsync<br>Binding.removeHanderAsync|
| CompressedFile    | PowerPoint<br>Word 2013 以降<br>Word 2016 for Mac<br>Word Online<br>Word for iPad<br/>Excel Online<br/>PowerPoint Online|Document.getFileAsync メソッドを使用するときの、<br>バイト配列 (Office.FileType.Compressed) としての Office Open XML (OOXML) 形式への出力をサポートします。|
| CustomXmlParts    | Word 2013 以降<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|CustomXmlNode.getNodesAsync<br>CustomXmlNode.getNodeValueAsync<br>CustomXmlNode.getXmlAsync<br>CustomXmlNode.setNodeValueAsync<br>CustomXmlNode.setXmlAsync<br>CustomXmlPart.addHandlerAsync<br>CustomXmlPart.deleteAsync<br>CustomXmlPart.getNodesAsync<br>CustomXmlPart.getXmlAsync<br>CustomXmlPart.removeHandlerAsync<br>CustomXmlParts.addAsync<br>CustomXmlParts.getByIdAsync<br>CustomXmlParts.getByNamespaceAsync<br>CustomXmlPrefixMappings.addNamespaceAsync<br>CustomXmlPrefixMappings.getNamespaceAsync<br>CustomXmlPrefixMappings.getPrefixAsync|
| DocumentEvents    | Excel<br>Excel Online<br>PowerPoint Online<br>Word 2013 以降<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|Document.addHandlerAsync<br>Document.removeHandlerAsync|
| File  | PowerPoint<br>Word 2013 以降<br>Word 2016 for Mac<br>Word Online<br>Word for iPad<br>PowerPoint Online|Document.getFileAsync<br>File.closeAsync<br>File.getSliceAsync|
| HtmlCoercion  | Word 2013 以降<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|Document.getSelectedDataAsync、Document.setSelectedDataAsync、Binding.getDataAsync、または Binding.setDataAsync メソッドを使用してデータを読み書きするときの、<br>HTML (Office.CoercionType.Html) への強制型変換をサポートします。|
| ImageCoercion | Word 2013 以降<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|Document.setSelectedDataAsync メソッドを使用してデータを書き込むときに、画像 (Office.CoercionType.Image) への変換をサポートしています。|
| メールボックス   |Outlook for Windows<br>Outlook for web<br>Outlook for Mac<br>Outlook Web App |「[Outlook API 要件セットについて](./outlook/tutorial-api-requirement-sets.md)」をご覧ください。|
| MatrixBindings    | Excel<br>Excel Online<br>Word<br>Word Online|Bindings.addFromNamedItemAsync<br>Bindings.addFromSelectionAsync<br>Bindings.getAllAsync<br>Bindings.getByIdAsync<br>Bindings.releaseByIdAsyncMatrix<br>Binding.getDataAsyncMatrix<br>Binding.setDataAsync|
| MatrixCoercion    | Excel<br>Excel Online<br>Word 2013 以降<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|Document.getSelectedDataAsync、Document.setSelectedDataAsync、Binding.getDataAsync、または Binding.setDataAsync の各メソッドを使用してデータを読み書きするとき、"matrix" (配列の配列) データ構造への強制型変換 (Office.CoercionType.Matrix) をサポートします。|
| OoxmlCoercion | Word 2013 以降<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|Document.getSelectedDataAsync、Document.setSelectedDataAsync、Binding.getDataAsync、または Binding.setDataAsync の各メソッドを使用してデータを読み書きするとき、Open Office XML (OOXML) 形式への強制型変換 (Office.CoercionType.Ooxml) をサポートします。|
| PartialTableBindings  | Access Web App||
| PdfFile   | PowerPoint<br/>PowerPoint Online<br/>Word 2013 以降<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|Document.getFileAsync メソッドを使用するときの、<br>PDF 形式 (Office.FileType.Pdf) への出力をサポートします。|
| Selection | Excel<br>Excel Online<br>PowerPoint<br>Project<br>Word 2013 以降<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|Document.getSelectedDataAsync<br>Document.setSelectedDataAsync|
| Settings  | Access Web App<br>Excel<br>Excel Online<br>PowerPoint<br>PowerPoint Online<br>Word 2013 以降<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|Settings.get<br>Settings.remove<br>Settings.saveAsync<br>Settings.set|
| TableBindings | Access Web App<br>Excel<br>Excel Online<br>Word 2013 以降<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|Bindings.addFromNamedItemAsync<br>Bindings.addFromSelectionAsync<br>Bindings.getAllAsync<br>Bindings.getByIdAsync<br>Bindings.releaseByIdAsyncTable<br>Binding.addColumnsAsyncTable<br>Binding.addRowsAsyncTable<br>Binding.deleteAllDataValuesAsyncTable<br>Binding.getDataAsyncTable<br>Binding.setDataAsync|
| TableCoercion | Access Web App<br>Excel<br>Excel Online<br>Word 2013 以降<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|Document.getSelectedDataAsync、Document.setSelectedDataAsync、Binding.getDataAsync、または Binding.setDataAsync の各メソッドを使用してデータを読み書きするとき、"table" データ構造への強制型変換 (Office.CoercionType.Table) をサポートします。|
| TextBindings  | Excel<br>Excel Online<br>Word 2013 以降<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|Bindings.addFromNamedItemAsync<br>Bindings.addFromSelectionAsync<br>Bindings.getAllAsync<br>Bindings.getByIdAsync<br>Bindings.releaseByIdAsyncText<br>Binding.getDataAsyncText<br>Binding.setDataAsync|
| TextCoercion  | Excel<br>Excel Online<br>PowerPoint<br>Project<br>Word 2013 以降<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|Document.getSelectedDataAsync、Document.setSelectedDataAsync、Binding.getDataAsync、または Binding.setDataAsync の各メソッドを使用してデータを読み書きするとき、テキスト形式への強制型変換 (Office.CoercionType.Text) をサポートします。|
| TextFile  | Word 2013 以降<br>Word 2016 for Mac<br>Word Online<br>Word for iPad<br/>|Document.getFileAsync メソッドを使用するとき、テキスト形式 (Office.FileType.Text) への出力をサポートします。|

## <a name="methods-that-arent-part-of-a-requirement-set"></a>要件セットの一部ではないメソッド

JavaScript API for Office の以下のメソッドは、要件セットの一部ではありません。アドインでこれらのメソッドが必要な場合は、アドインのマニフェストで **Methods** 要素と **Method** 要素を使用してメソッドが必要であると宣言するか、または if ステートメントを使用してランタイム チェックを実行します。詳細については、「[Office のホストと API 要件を指定する](../docs/overview/specify-office-hosts-and-api-requirements.md)」をご覧ください。

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

## <a name="additional-resources"></a>その他のリソース

- [Office のホストと API の要件を指定する](../docs/overview/specify-office-hosts-and-api-requirements.md)



