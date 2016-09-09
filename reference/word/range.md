# Range オブジェクト (JavaScript API for Word)

文書内の連続した領域を表します。

_適用対象:Word 2016、Word for iPad、Word for Mac_

## プロパティ
| プロパティ     | 型   |説明
|:---------------|:--------|:----------|
|style|string|範囲に使用されるスタイルを取得または設定します。これは、事前にインストールされているスタイルまたはユーザー設定のスタイルの名前です。|
|text|string|範囲のテキストを取得します。読み取り専用です。|

## リレーションシップ
| リレーションシップ | 型   |説明|
|:---------------|:--------|:----------|
|contentControls|[ContentControlCollection](contentcontrolcollection.md)|範囲に含まれるコンテンツ コントロール オブジェクトのコレクションを取得します。読み取り専用です。|
|Font|[フォント](font.md)|範囲のテキスト形式を取得します。これを使用して、フォント名、サイズ、色、およびその他のプロパティを取得および設定します。読み取り専用です。|
|inlinePictures|[InlinePictureCollection](inlinepicturecollection.md)|範囲に含まれる inlinePicture オブジェクトのコレクションを取得します。読み取り専用です。|
|paragraphs|[ParagraphCollection](paragraphcollection.md)|範囲に含まれる段落オブジェクトのコレクションを取得します。読み取り専用です。|
|parentContentControl|[ContentControl](contentcontrol.md)|範囲を格納するコンテンツ コントロールを取得します。親コンテンツ コントロールがない場合は、null を返します。読み取り専用です。|

## メソッド

| メソッド           | 戻り値の型    |説明|
|:---------------|:--------|:----------|
|[clear()](#clear)|void|範囲オブジェクトの内容をクリアします。ユーザーは、クリアしたコンテンツを元に戻す操作を実行できます。|
|[delete()](#delete)|void|文書から範囲と、その範囲の内容を削除します。|
|[getHtml()](#gethtml)|string|範囲オブジェクトの HTML 表記を取得します。|
|[getOoxml()](#getooxml)|string|Range オブジェクトの OOXML 表記を取得します。|
|[insertBreak(breakType: BreakType, insertLocation: InsertLocation)](#insertbreakbreaktype-breaktype-insertlocation-insertlocation)|void|指定した位置に、区切りを挿入します。 改行以外の区切りは、メイン文書本文内に含まれた範囲オブジェクトにのみ挿入できます。改行はどの本文オブジェクトにも挿入できます。 有効な insertLocation の値は、'Before' または 'After' です。|
|[insertContentControl()](#insertcontentcontrol)|[ContentControl](contentcontrol.md)|範囲オブジェクトを、リッチ テキストのコンテンツ コントロールでラップします。|
|[insertFileFromBase64(base64File: string, insertLocation:InsertLocation)](#insertfilefrombase64base64file-string-insertlocation-insertlocation)|[範囲](range.md)|範囲の指定した位置に文書を挿入します。insertLocation の値には、'Replace'、'Start'、'End' のいずれかを指定できます。|
|[insertHtml(html: string, insertLocation:InsertLocation)](#inserthtmlhtml-string-insertlocation-insertlocation)|[範囲](range.md)|範囲の指定した位置に HTML を挿入します。insertLocation の値には、'Replace'、'Start'、'End' のいずれかを指定できます。|
|[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: InsertLocation)](#insertinlinepicturefrombase64base64encodedimage-string-insertlocation-insertlocation)|[InlinePicture](inlinepicture.md)|画像を範囲の指定された位置に挿入します。insertLocation の値は、'Replace'、'Start'、'End'、'Before' 、'After' のいずれかになります。
|[insertOoxml(ooxml: string, insertLocation: InsertLocation)](#insertooxmlooxml-string-insertlocation-insertlocation)|[範囲](range.md)|範囲内の指定された位置に OOXML または wordProcessingML を挿入します。insertLocation の値には、'Replace'、'Start'、'End' のいずれかを指定できます。|
|[insertParagraph(paragraphText: string, insertLocation: InsertLocation)](#insertparagraphparagraphtext-string-insertlocation-insertlocation)|[段落](paragraph.md)|範囲の指定した位置に段落を挿入します。有効な insertLocation の値は、'Before' または 'After' です。|
|[insertText(text: string, insertLocation:InsertLocation)](#inserttexttext-string-insertlocation-insertlocation)|[範囲](range.md)|範囲の指定した位置にテキストを挿入します。insertLocation の値には、'Replace'、'Start'、'End' のいずれかを指定できます。|
|[load(param: object)](#loadparam-object)|void|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|
|[search(searchText: string, searchOptions:ParamTypeStrings.SearchOptions)](#searchsearchtext-string-searchoptions-paramtypestrings.searchoptions)|[SearchResultCollection](searchresultcollection.md)|範囲オブジェクトの範囲で、searchOptions を指定した検索を実行します。検索結果は、範囲オブジェクトのコレクションになります。|
|[select(selectionMode: SelectionMode)](#selectselectionmode-selectionmode)|void|範囲を選択して、その範囲に Word の UI を移動します。selectionMode 値は、'Select'、'Start'、'End' のいずれかになります。|

## メソッドの詳細

### clear()
範囲オブジェクトの内容をクリアします。ユーザーは、クリアしたコンテンツを元に戻す操作を実行できます。

#### 構文
```js
rangeObject.clear();
```

#### パラメーター
なし

#### 戻り値
void

#### 例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Queue a command to get the current selection and then
    // create a proxy range object with the results.
    var range = context.document.getSelection();

    // Queue a commmand to clear the contents of the proxy range object.
    range.clear();

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Cleared the selection (range object)');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```
### delete()
文書から範囲と、その範囲の内容を削除します。

#### 構文
```js
rangeObject.delete();
```

#### パラメーター
なし

#### 戻り値
void

#### 例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Queue a command to get the current selection and then
    // create a proxy range object with the results.
    var range = context.document.getSelection();

    // Queue a commmand to delete the range object.
    range.delete();

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Deleted the selection (range object)');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### getHtml()
範囲オブジェクトの HTML 表記を取得します。

#### 構文
```js
rangeObject.getHtml();
```

#### パラメーター
なし

#### 戻り値
string

#### 例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Queue a command to get the current selection and then
    // create a proxy range object with the results.
    var range = context.document.getSelection();

    // Queue a commmand to get the HTML of the current selection.
    var html = range.getHtml();

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('The HTML read from the document was: ' + html.value);
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### getOoxml()
Range オブジェクトの OOXML 表記を取得します。

#### 構文
```js
rangeObject.getOoxml();
```

#### パラメーター
なし

#### 戻り値
string

#### 例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Queue a command to get the current selection and then
    // create a proxy range object with the results.
    var range = context.document.getSelection();

    // Queue a commmand to get the OOXML of the current selection.
    var ooxml = range.getOoxml();

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('The OOXML read from the document was:  ' + ooxml.value);
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### insertBreak(breakType: BreakType, insertLocation: InsertLocation)
指定した位置に、区切りを挿入します。 改行以外の区切りは、メイン文書本文内に含まれた範囲オブジェクトにのみ挿入できます。改行はどの本文オブジェクトにも挿入できます。 有効な insertLocation の値は、'Before' または 'After' です。

#### 構文
```js
rangeObject.insertBreak(breakType, insertLocation);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|breakType|BreakType|必須。範囲に追加する区切りの種類。|
|insertLocation|InsertLocation|必須。有効な値は、'Before' または 'After' です。|

#### 戻り値
void

#### 追加の詳細
ヘッダー、フッター、脚注、文末脚注、コメント、テキスト ボックスのオブジェクトには改行以外の区切りを挿入できません。

#### 例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Queue a command to get the current selection and then
    // create a proxy range object with the results.
    var range = context.document.getSelection();

    // Queue a commmand to insert a page break after the selected text.
    range.insertBreak('page', 'After');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Inserted a page break after the selected text.');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### insertContentControl()
範囲オブジェクトを、リッチ テキストのコンテンツ コントロールでラップします。

#### 構文
```js
rangeObject.insertContentControl();
```

#### パラメーター
なし

#### 戻り値
[ContentControl](contentcontrol.md)

#### 例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Queue a command to get the current selection and then
    // create a proxy range object with the results.
    var range = context.document.getSelection();

    // Queue a commmand to insert a content control around the selected text,
    // and create a proxy content control object. We'll update the properties
    // on the content control.
    var myContentControl = range.insertContentControl();
    myContentControl.tag = "Customer-Address";
    myContentControl.title = "Enter Customer Address Here:";
    myContentControl.style = "Normal";
    myContentControl.insertText("One Microsoft Way, Redmond, WA 98052", 'replace');
    myContentControl.cannotEdit = true;

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Wrapped a content control around the selected text.');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### insertFileFromBase64(base64File: string, insertLocation:InsertLocation)
範囲の指定した位置に文書を挿入します。insertLocation の値には、'Replace'、'Start'、'End' のいずれかを指定できます。

#### 構文
```js
rangeObject.insertFileFromBase64(base64File, insertLocation);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|base64File|string|必須。挿入するファイルの内容が base64 エンコードされているファイル。|
|insertLocation|InsertLocation|必須。値には、'Replace'、'Start'、'End' のいずれかを指定できます。|

#### 戻り値
[範囲](range.md)

#### 例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Queue a command to get the current selection and then
    // create a proxy range object with the results.
    var range = context.document.getSelection();

    // Queue a commmand to insert base64 encoded .docx at the beginning of the range.
    // You'll need to implement getBase64() to make this work.
    range.insertFileFromBase64(getBase64(), Word.InsertLocation.start);

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Added base64 encoded text to the beginning of the range.');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### insertHtml(html: string, insertLocation:InsertLocation)
範囲の指定した位置に HTML を挿入します。insertLocation の値には、'Replace'、'Start'、'End' のいずれかを指定できます。

#### 構文
```js
rangeObject.insertHtml(html, insertLocation);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|Html|string|必須。範囲に挿入する HTML。|
|insertLocation|InsertLocation|必須。値には、'Replace'、'Start'、'End' のいずれかを指定できます。|

#### 戻り値
[範囲](range.md)

#### 例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Queue a command to get the current selection and then
    // create a proxy range object with the results.
    var range = context.document.getSelection();

    // Queue a commmand to insert HTML in to the beginning of the range.
    range.insertHtml('<strong>This is text inserted with range.insertHtml()</strong>', Word.InsertLocation.start);

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('HTML added to the beginning of the range.');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: InsertLocation)
画像を範囲の指定された位置に挿入します。insertLocation の値は、'Replace'、'Start'、'End'、'Before' 、'After' のいずれかになります。

#### 構文
rangeObject.insertInlinePictureFromBase64(image, insertLocation);

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|base64EncodedImage|string|必須。範囲に挿入される base64 でエンコードされた画像。|
|insertLocation|InsertLocation|必須。値は、'Replace'、'Start'、'End'、'Before' 、'After' のいずれかになります。|

#### 戻り値
[InlinePicture](inlinepicture.md)

### insertOoxml(ooxml: string, insertLocation: InsertLocation)
範囲内の指定された位置に OOXML または wordProcessingML を挿入します。insertLocation の値には、'Replace'、'Start'、'End' のいずれかを指定できます。

#### 構文
```js
rangeObject.insertOoxml(ooxml, insertLocation);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|ooxml|string|必須。範囲に挿入する OOXML または wordProcessingML。|
|insertLocation|InsertLocation|必須。値には、'Replace'、'Start'、'End' のいずれかを指定できます。|

#### 戻り値
[範囲](range.md)

#### 例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Queue a command to get the current selection and then
    // create a proxy range object with the results.
    var range = context.document.getSelection();

    // Queue a commmand to insert OOXML in to the beginning of the range.
    range.insertOoxml("<pkg:package xmlns:pkg='http://schemas.microsoft.com/office/2006/xmlPackage'><pkg:part pkg:name='/_rels/.rels' pkg:contentType='application/vnd.openxmlformats-package.relationships+xml' pkg:padding='512'><pkg:xmlData><Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument' Target='word/document.xml'/></Relationships></pkg:xmlData></pkg:part><pkg:part pkg:name='/word/document.xml' pkg:contentType='application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'><pkg:xmlData><w:document xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' ><w:body><w:p><w:pPr><w:spacing w:before='360' w:after='0' w:line='480' w:lineRule='auto'/><w:rPr><w:color w:val='70AD47' w:themeColor='accent6'/><w:sz w:val='28'/></w:rPr></w:pPr><w:r><w:rPr><w:color w:val='70AD47' w:themeColor='accent6'/><w:sz w:val='28'/></w:rPr><w:t>This text has formatting directly applied to achieve its font size, color, line spacing, and paragraph spacing.</w:t></w:r></w:p></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>", Word.InsertLocation.start);

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('OOXML added to the beginning of the range.');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

#### 追加情報
OOXML の操作の詳細については、「[Office Open XML を使用して Word のより良いアドインを作成する](https://msdn.microsoft.com/en-us/library/office/dn423225.aspx)」をお読みください。

### insertParagraph(paragraphText: string, insertLocation: InsertLocation)
範囲の指定した位置に段落を挿入します。有効な insertLocation の値は、'Before' または 'After' です。

#### 構文
```js
rangeObject.insertParagraph(paragraphText, insertLocation);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|paragraphText|string|必須。挿入する段落テキスト。|
|insertLocation|InsertLocation|必須。有効な値は、'Before' または 'After' です。|

#### 戻り値
[段落](paragraph.md)

#### 例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Queue a command to get the current selection and then
    // create a proxy range object with the results.
    var range = context.document.getSelection();

    // Queue a commmand to insert the paragraph after the range.
    range.insertParagraph('Content of a new paragraph', Word.InsertLocation.after);

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Paragraph added to the end of the range.');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### insertText(text: string, insertLocation:InsertLocation)
範囲の指定した位置にテキストを挿入します。insertLocation の値には、'Replace'、'Start'、'End' のいずれかを指定できます。

#### 構文
```js
rangeObject.insertText(text, insertLocation);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|text|string|必須。挿入するテキスト。|
|insertLocation|InsertLocation|必須。値には、'Replace'、'Start'、'End' のいずれかを指定できます。|

#### 戻り値
[範囲](range.md)

#### 例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Queue a command to get the current selection and then
    // create a proxy range object with the results.
    var range = context.document.getSelection();

    // Queue a commmand to insert the paragraph at the end of the range.
    range.insertText('New text inserted into the range.', Word.InsertLocation.end);

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Text added to the end of the range.');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### load(param: object)
JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。

#### 構文
```js
object.load(param);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|param|object|省略可能。パラメーターとリレーションシップ名を、区切られた文字列または 1 つの配列として受け入れます。あるいは、[loadOption](loadoption.md) オブジェクトを提供します。|

#### 戻り値
void

#### 例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Queue a command to get the current selection and then
    // create a proxy range object with the results.
    var range = context.document.getSelection();

    // Queue a commmand to load font and style information for the range.
    context.load(range, 'font/size, font/name, font/color, style');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {

        // Show the results of the load method. Here we show the
        // property values on the range object.
        var results = "  ---Font size: " + range.font.size +
                      "  ---Font name: " + range.font.name +
                      "  ---Font color: " + range.font.color +
                      "  ---Style: " + range.style;
        console.log(results);
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### search(searchText: string, searchOptions:ParamTypeStrings.SearchOptions)
範囲オブジェクトの範囲で、searchOptions を指定した検索を実行します。検索結果は、範囲オブジェクトのコレクションになります。

#### 構文
```js
rangeObject.search(searchText, searchOptions);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|searchText|string|必須。検索テキスト。|
|[searchOptions](searchoptions.md)|ParamTypeStrings.SearchOptions|省略可能。検索のオプション。|

#### 戻り値
[SearchResultCollection](searchresultcollection.md)


### select(selectionMode: SelectionMode)
範囲を選択して、その範囲に Word の UI を移動します。selectionMode 値は、'Select'、'Start'、'End' のいずれかになります。

#### 構文
```js
rangeObject.select(selectionMode);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|selectionMode|SelectionMode|省略可能。選択モードは、'Select'、'Start'、'End' のいずれかになります。'Select' が既定値です。|

#### 戻り値
void

#### 例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Queue a command to get the current selection and then
    // create a proxy range object with the results.
    var range = context.document.getSelection();

    // Queue a commmand to insert HTML in to the beginning of the range.
    range.insertHtml('<strong>This is text inserted with range.insertHtml()</strong>', Word.InsertLocation.start);

    // Queue a command to select the HTML that was inserted.
    range.select();

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Selected the range.');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

## サポートの詳細
実行時のチェックで[要件セット](../office-add-in-requirement-sets.md)を使用して、アプリケーションが Word のホスト バージョンによってサポートされていることを確かめます。Office ホスト アプリケーションとサーバーの要件の詳細については、「[Office アドインを実行するための要件](../../docs/overview/requirements-for-running-office-add-ins.md)」を参照してください。
