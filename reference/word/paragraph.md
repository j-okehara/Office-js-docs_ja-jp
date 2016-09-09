# 段落オブジェクト (JavaScript API for Word)

選択部分、範囲、コンテンツ コントロール、または文書本文に含まれる 1 つの段落を表します。

_適用対象:Word 2016、Word for iPad、Word for Mac_

## プロパティ
| プロパティ     | 型   |説明
|:---------------|:--------|:----------|
|outlineLevel|int|段落のアウトライン レベルを取得または設定します。|
|style|string|段落に使用するスタイルを取得または設定します。 これは、事前にインストールされているスタイルまたはユーザー設定のスタイルの名前です。 [Word-Add-in-DocumentAssembly][paragraph.style] サンプルは、段落スタイルを設定する方法を示しています。|
|text|string|段落のテキストを取得します。読み取り専用です。|

## リレーションシップ
| リレーションシップ | 型   |説明|
|:---------------|:--------|:----------|
|整列|**位置揃え**|段落の配置を取得または設定します。有効な値は、"left"、"centered"、"right"、または "justified" です。|
|contentControls|[ContentControlCollection](contentcontrolcollection.md)|段落に含まれるコンテンツ コントロール オブジェクトのコレクションを取得します。読み取り専用です。|
|firstLineIndent|**浮動小数点数**|最初の行のインデントまたはぶら下げインデントの値をポイント数単位で取得または設定します。最初の行のインデントを設定するには、正の値を使用します。また、ぶら下げインデントを設定するには、負の値を使用します。|
|Font|[フォント](font.md)|段落のテキスト形式を取得します。これを使用して、フォント名、サイズ、色、およびその他のプロパティを取得および設定します。読み取り専用です。|
|inlinePictures|[InlinePictureCollection](inlinepicturecollection.md)|段落に含まれる inlinePicture オブジェクトのコレクションを取得します。コレクションに浮動イメージは含まれません。読み取り専用です。|
|leftIndent|**浮動小数点数**|段落の左インデントの値をポイント数単位で取得または設定します。|
|lineSpacing|**浮動小数点数**|段落の行間をポイント数単位で取得または設定します。Word UI では、この値が 12 で除算されます。|
|lineUnitAfter|**浮動小数点数**|段落後の間隔の幅をグリッド線数単位で取得または設定します。|
|lineUnitBefore|**浮動小数点数**|段落前の間隔の幅をグリッド線数単位で取得または設定します。|
|parentContentControl|[ContentControl](contentcontrol.md)|段落を格納しているコンテンツ コントロールを取得します。親コンテンツ コントロールがない場合は null を返します。読み取り専用です。|
|rightIndent|**浮動小数点数**|段落の右インデントの値をポイント数単位で取得または設定します。|
|spaceAfter|**浮動小数点数**|段落後の間隔をポイント数単位で取得または設定します。|
|spaceBefore|**浮動小数点数**|段落前の間隔をポイント数単位で取得または設定します。|

## メソッド

| メソッド           | 戻り値の型    |説明|
|:---------------|:--------|:----------|
|[clear()](#clear)|void|段落オブジェクトの内容をクリアします。ユーザーは、消去された内容を元に戻す操作を実行できます。|
|[delete()](#delete)|void|文書から段落と、その段落の内容を削除します。|
|[getHtml()](#gethtml)|string|Paragraph オブジェクトの HTML 表記を取得します。|
|[getOoxml()](#getooxml)|string|Paragraph オブジェクトの Office Open XML (OOXML) 表記を取得します。|
|[insertBreak(breakType: BreakType, insertLocation: InsertLocation)](#insertbreakbreaktype-breaktype-insertlocation-insertlocation)|void|指定した位置に区切りを挿入します。改行以外の区切りは、メイン文書の本体に含まれている段落にのみ挿入できます。改行はどの本文オブジェクトにも挿入できます。insertLocation の値には 'After' または 'Before' を指定できます。|
|[insertContentControl()](#insertcontentcontrol)|[ContentControl](contentcontrol.md)|段落オブジェクトを、リッチ テキストのコンテンツ コントロールでラップします。|
|[insertFileFromBase64(base64File: string, insertLocation:InsertLocation)](#insertfilefrombase64base64file-string-insertlocation-insertlocation)|[範囲](range.md)|現在の段落の指定した位置に、文書を挿入します。insertLocation の値には、'Start' または 'End' を指定できます。|
|[insertHtml(html: string, insertLocation:InsertLocation)](#inserthtmlhtml-string-insertlocation-insertlocation)|[範囲](range.md)|段落の指定した位置に、HTML を挿入します。insertLocation の値には、'Replace'、'Start'、'End' のいずれかを指定できます。|
|[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: InsertLocation)](#insertinlinepicturefrombase64base64encodedimage-string-insertlocation-insertlocation)|[InlinePicture](inlinepicture.md)|段落の指定した位置に、図を挿入します。insertLocation の値には、'Before'、'After'、'Start'、'End' のいずれかを指定できます。|
|[insertOoxml(ooxml: string, insertLocation: InsertLocation)](#insertooxmlooxml-string-insertlocation-insertlocation)|[範囲](range.md)|段落内の指定された位置に OOXML または wordProcessingML を挿入します。insertLocation の値には、'Replace'、'Start'、'End' のいずれかを指定できます。|
|[insertParagraph(paragraphText: string, insertLocation: InsertLocation)](#insertparagraphparagraphtext-string-insertlocation-insertlocation)|[段落](paragraph.md)|指定した位置に、段落を挿入します。有効な insertLocation の値は、'Before' または 'After' です。|
|[insertText(text: string, insertLocation:InsertLocation)](#inserttexttext-string-insertlocation-insertlocation)|[範囲](range.md)|段落の指定した位置に、テキストを挿入します。insertLocation の値には、'Replace'、'Start'、'End' のいずれかを指定できます。|
|[load(param: object)](#loadparam-object)|void|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|
|[search(searchText: string, searchOptions:ParamTypeStrings.SearchOptions)](#searchsearchtext-string-searchoptions-paramtypestrings.searchoptions)|[SearchResultCollection](searchresultcollection.md)|段落オブジェクトの範囲で、searchOptions を指定した検索を実行します。検索結果は、範囲オブジェクトのコレクションです。|
|[select(selectionMode: SelectionMode)](#selectselectionmode-selectionmode)|void|段落を選択して、その段落に Word の UI を移動します。選択モードは、'Select'、'Start'、'End' のいずれかになります。'Select' が既定値です。|

## メソッドの詳細

### clear()
段落オブジェクトの内容をクリアします。ユーザーは、消去された内容を元に戻す操作を実行できます。

#### 構文
```js
paragraphObject.clear();
```

#### パラメーター
なし

#### 戻り値
void

#### 例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;

    // Queue a commmand to load the style property for all of the paragraphs.
    context.load(paragraphs, 'style');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {

        // Queue a command to clear the contents of the first paragraph.
        paragraphs.items[0].clear();

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Cleared the contents of the first paragraph.');
        });
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
文書から段落と、その段落の内容を削除します。

#### 構文
```js
paragraphObject.delete();
```

#### パラメーター
なし

#### 戻り値
void

#### 例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;

    // Queue a commmand to load the text property for all of the paragraphs.
    context.load(paragraphs, 'text');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {

        // Queue a command to delete the first paragraph.
        paragraphs.items[0].delete();

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Deleted the first paragraph.');
        });
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
Paragraph オブジェクトの HTML 表記を取得します。

#### 構文
```js
paragraphObject.getHtml();
```

#### パラメーター
なし

#### 戻り値
string

#### 例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;

    // Queue a commmand to load the style property for all of the paragraphs.
    context.load(paragraphs, 'style');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {

        // Queue a a set of commands to get the HTML of the first paragraph.
        var html = paragraphs.items[0].getHtml();

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Paragraph HTML: ' + html.value);
        });
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
Paragraph オブジェクトの Office Open XML (OOXML) 表記を取得します。

#### 構文
```js
paragraphObject.getOoxml();
```

#### パラメーター
なし

#### 戻り値
string

#### 例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;

    // Queue a commmand to load the style property for the top 2 paragraphs.
    context.load(paragraphs, {select: 'style', top: 2} );

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {

        // Queue a a set of commands to get the OOXML of the first paragraph.
        var ooxml = paragraphs.items[0].getOoxml();

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Paragraph OOXML: ' + ooxml.value);
        });
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
指定した位置に区切りを挿入します。改行以外の区切りは、メイン文書の本体に含まれている段落にのみ挿入できます。改行はどの本文オブジェクトにも挿入できます。insertLocation の値には 'Before' または 'After' を指定できます。

#### 構文
```js
paragraphObject.insertBreak(breakType, insertLocation);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|breakType|BreakType|必須。文書に追加する区切りの種類。|
|insertLocation|InsertLocation|必須。有効な値は、'Before' または 'After' です。|

#### 戻り値
void

#### 追加の詳細
ヘッダー、フッター、脚注、文末脚注、コメント、テキスト ボックスには、区切りを挿入できません。

#### 例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;

    // Queue a commmand to load the style property for the top 2 paragraphs.
    // We never perform an empty load. We always must request a property.
    context.load(paragraphs, {select: 'style', top: 2} );

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {

        // Queue a command to get the first paragraph.
        var paragraph = paragraphs.items[0];

        // Queue a command to insert a page break after the first paragraph.
        paragraph.insertBreak('page', 'After');

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Inserted a page break after the paragraph.');
        });
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
段落オブジェクトを、リッチ テキストのコンテンツ コントロールでラップします。

#### 構文
```js
paragraphObject.insertContentControl();
```

#### パラメーター
なし

#### 戻り値
[ContentControl](contentcontrol.md)

#### 例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;

    // Queue a commmand to load the style property for the top 2 paragraphs.
    // We never perform an empty load. We always must request a property.
    context.load(paragraphs, {select: 'style', top: 2} );

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {

        // Queue a command to get the first paragraph.
        var paragraph = paragraphs.items[0];

        // Queue a command to wrap the first paragraph in a rich text content control.
        paragraph.insertContentControl();

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Wrapped the first paragraph in a content control.');
        });
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
[Word-Add-in-DocumentAssembly][paragraph.insertContentControl] サンプルは、insertContentControl メソッドを使う方法を示しています。

### insertFileFromBase64(base64File: string, insertLocation:InsertLocation)
現在の段落の指定した位置に、文書を挿入します。insertLocation の値には、'Start' または 'End' を指定できます。

#### 構文
```js
paragraphObject.insertFileFromBase64(base64File, insertLocation);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|base64File|string|必須。挿入するファイルの内容が base64 エンコードされているファイル。|
|insertLocation|InsertLocation|必須。有効な値は、'Start' または 'End' です。|

#### 戻り値
[範囲](range.md)

#### 例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;

    // Queue a commmand to load the style property for all of the paragraphs.
    context.load(paragraphs, 'style');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {

        // Queue a command to get the first paragraph.
        var paragraph = paragraphs.items[0];

        // Queue a command to insert base64 encoded .docx at the beginning of the first paragraph.
        // This won't work unless you have a definition for getBase64().
        paragraph.insertFileFromBase64(getBase64(), Word.InsertLocation.start);

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Inserted base64 encoded content at the beginning of the first paragraph.');
        });
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
段落の指定した位置に、HTML を挿入します。insertLocation の値には、'Replace'、'Start'、'End' のいずれかを指定できます。

#### 構文
```js
paragraphObject.insertHtml(html, insertLocation);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|Html|string|必須。段落に挿入する HTML。|
|insertLocation|InsertLocation|必須。値には、'Replace'、'Start'、'End' のいずれかを指定できます。|

#### 戻り値
[範囲](range.md)

#### 例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;

    // Queue a commmand to load the style property for the top 2 paragraphs.
    // We never perform an empty load. We always must request a property.
    context.load(paragraphs, {select: 'style', top: 2} );

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {

        // Queue a command to get the first paragraph.
        var paragraph = paragraphs.items[0];

        // Queue a command to insert HTML content at the end of the first paragraph.
        paragraph.insertHtml('<strong>Inserted HTML.</strong>', Word.InsertLocation.end);

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Inserted HTML content at the end of the first paragraph.');
        });
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});

```

### insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation:InsertLocation)
段落の指定した位置に、図を挿入します。insertLocation の値には、'Before'、'After'、'Start'、'End' のいずれかを指定できます。

#### 構文
```js
paragraphObject.insertInlinePictureFromBase64(base64EncodedImage, insertLocation);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|base64EncodedImage|string|必須。段落に挿入する HTML。|
|insertLocation|InsertLocation|必須。値には、'Before'、'After'、'Start'、または 'End' を指定できます。|

#### 戻り値
[InlinePicture](inlinepicture.md)

#### 例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;

    // Queue a commmand to load the style property for all of the paragraphs.
    context.load(paragraphs, 'style');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {

        // Queue a command to get the first paragraph.
        var paragraph = paragraphs.items[0];

        var b64encodedImg = "iVBORw0KGgoAAAANSUhEUgAAAB4AAAANCAIAAAAxEEnAAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAACFSURBVDhPtY1BEoQwDMP6/0+XgIMTBAeYoTqso9Rkx1zG+tNj1H94jgGzeNSjteO5vtQQuG2seO0av8LzGbe3anzRoJ4ybm/VeKEerAEbAUpW4aWQCmrGFWykRzGBCnYy2ha3oAIq2MloW9yCCqhgJ6NtcQsqoIKdjLbFLaiACnYyf2fODbrjZcXfr2F4AAAAAElFTkSuQmCC";

        // Queue a command to insert a base64 encoded image at the beginning of the first paragraph.
        paragraph.insertInlinePictureFromBase64(b64encodedImg, Word.InsertLocation.start);

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Added an image to the first paragraph.');
        });
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
[Word-Add-in-DocumentAssembly][paragraph.insertpicture] サンプルは、段落に画像を挿入する別の例を提供しています。

### insertOoxml(ooxml: string, insertLocation:InsertLocation)
段落内の指定された位置に OOXML または wordProcessingML を挿入します。insertLocation の値には、'Replace'、'Start'、'End' のいずれかを指定できます。

#### 構文
```js
paragraphObject.insertOoxml(ooxml, insertLocation);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|ooxml|string|必須。段落に挿入する OOXML または wordProcessingML。|
|insertLocation|InsertLocation|必須。値には、'Replace'、'Start'、'End' のいずれかを指定できます。|

#### 戻り値
[範囲](range.md)

#### 例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;

    // Queue a commmand to load the style property for the top 2 paragraphs.
    // We never perform an empty load. We always must request a property.
    context.load(paragraphs, {select: 'style', top: 2} );

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {

        // Queue a command to get the first paragraph.
        var paragraph = paragraphs.items[0];

        // Queue a command to insert Ooxml content into the first paragraph.
        var ooxmlContent = "<pkg:package xmlns:pkg='http://schemas.microsoft.com/office/2006/xmlPackage'><pkg:part pkg:name='/_rels/.rels' pkg:contentType='application/vnd.openxmlformats-package.relationships+xml' pkg:padding='512'><pkg:xmlData><Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument' Target='word/document.xml'/></Relationships></pkg:xmlData></pkg:part><pkg:part pkg:name='/word/document.xml' pkg:contentType='application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'><pkg:xmlData><w:document xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' ><w:body><w:p><w:pPr><w:spacing w:before='360' w:after='0' w:line='480' w:lineRule='auto'/><w:rPr><w:color w:val='70AD47' w:themeColor='accent6'/><w:sz w:val='28'/></w:rPr></w:pPr><w:r><w:rPr><w:color w:val='70AD47' w:themeColor='accent6'/><w:sz w:val='28'/></w:rPr><w:t>This text has formatting directly applied to achieve its font size, color, line spacing, and paragraph spacing.</w:t></w:r></w:p></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>";
        paragraph.insertOoxml(ooxmlContent, Word.InsertLocation.end);

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Inserted OOXML at the end of the first paragraph.');
        });
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
指定した位置に、段落を挿入します。有効な insertLocation の値は、'Before' または 'After' です。

#### 構文
```js
paragraphObject.insertParagraph(paragraphText, insertLocation);
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

    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;

    // Queue a commmand to load the style property for the top 2 paragraphs.
    // We never perform an empty load. We always must request a property.
    context.load(paragraphs, {select: 'style', top: 2} );

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {

        // Queue a command to get the first paragraph.
        var paragraph = paragraphs.items[0];

        // Queue a command to insert the paragraph after the current paragraph.
        paragraph.insertParagraph('Content of a new paragraph', Word.InsertLocation.after);

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Inserted a new paragraph at the end of the first paragraph.');
        });
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
段落の指定した位置に、テキストを挿入します。insertLocation の値には、'Replace'、'Start'、'End' のいずれかを指定できます。

#### 構文
```js
paragraphObject.insertText(text, insertLocation);
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

    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;

    // Queue a commmand to load the style property for the top 2 paragraphs.
    // We never perform an empty load. We always must request a property.
    context.load(paragraphs, {select: 'style', top: 2} );

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {

        // Queue a command to get the first paragraph.
        var paragraph = paragraphs.items[0];

        // Queue a command to insert text into the end of the paragraph.
        paragraph.insertText('New text inserted into the paragraph.', Word.InsertLocation.end);

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Inserted text at the end of the first paragraph.');
        });
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

    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;

    // Queue a commmand to load the style property for the top 2 paragraphs.
    // We never perform an empty load. We always must request a property.
    context.load(paragraphs, {select: 'style', top: 2} );

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {

        // Queue a command to get the first paragraph.
        var paragraph = paragraphs.items[0];

        // Queue a command to load font information for the paragraph.
        context.load(paragraph, 'font/size, font/name, font/color');

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            // Show the results of the load method. Here we show the
            // property values on the paragraph object. Note that we
            // requested the style property in the first load command.
            var results = "<strong>Paragraph</strong>--" +
                          "--Font size: " + paragraph.font.size +
                          "--Font name: " + paragraph.font.name +
                          "--Font color: " + paragraph.font.color +
                          "--Style: " + paragraph.style;

            console.log(results);
        });
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
段落オブジェクトの範囲で、searchOptions を指定した検索を実行します。検索結果は、範囲オブジェクトのコレクションです。

#### 構文
```js
paragraphObject.search(searchText, searchOptions);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|searchText|string|必須。検索テキスト。|
|[searchOptions](searchoptions.md)|ParamTypeStrings.SearchOptions|省略可能。検索のオプション。|

#### 戻り値
[SearchResultCollection](searchresultcollection.md)

### select(selectionMode: SelectionMode)
段落を選択して、その段落に Word の UI を移動します。

#### 構文
```js
paragraphObject.select(selectionMode);
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

    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;

    // Queue a commmand to load the style property for all of the paragraphs.
    context.load(paragraphs, 'style');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {

        // Queue a command to get the last paragraph a create a
        // proxy paragraph object.
        var paragraph = paragraphs.items[paragraphs.items.length - 1];

        // Queue a command to select the paragraph. The Word UI will
        // move to the selected paragraph.
        paragraph.select();

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Selected the last paragraph.');
        });
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


[paragraph.insertContentControl]: https://github.com/OfficeDev/Word-Add-in-DocumentAssembly/blob/master/WordAPIDocAssemblySampleWeb/App/Home/Home.js#L161 "iコンテンツ コントロールを挿入します"
[paragraph.style]: https://github.com/OfficeDev/Word-Add-in-DocumentAssembly/blob/master/WordAPIDocAssemblySampleWeb/App/Home/Home.js#L172 "スタイルを設定します"
[paragraph.insertpicture]: https://github.com/OfficeDev/Word-Add-in-DocumentAssembly/blob/master/WordAPIDocAssemblySampleWeb/App/Home/Home.js#L236 "画像を挿入します"