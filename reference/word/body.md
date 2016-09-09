# 本文オブジェクト (JavaScript API for Word)

文書またはセクションの本文を表します。

_適用対象:Word 2016、Word for iPad、Word for Mac_

## プロパティ
| プロパティ     | 型   |説明
|:---------------|:--------|:----------|
|style|string|本文に使用されるスタイルを取得または設定します。これは、事前にインストールされている、またはユーザー定義のスタイルの名前です。|
|text|string|本文のテキストを取得します。insertText メソッドを使用して、テキストを挿入します。読み取り専用です。|

_プロパティのアクセスの[例](#例)を参照してください。_

## リレーションシップ
| リレーションシップ | 型   |説明|
|:---------------|:--------|:----------|
|contentControls|[ContentControlCollection](contentcontrolcollection.md)|本文に含まれるリッチ テキストのコンテンツ コントロール オブジェクトのコレクションを取得します。読み取り専用です。|
|Font|[フォント](font.md)|本文のテキスト形式を取得します。これを使用して、フォント名、サイズ、色、およびその他のプロパティを取得および設定します。読み取り専用です。|
|inlinePictures|[InlinePictureCollection](inlinepicturecollection.md)|本文に含まれる inlinePicture オブジェクトのコレクションを取得します。コレクションには、浮動イメージは含まれません。読み取り専用です。|
|paragraphs|[ParagraphCollection](paragraphcollection.md)|本文に含まれる段落オブジェクトのコレクションを取得します。読み取り専用です。|
|parentContentControl|[ContentControl](contentcontrol.md)|本文を含むコンテンツ コントロールを取得します。親コンテンツ コントロールがない場合は null を返します。読み取り専用です。|

## メソッド

| メソッド           | 戻り値の型    |説明|
|:---------------|:--------|:----------|
|[clear()](#clear)|void|本文オブジェクトの内容を消去します。ユーザーは、消去された内容を元に戻す操作を実行できます。|
|[getHtml()](#gethtml)|string|本文のオブジェクトの HTML 表記を取得します。|
|[getOoxml()](#getooxml)|string|本文オブジェクトの OOXML (Office オープン XML) 表記を取得します。|
|[insertBreak(breakType: BreakType, insertLocation: InsertLocation)](#insertbreakbreaktype-breaktype-insertlocation-insertlocation)|void|指定した位置に区切りを挿入します。区切りは、どの本文オブジェクトにも挿入可能な改行である場合を除き、メイン文書本文にのみ挿入できます。insertLocation の値には、'Start' または 'End' を指定できます。|
|[insertContentControl()](#insertcontentcontrol)|[ContentControl](contentcontrol.md)|リッチ テキスト コンテンツ コントロールで本文オブジェクトをラップします。|
|[insertFileFromBase64(base64File: string, insertLocation:InsertLocation)](#insertfilefrombase64base64file-string-insertlocation-insertlocation)|[範囲](range.md)|文書を本文の指定された位置に挿入します。insertLocation の値には、'Replace'、'Start'、'End' のいずれかを指定できます。|
|[insertHtml(html: string, insertLocation:InsertLocation)](#inserthtmlhtml-string-insertlocation-insertlocation)|[範囲](range.md)|指定した位置に HTML を挿入します。insertLocation の値には、'Replace'、'Start'、'End' のいずれかを指定できます。|
|[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: InsertLocation)](#insertinlinepicturefrombase64base64encodedimage-string-insertlocation-insertlocation)|[InlinePicture](inlinepicture.md)|画像を本文の指定された位置に挿入します。insertLocation の値は、'Start' か 'End' になります。 |
|[insertOoxml(ooxml: string, insertLocation: InsertLocation)](#insertooxmlooxml-string-insertlocation-insertlocation)|[範囲](range.md)|指定した位置に OOXML を挿入します。  insertLocation の値には、'Replace'、'Start'、'End' のいずれかを指定できます。|
|[insertParagraph(paragraphText: string, insertLocation:InsertLocation)](#insertparagraphparagraphtext-string-insertlocation-insertlocation)|[段落](paragraph.md)|指定した位置に段落を挿入します。insertLocation の値には、'Start' または 'End' を指定できます。|
|[insertText(text: string, insertLocation:InsertLocation)](#inserttexttext-string-insertlocation-insertlocation)|[範囲](range.md)|テキストを本文の指定された位置に挿入します。insertLocation の値には、'Replace'、'Start'、'End' のいずれかを指定できます。|
|[load(param: object)](#loadparam-object)|void|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|
|[search(searchText: string, searchOptions:ParamTypeStrings.SearchOptions)](#searchsearchtext-string-searchoptions-paramtypestrings.searchoptions)|[SearchResultCollection](searchresultcollection.md)|本文オブジェクトの範囲で、指定した searchOptions を使って検索を実行します。検索結果は、範囲オブジェクトのコレクションです。|
|[select(selectionMode: SelectionMode)](#selectselectionmode-selectionmode)|void|本文を選択し、その本文に Word の UI を移動します。selectionMode 値は、'Select'、'Start'、'End' のいずれかになります。|

## メソッドの詳細

### clear()
本文オブジェクトの内容を消去します。ユーザーは、消去された内容を元に戻す操作を実行できます。

#### 構文
```js
bodyObject.clear();
```

#### パラメーター
なし

#### 戻り値
void

#### 例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to clear the contents of the body.
    body.clear();

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Cleared the body contents.');
    });
})
.catch(function (error) {
    console.log("Error: " + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

[Silly stories](https://aka.ms/sillystorywordaddin) アドイン サンプルは、**clear** メソッドを使用して文書のコンテンツをクリアする方法を示します。

### getHtml()
本文のオブジェクトの HTML 表記を取得します。

#### 構文
```js
bodyObject.getHtml();
```

#### パラメーター
なし

#### 戻り値
string

#### 例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to get the HTML contents of the body.
    var bodyHTML = body.getHtml();

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log("Body HTML contents: " + bodyHTML.value);
    });
})
.catch(function (error) {
    console.log("Error: " + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

### getOoxml()
本文オブジェクトの OOXML (Office オープン XML) 表記を取得します。

#### 構文
```js
bodyObject.getOoxml();
```

#### パラメーター
なし

#### 戻り値
string

#### 例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to get the OOXML contents of the body.
    var bodyOOXML = body.getOoxml();

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log("Body OOXML contents: " + bodyOOXML.value);
    });
})
.catch(function (error) {
    console.log("Error: " + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

### insertBreak(breakType: BreakType, insertLocation: InsertLocation)
指定した位置に区切りを挿入します。区切りは、どの本文オブジェクトにも挿入可能な改行である場合を除き、メイン文書本文にのみ挿入できます。insertLocation の値には、'Start' または 'End' を指定できます。

#### 構文
```js
bodyObject.insertBreak(breakType, insertLocation);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|breakType|BreakType|必須。本文に追加する区切りの種類。|
|insertLocation|InsertLocation|必須。有効な値は、'Start' または 'End' です。|

#### 戻り値
void

#### 詳細
ヘッダー、フッター、脚注、文末脚注、コメント、テキスト ボックスに、改行以外の区切りを挿入することはできません。

#### 例
```js
// Run a batch operation against the Word object model.
Word.run(function (ctx) {

    // Create a proxy object for the document body.
    var body = ctx.document.body;

    // Queue a commmand to insert a page break at the start of the document body.
    body.insertBreak(Word.BreakType.page, Word.InsertLocation.start);

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return ctx.sync().then(function () {
        console.log('Added a page break at the start of the document body.');
    });
})
.catch(function (error) {
    console.log("Error: " + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```
### insertContentControl()
リッチ テキスト コンテンツ コントロールで本文オブジェクトをラップします。

#### 構文
```js
bodyObject.insertContentControl();
```

#### パラメーター
なし

#### 戻り値
[ContentControl](contentcontrol.md)

#### 例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to wrap the body in a content control.
    body.insertContentControl();

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Wrapped the body in a content control.');
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
文書を本文の指定された位置に挿入します。insertLocation の値には、'Replace'、'Start'、'End' のいずれかを指定できます。

#### 構文
```js
bodyObject.insertFileFromBase64(base64File, insertLocation);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|base64File|string|必須。挿入する base64 エンコード ファイルの内容。|
|insertLocation|InsertLocation|必須。値には、'Replace'、'Start'、'End' のいずれかを指定できます。|

#### 戻り値
[範囲](range.md)

#### 例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to insert base64 encoded .docx at the beginning of the content body.
    // You will need to implement getBase64() to pass in a string of a base64 encoded docx file.
    body.insertFileFromBase64(getBase64(), Word.InsertLocation.start);

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Added base64 encoded text to the beginning of the document body.');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

[Silly stories](https://aka.ms/sillystorywordaddin) アドイン サンプルは、**insertFileFromBase64** メソッドを使用して、サービスから docx ファイルを挿入する方法を示しています。

### insertHtml(html: string, insertLocation:InsertLocation)
指定した位置に HTML を挿入します。insertLocation の値には、'Replace'、'Start'、'End' のいずれかを指定できます。

#### 構文
```js
bodyObject.insertHtml(html, insertLocation);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|Html|string|必須。文書に挿入する HTML。|
|insertLocation|InsertLocation|必須。値には、'Replace'、'Start'、'End' のいずれかを指定できます。|

#### 戻り値
[範囲](range.md)

#### 例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to insert HTML in to the beginning of the body.
    body.insertHtml('<strong>This is text inserted with body.insertHtml()</strong>', Word.InsertLocation.start);

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('HTML added to the beginning of the document body.');
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
画像を本文の指定された位置に挿入します。insertLocation の値は、'Start' か 'End' になります。

#### 構文
bodyObject.insertInlinePictureFromBase64(image, insertLocation);

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|base64EncodedImage|string|必須。本文に挿入される base64 でエンコードされた画像。|
|insertLocation|InsertLocation|必須。有効な値は、'Start' または 'End' です。|

#### 戻り値
[InlinePicture](inlinepicture.md)

### insertOoxml(ooxml: string, insertLocation:InsertLocation)
指定した位置に OOXML を挿入します。insertLocation の値には、'Replace'、'Start'、'End' のいずれかを指定できます。

#### 構文
```js
bodyObject.insertOoxml(ooxml, insertLocation);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|ooxml|string|必須。挿入する OOXML または wordProcessingML。|
|insertLocation|InsertLocation|必須。値には、'Replace'、'Start'、'End' のいずれかを指定できます。|

#### 戻り値
[範囲](range.md)

#### 例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to insert OOXML in to the beginning of the body.
    body.insertOoxml("<pkg:package xmlns:pkg='http://schemas.microsoft.com/office/2006/xmlPackage'><pkg:part pkg:name='/_rels/.rels' pkg:contentType='application/vnd.openxmlformats-package.relationships+xml' pkg:padding='512'><pkg:xmlData><Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument' Target='word/document.xml'/></Relationships></pkg:xmlData></pkg:part><pkg:part pkg:name='/word/document.xml' pkg:contentType='application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'><pkg:xmlData><w:document xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' ><w:body><w:p><w:pPr><w:spacing w:before='360' w:after='0' w:line='480' w:lineRule='auto'/><w:rPr><w:color w:val='70AD47' w:themeColor='accent6'/><w:sz w:val='28'/></w:rPr></w:pPr><w:r><w:rPr><w:color w:val='70AD47' w:themeColor='accent6'/><w:sz w:val='28'/></w:rPr><w:t>This text has formatting directly applied to achieve its font size, color, line spacing, and paragraph spacing.</w:t></w:r></w:p></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>", Word.InsertLocation.start);

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('OOXML added to the beginning of the document body.');
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
OOXML の操作の詳細については、「[Office Open XML を使用してより良い Word 用アドインを作成する](https://msdn.microsoft.com/en-us/library/office/dn423225.aspx)」をご覧ください。 [[Word-Add-in-DocumentAssembly]][[body.insertOoxml]] サンプルは、この API を使ってドキュメントを組み立てる方法を示しています。

### insertParagraph(paragraphText: string, insertLocation:InsertLocation)
指定した位置に段落を挿入します。insertLocation の値には、'Start' または 'End' を指定できます。

#### 構文
```js
bodyObject.insertParagraph(paragraphText, insertLocation);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|paragraphText|string|必須。挿入する段落テキスト。|
|insertLocation|InsertLocation|必須。有効な値は、'Start' または 'End' です。|

#### 戻り値
[段落](paragraph.md)

#### 例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to insert the paragraph at the end of the document body.
    body.insertParagraph('Content of a new paragraph', Word.InsertLocation.end);

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Paragraph added at the end of the document body.');
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
[[Word-Add-in-DocumentAssembly]][[body.insertParagraph]] サンプルは、insertParagraph メソッドを使ってドキュメントを組み立てる方法を示しています。

### insertText(text: string, insertLocation:InsertLocation)
テキストを本文の指定された位置に挿入します。insertLocation の値には、'Replace'、'Start'、'End' のいずれかを指定できます。

#### 構文
```js
bodyObject.insertText(text, insertLocation);
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

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to insert text in to the beginning of the body.
    body.insertText('This is text inserted with body.insertText()', Word.InsertLocation.start);

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Text added to the beginning of the document body.');
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

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to load font and style information for the document body.
    context.load(body, 'font/size, font/name, font/color, style');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        // Show the results of the load method. Here we show the
        // property values on the body object.
        var results = 'Font size: ' + body.font.size +
                      '; Font name: ' + body.font.name +
                      '; Font color: ' + body.font.color +
                      '; Body style: ' + body.style;

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
本文オブジェクトの範囲で、指定した検索オプションを使って検索を実行します。検索結果は、範囲オブジェクトのコレクションです。

#### 構文
```js
bodyObject.search(searchText, searchOptions);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|searchText|string|必須。検索テキスト。|
|[searchOptions](searchoptions.md)|ParamTypeStrings.SearchOptions|省略可能。検索のオプション。|

#### 戻り値
[SearchResultCollection](searchresultcollection.md)

#### 例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to search the document.
    var searchResults = context.document.body.search('video', {matchCase: false});

    // Queue a commmand to load the results.
    context.load(searchResults, 'text, font');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        var results = 'Found count: ' + searchResults.items.length +
                      '; we highlighted the results.';

        // Queue a command to change the font for each found item.
        for (var i = 0; i < searchResults.items.length; i++) {
          searchResults.items[i].font.color = '#FF0000'    // Change color to Red
          searchResults.items[i].font.highlightColor = '#FFFF00';
          searchResults.items[i].font.bold = true;
        }

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
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

#### 追加情報
[Word-Add-in-DocumentAssembly][body.search] サンプルは、ドキュメントを検索する方法の別の例を示しています。

### select(selectionMode: SelectionMode)
本文を選択し、その本文に Word の UI を移動します。selectionMode 値は、'Select'、'Start'、'End' のいずれかになります。

#### 構文
```js
bodyObject.select(selectionMode);
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

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to select the document body. The Word UI will
    // move to the selected document body.
    body.select();

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Selected the document body.');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

## プロパティのアクセスの例

### 本文オブジェクトのテキスト プロパティを取得します。
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to load the text in document body.
    context.load(body, 'text');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log("Body contents: " + body.text);
    });
})
.catch(function (error) {
    console.log("Error: " + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```
### 本文オブジェクトのスタイルとフォント サイズ、フォント名、フォントの色のプロパティを取得します。

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to load font and style information for the document body.
    context.load(body, 'font/size, font/name, font/color, style');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        // Show the results of the load method. Here we show the
        // property values on the body object.
        var results = 'Font size: ' + body.font.size +
                      '; Font name: ' + body.font.name +
                      '; Font color: ' + body.font.color +
                      '; Body style: ' + body.style;

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

## サポートの詳細

実行時のチェックで[要件セット](../office-add-in-requirement-sets.md)を使用して、アプリケーションが Word のホスト バージョンによってサポートされていることを確かめます。Office ホスト アプリケーションとサーバーの要件の詳細については、「[Office アドインを実行するための要件](../../docs/overview/requirements-for-running-office-add-ins.md)」を参照してください。


[body.insertOoxml]: https://github.com/OfficeDev/Word-Add-in-DocumentAssembly/blob/master/WordAPIDocAssemblySampleWeb/App/Home/Home.js#L127 "insert OOXML"
[body.insertParagraph]: https://github.com/OfficeDev/Word-Add-in-DocumentAssembly/blob/master/WordAPIDocAssemblySampleWeb/App/Home/Home.js#L153 "insert paragraph"
[body.search]: https://github.com/OfficeDev/Word-Add-in-DocumentAssembly/blob/master/WordAPIDocAssemblySampleWeb/App/Home/Home.js#L261 "body search"
