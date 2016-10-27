# <a name="contentcontrol-object-(javascript-api-for-word)"></a>ContentControl オブジェクト (JavaScript API for Word)

コンテンツ コントロールを表します。コンテンツ コントロールは、特定の種類のコンテンツのコンテナーとして機能し、ドキュメント内で境界線で区切られ、ラベルが付いた領域になる場合もあります。個々のコンテンツ コントロールには、画像、表、書式設定されたテキストの段落などの内容が含まれていることがあります。現時点では、リッチ テキスト コンテンツ コントロールのみがサポートされています。

_適用対象:Word 2016、Word for iPad、Word for Mac、Word Online_

## <a name="properties"></a>プロパティ
| プロパティ     | 型   |説明
|:---------------|:--------|:----------|
|cannotDelete|bool|ユーザーがコンテンツ コントロールを削除できるかどうかを示す値を取得または設定します。removeWhenEdited と同時に使用することはできません。|
|cannotEdit|bool|ユーザーがコンテンツ コントロールのコンテンツを編集できるかどうかを示す値を取得または設定します。|
|color|string|コンテンツ コントロールの色を取得または設定します。色は、"#RRGGBB" 形式で設定するか、色の名前を使用して設定します。|
|placeholderText|文字列|コンテンツ コントロールのプレースホルダー テキストを取得または設定します。コンテンツ コントロールが空の場合は、淡色のテキストが表示されます。このプロパティは Word オンラインで現在サポートされていません。|
|removeWhenEdited|bool|コンテンツ コントロールを編集後に削除できるかどうかを示す値を取得または設定します。cannotDelete と同時に使用することはできません。|
|style|string|コンテンツ コントロールに使用するスタイルを取得または設定します。これは、事前にインストールされている、またはユーザー定義のスタイルの名前です。|
|tag|string|コンテンツ コントロールを識別するタグを取得または設定します。[Silly stories](https://aka.ms/sillystorywordaddin) サンプル アドインは、**tag** プロパティの使用方法を示しています。|
|text|string|コンテンツ コントロールのテキストを取得します。読み取り専用です。|
|title|string|コンテンツ コントロールのタイトルを取得または設定します。|

_プロパティのアクセスの[例を参照してください。](#property-access-examples)_

## <a name="relationships"></a>関係
| リレーションシップ | 型   |説明|
|:---------------|:--------|:----------|
|appearance|**ContentControlAppearance**|コンテンツ コントロールの外観を取得または設定します。値には 'boundingBox'、'tags'、または 'hidden' を指定できます。|
|contentControls|[ContentControlCollection](contentcontrolcollection.md)|コンテンツ コントロールのコンテンツ コントロール オブジェクトのコレクションを取得します。読み取り専用です。|
|font|[Font](font.md)|コンテンツ コントロールのテキストの書式設定を取得します。これを使用して、フォント名、サイズ、色、およびその他のプロパティを取得および設定します。読み取り専用です。|
|id|**uint**|コンテンツ コントロールの識別子を表す整数値を取得します。読み取り専用です。|
|inlinePictures|[InlinePictureCollection](inlinepicturecollection.md)|コンテンツ コントロールに含まれる inlinePicture オブジェクトのコレクションを取得します。コレクションに浮動イメージは含まれません。読み取り専用です。|
|paragraphs|[ParagraphCollection](paragraphcollection.md)|コンテンツ コントロールにある Paragraph オブジェクトのコレクションを取得します。読み取り専用です。|
|parentContentControl|[ContentControl](contentcontrol.md)|コンテンツ コントロールを含むコンテンツ コントロールを取得します。親コンテンツ コントロールがない場合は null を返します。読み取り専用です。|
|型|**ContentControlType**|コンテンツ コントロールの種類を取得します。現在、リッチ テキストのコンテンツ コントロールのみがサポートされています。読み取り専用です。|

## <a name="methods"></a>メソッド

| メソッド           | 戻り値の型    |説明|
|:---------------|:--------|:----------|
|[clear()](#clear)|void|コンテンツ コントロールの内容をクリアします。ユーザーは、消去された内容を元に戻す操作を実行できます。|
|[delete(keepContent: bool)](#deletekeepcontent-bool)|(非推奨)|コンテンツ コントロールとそのコンテンツを削除します。keepContent が true の場合、コンテンツは削除されません。|
|[getHtml()](#gethtml)|string|コンテンツ コントロール オブジェクトの HTML 表記を取得します。|
|[getOoxml()](#getooxml)|string|コンテンツ コントロール オブジェクトの Office Open XML (OOXML) 表記を取得します。|
|[insertBreak(breakType: BreakType, insertLocation: InsertLocation)](#insertbreakbreaktype-breaktype-insertlocation-insertlocation)|void|指定した位置に、区切りを挿入します。改行以外の区切りは、メインドキュメント本文に含まれるオブジェクトにのみ挿入できます。改行は、いずれの本文オブジェクトにも挿入できます。insertLocation の値には、'Before'、'After'、'Start'、'End' のいずれかを指定できます。|
|[insertFileFromBase64(base64File: string, insertLocation:InsertLocation)](#insertfilefrombase64base64file-string-insertlocation-insertlocation)|[Range](range.md)|現在のコンテンツ コントロール内の指定された位置にドキュメントを挿入します。insertLocation の値には、'Replace'、'Start'、'End' のいずれかを指定できます。|
|[insertHtml(html: string, insertLocation:InsertLocation)](#inserthtmlhtml-string-insertlocation-insertlocation)|[Range](range.md)|コンテンツ コントロール内の指定された位置に HTML を挿入します。insertLocation の値には、'Replace'、'Start'、'End' のいずれかを指定できます。|
|[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation:InsertLocation)](#insertInlinePictureFromBase64base64EncodedImage-string-insertlocation-insertlocation)|[InlinePicture](inlinepicture.md)|コンテンツ コントロール内の指定された位置にインライン画像を挿入します。insertLocation の値には、'Replace'、'Start'、'End' のいずれかを指定できます。 |
|[insertOoxml(ooxml: string, insertLocation:InsertLocation)](#insertooxmlooxml-string-insertlocation-insertlocation)|[Range](range.md)|コンテンツ コントロール内の指定された位置に OOXML または wordProcessingML を挿入します。insertLocation の値には、'Replace'、'Start'、'End' のいずれかを指定できます。|
|[insertParagraph(paragraphText: string, insertLocation:InsertLocation)](#insertparagraphparagraphtext-string-insertlocation-insertlocation)|[Paragraph](paragraph.md)|指定した位置に、段落を挿入します。insertLocation の値には、'Before'、'After'、'Start'、'End' のいずれかを指定できます。|
|[insertText(text: string, insertLocation:InsertLocation)](#inserttexttext-string-insertlocation-insertlocation)|[Range](range.md)|コンテンツ コントロール内の指定された位置にテキストを挿入します。insertLocation の値には、'Replace'、'Start'、'End' のいずれかを指定できます。|
|[load(param: object)](#loadparam-object)|(非推奨)|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|
|[search(searchText: string, searchOptions:ParamTypeStrings.SearchOptions)](#searchsearchtext-string-searchoptions-paramtypestringssearchoptions)|[SearchResultCollection](searchresultcollection.md)|コンテンツ コントロール オブジェクトの範囲で、指定した searchOptions を使って検索を実行します。検索結果は、範囲オブジェクトのコレクションです。|
|[select(selectionMode: SelectionMode)](#selectselectionmode-selectionmode)|void|コンテンツ コントロールを選択します。その結果、Word は選択範囲にスクロールされます。選択モードは、'Select'、'Start'、'End' のいずれかになります。|

## <a name="method-details"></a>メソッドの詳細

### <a name="clear()"></a>clear()
コンテンツ コントロールの内容をクリアします。ユーザーは、消去された内容を元に戻す操作を実行できます。

#### <a name="syntax"></a>構文
```js
contentControlObject.clear();
```

#### <a name="parameters"></a>パラメーター
なし

#### <a name="returns"></a>戻り値
void

#### <a name="examples"></a>例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the content controls collection.
    var contentControls = context.document.contentControls;
    
    // Queue a command to load the content controls collection.
    contentControls.load('text');
     
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        
        if (contentControls.items.length === 0) {
            console.log("There isn't a content control in this document.");
        } else {
            
            // Queue a command to clear the contents of the first content control.
            contentControls.items[0].clear();
            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync().then(function () {
                console.log('Content control cleared of contents.');
            });      
        }
            
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});

```

### <a name="delete(keepcontent:-bool)"></a>delete(keepContent: bool)
コンテンツ コントロールとそのコンテンツを削除します。keepContent が true の場合、コンテンツは削除されません。

#### <a name="syntax"></a>構文
```js
contentControlObject.delete(keepContent);
```

#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|keepContent|bool|必須。コンテンツ コントロールを使用してコンテンツを削除する必要があるかどうかを示します。keepContent が true の場合、コンテンツは削除されません。|

#### <a name="returns"></a>戻り値
void

#### <a name="examples"></a>例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the content controls collection.
    var contentControls = context.document.contentControls;
    
    // Queue a command to load the content controls collection.
    contentControls.load('text');
     
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        
        if (contentControls.items.length === 0) {
            console.log("There isn't a content control in this document.");
        } else {
            
            // Queue a command to delete the first content control. The
            // contents will remain in the document.
            contentControls.items[0].delete(true);
            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync().then(function () {
                console.log('Content control cleared of contents.');
            });      
        }
            
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```


### <a name="gethtml()"></a>getHtml()
コンテンツ コントロール オブジェクトの HTML 表記を取得します。

#### <a name="syntax"></a>構文
```js
contentControlObject.getHtml();
```

#### <a name="parameters"></a>パラメーター
なし

#### <a name="returns"></a>戻り値
string

#### <a name="examples"></a>例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the content controls collection that contains a specific tag.
    var contentControlsWithTag = context.document.contentControls.getByTag('Customer-Address');
    
    // Queue a command to load the tag property for all of content controls. 
    context.load(contentControlsWithTag, 'tag');
     
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        if (contentControlsWithTag.items.length === 0) {
            console.log('No content control found.');
        }
        else {
            // Queue a command to get the HTML contents of the first content control.
            var html = contentControlsWithTag.items[0].getHtml();
        
            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    console.log('Content control HTML: ' + html.value);
            });
        }
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### <a name="getooxml()"></a>getOoxml()
コンテンツ コントロール オブジェクトの Office Open XML (OOXML) 表記を取得します。

#### <a name="syntax"></a>構文
```js
contentControlObject.getOoxml();
```

#### <a name="parameters"></a>パラメーター
なし

#### <a name="returns"></a>戻り値
string

#### <a name="examples"></a>例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the content controls collection.
    var contentControls = context.document.contentControls;
    
    // Queue a command to load the id property for all of the content controls. 
    context.load(contentControls, 'id');
     
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        if (contentControls.items.length === 0) {
            console.log('No content control found.');
        }
        else {
            // Queue a command to get the OOXML contents of the first content control.
            var ooxml = contentControls.items[0].getOoxml();
        
            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    console.log('Content control OOXML: ' + ooxml.value);
            });
        }
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### <a name="insertbreak(breaktype:-breaktype,-insertlocation:-insertlocation)"></a>insertBreak(breakType: BreakType, insertLocation: InsertLocation)
指定した位置に、区切りを挿入します。改行以外の区切りは、メインドキュメント本文に含まれるオブジェクトにのみ挿入できます。改行は、いずれの本文オブジェクトにも挿入できます。insertLocation の値には、'Before'、'After'、'Start'、'End' のいずれかを指定できます。

#### <a name="syntax"></a>構文
```js
contentControlObject.insertBreak(breakType, insertLocation);
```

#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|breakType|BreakType|必須。区切りの種類 (breakType.md)|
|insertLocation|InsertLocation|必須。値には、'Before'、'After'、'Start'、または 'End' を指定できます。|

#### <a name="returns"></a>戻り値
void

#### <a name="additional-details"></a>追加の詳細
ヘッダー、フッター、脚注、文末脚注、コメント、テキスト ボックスに含まれたオブジェクトに改行以外の区切りを挿入することはできません。  

#### <a name="examples"></a>例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the content controls collection.
    var contentControls = context.document.contentControls;
    
    // Queue a commmand to load the id property for all of content controls. 
    context.load(contentControls, 'id');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion. We now will have 
    // access to the content control collection.
    return context.sync().then(function () {
        if (contentControls.items.length === 0) {
            console.log('No content control found.');
        }
        else {
            // Queue a command to insert a page break after the first content control. 
            contentControls.items[0].insertBreak('page', "After");
            
            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion. 
            return context.sync()
                .then(function () {
                    console.log('Inserted a page break after the first content control.');    
            });
        }
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### <a name="insertfilefrombase64(base64file:-string,-insertlocation:-insertlocation)"></a>insertFileFromBase64(base64File: string, insertLocation:InsertLocation)
現在のコンテンツ コントロール内の指定された位置にドキュメントを挿入します。insertLocation の値には、'Replace'、'Start'、'End' のいずれかを指定できます。

#### <a name="syntax"></a>構文
```js
contentControlObject.insertFileFromBase64(base64File, insertLocation);
```

#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|base64File|string|必須。base64 でエンコードされた挿入するファイルのコンテンツ。|
|insertLocation|InsertLocation|必須。値には、'Replace'、'Start'、'End' のいずれかを指定できます。|

#### <a name="returns"></a>戻り値
[Range](range.md)

### <a name="inserthtml(html:-string,-insertlocation:-insertlocation)"></a>insertHtml(html: string, insertLocation:InsertLocation)
コンテンツ コントロール内の指定された位置に HTML を挿入します。insertLocation の値には、'Replace'、'Start'、'End' のいずれかを指定できます。

#### <a name="syntax"></a>構文
```js
contentControlObject.insertHtml(html, insertLocation);
```

#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|Html|string|必須。コンテンツ コントロールに挿入する HTML。|
|insertLocation|InsertLocation|必須。値には、'Replace'、'Start'、'End' のいずれかを指定できます。|

#### <a name="returns"></a>戻り値
[Range](range.md)

#### <a name="examples"></a>例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the content controls collection.
    var contentControls = context.document.contentControls;
    
    // Queue a command to load the id property for all of the content controls. 
    context.load(contentControls, 'id');
     
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        if (contentControls.items.length === 0) {
            console.log('No content control found.');
        }
        else {
            // Queue a command to put HTML into the contents of the first content control.
            contentControls.items[0].insertHtml('<strong>HTML content inserted into the content control.</strong>', 'Start');
        
            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    console.log('Inserted HTML in the first content control.');
            });
        }
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### <a name="insertinlinepicturefrombase64(base64encodedimage:-string,-insertlocation:-insertlocation)"></a>insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: InsertLocation)
コンテンツ コントロール内の指定された位置にインライン画像を挿入します。insertLocation の値は、'Replace'、'Start'、'End' のいずれかになります。

#### <a name="syntax"></a>構文
contentControlObject.insertInlinePictureFromBase64(image, insertLocation);

#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|base64EncodedImage|string|必須。コンテンツ コントロールに挿入される base64 でエンコードされた画像。|
|insertLocation|InsertLocation|必須。値には、'Replace'、'Start'、'End' のいずれかを指定できます。|

#### <a name="returns"></a>戻り値
[InlinePicture](inlinepicture.md)

#### <a name="known-issues"></a>既知の問題
Word オンラインでは、_insertLocation_ パラメーターに対して 'Replace' 値のみがサポートされています。'Start' または 'End' 値を使用した場合は、操作は失敗します。

### <a name="insertooxml(ooxml:-string,-insertlocation:-insertlocation)"></a>insertOoxml(ooxml: string, insertLocation:InsertLocation)
コンテンツ コントロール内の指定された位置に OOXML または wordProcessingML を挿入します。insertLocation の値には、'Replace'、'Start'、'End' のいずれかを指定できます。

#### <a name="syntax"></a>構文
```js
contentControlObject.insertOoxml(ooxml, insertLocation);
```

#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|ooxml|string|必須。コンテンツ コントロールに挿入する OOXML または wordProcessingML。|
|insertLocation|InsertLocation|必須。値には、'Replace'、'Start'、'End' のいずれかを指定できます。|

#### <a name="returns"></a>戻り値
[範囲](range.md)

#### <a name="known-issues"></a>既知の問題
このメソッドを使用すると Word オンラインの待機時間が長くなります。これはアドインのユーザー エクスペリエンスに影響を与える可能性があります。他のソリューションが利用できない場合にのみ、このメソッドを使用することをお勧めします。 

#### <a name="examples"></a>例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the content controls collection.
    var contentControls = context.document.contentControls;
    
    // Queue a command to load the id property for all of the content controls. 
    context.load(contentControls, 'id');
     
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        if (contentControls.items.length === 0) {
            console.log('No content control found.');
        }
        else {
            // Queue a command to put OOXML into the contents of the first content control.
            contentControls.items[0].insertOoxml("<pkg:package xmlns:pkg='http://schemas.microsoft.com/office/2006/xmlPackage'><pkg:part pkg:name='/_rels/.rels' pkg:contentType='application/vnd.openxmlformats-package.relationships+xml' pkg:padding='512'><pkg:xmlData><Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument' Target='word/document.xml'/></Relationships></pkg:xmlData></pkg:part><pkg:part pkg:name='/word/document.xml' pkg:contentType='application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'><pkg:xmlData><w:document xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' ><w:body><w:p><w:pPr><w:spacing w:before='360' w:after='0' w:line='480' w:lineRule='auto'/><w:rPr><w:color w:val='70AD47' w:themeColor='accent6'/><w:sz w:val='28'/></w:rPr></w:pPr><w:r><w:rPr><w:color w:val='70AD47' w:themeColor='accent6'/><w:sz w:val='28'/></w:rPr><w:t>This text has formatting directly applied to achieve its font size, color, line spacing, and paragraph spacing.</w:t></w:r></w:p></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>", "End");
        
            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    console.log('Inserted OOXML in the first content control.');
            });
        }
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

#### <a name="additional-information"></a>追加情報
OOXML の操作の詳細については、「[Office Open XML を使用して Word のより良いアドインを作成する](https://msdn.microsoft.com/en-us/library/office/dn423225.aspx)」をお読みください。

### <a name="insertparagraph(paragraphtext:-string,-insertlocation:-insertlocation)"></a>insertParagraph(paragraphText: string, insertLocation: InsertLocation)
指定した位置に、段落を挿入します。insertLocation の値には、'Before'、'After'、'Start'、'End' のいずれかを指定できます。

#### <a name="syntax"></a>構文
```js
contentControlObject.insertParagraph(paragraphText, insertLocation);
```

#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|paragraphText|string|必須。挿入する段落テキスト。|
|insertLocation|InsertLocation|必須。値には、'Before'、'After'、'Start'、または 'End' を指定できます。|

#### <a name="returns"></a>戻り値
[Paragraph](paragraph.md)

#### <a name="examples"></a>例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the content controls collection.
    var contentControls = context.document.contentControls;
    
    // Queue a command to load the id property for all of the content controls. 
    context.load(contentControls, 'id');
     
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        if (contentControls.items.length === 0) {
            console.log('No content control found.');
        }
        else {
            // Queue a command to insert a paragraph after the first content control. 
            contentControls.items[0].insertParagraph('Text of the inserted paragraph.', 'After');
        
            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    console.log('Inserted a paragraph after the first content control.');
            });
        }
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### <a name="inserttext(text:-string,-insertlocation:-insertlocation)"></a>insertText(text: string, insertLocation:InsertLocation)
コンテンツ コントロール内の指定された位置にテキストを挿入します。insertLocation の値には、'Replace'、'Start'、'End' のいずれかを指定できます。

#### <a name="syntax"></a>構文
```js
contentControlObject.insertText(text, insertLocation);
```

#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|text|string|必須。コンテンツ コントロールに挿入する テキスト。|
|insertLocation|InsertLocation|必須。値には、'Replace'、'Start'、'End' のいずれかを指定できます。|

#### <a name="returns"></a>戻り値
[範囲](range.md)

#### <a name="known-issues"></a>既知の問題
Word オンラインでは、_insertLocation_ パラメーターに対して 'Replace' 値のみがサポートされています。'Start' または 'End' 値を使用した場合は、操作は失敗します。

#### <a name="examples"></a>例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the content controls collection.
    var contentControls = context.document.contentControls;
    
    // Queue a command to load the id property for all of the content controls. 
    context.load(contentControls, 'id');
     
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        if (contentControls.items.length === 0) {
            console.log('No content control found.');
        }
        else {
            // Queue a command to replace text in the first content control. 
            contentControls.items[0].insertText('Replaced text in the first content control.', 'Replace');
        
            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    console.log('Replaced text in the first content control.');
            });
        }
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

[Silly stories](https://aka.ms/sillystorywordaddin) サンプル アドインは、**insertText** メソッドの使用方法を示しています。

### <a name="load(param:-object)"></a>load(param: object)
JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。

#### <a name="syntax"></a>構文
```js
object.load(param);
```

#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|param|object|省略可能。パラメーターとリレーションシップ名を、区切られた文字列または 1 つの配列として受け入れます。あるいは、[loadOption](loadoption.md) オブジェクトを提供します。|

#### <a name="returns"></a>戻り値
void

#### <a name="examples"></a>例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy range object for the current selection.
    var range = context.document.getSelection();
    
    // Queue a commmand to create the content control.
    var myContentControl = range.insertContentControl();
    myContentControl.tag = 'Customer-Address';
    myContentControl.title = ' has t';
    myContentControl.style = 'Heading 2';
    myContentControl.insertText('One Microsoft Way, Redmond, WA 98052', 'replace');
    myContentControl.cannotEdit = true;
    
    // Queue a command to load the id property for the content control you created.
    context.load(myContentControl, 'id');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Created content control with id: ' + myContentControl.id);
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### <a name="search(searchtext:-string,-searchoptions:-paramtypestrings.searchoptions)"></a>search(searchText: string, searchOptions:ParamTypeStrings.SearchOptions)
コンテンツ コントロール オブジェクトの範囲で、指定した searchOptions を使って検索を実行します。検索結果は、範囲オブジェクトのコレクションです。

#### <a name="syntax"></a>構文
```js
contentControlObject.search(searchText, searchOptions);
```

#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|searchText|文字列|必須。検索テキスト。|
|[searchOptions](searchoptions.md)|ParamTypeStrings.SearchOptions|省略可能。検索のオプション。|

#### <a name="returns"></a>戻り値
[SearchResultCollection](searchresultcollection.md)

### <a name="select(selectionmode:-selectionmode)"></a>select(selectionMode: SelectionMode)
コンテンツ コントロールを選択します。その結果、Word は選択範囲にスクロールされます。選択モードは、'Select'、'Start'、'End' のいずれかになります。

#### <a name="syntax"></a>構文
```js
contentControlObject.select(selectionMode);
```

#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|selectionMode|SelectionMode|省略可能。選択モードは、'Select'、'Start'、'End' のいずれかになります。'Select' が既定値です。|

#### <a name="returns"></a>戻り値
void

#### <a name="examples"></a>例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the content controls collection.
    var contentControls = context.document.contentControls;
    
    // Queue a command to load the id property for all of the content controls. 
    context.load(contentControls, 'id');
     
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        if (contentControls.items.length === 0) {
            console.log('No content control found.');
        }
        else {
            // Queue a command to select the first content control.
            contentControls.items[0].select();
        
            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    console.log('Selected the first content control.');
            });
        }
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

## <a name="property-access-examples"></a>プロパティのアクセスの例

### <a name="load-all-of-the-content-control-properties"></a>すべてのコンテンツ コントロールのプロパティを読み込む
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the content controls collection.
    var contentControls = context.document.contentControls;
    
    // Queue a command to load the id property for all of the content controls. 
    context.load(contentControls, 'id');
     
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        if (contentControls.items.length === 0) {
            console.log('No content control found.');
        }
        else {
            // Queue a command to load the properties on the first content control. 
            contentControls.items[0].load(  'appearance,' +
                                            'cannotDelete,' +
                                            'cannotEdit,' +
                                            'id,' +
                                            'placeHolderText,' +
                                            'removeWhenEdited,' +
                                            'title,' +
                                            'text,' +
                                            'type,' +
                                            'style,' +
                                            'tag,' +
                                            'font/size,' +
                                            'font/name,' +
                                            'font/color');             
        
            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    console.log('Property values of the first content control:' + 
                        '   ----- appearance: ' + contentControls.items[0].appearance + 
                        '   ----- cannotDelete: ' + contentControls.items[0].cannotDelete +
                        '   ----- cannotEdit: ' + contentControls.items[0].cannotEdit +
                        '   ----- color: ' + contentControls.items[0].color +
                        '   ----- id: ' + contentControls.items[0].id +
                        '   ----- placeHolderText: ' + contentControls.items[0].placeholderText +
                        '   ----- removeWhenEdited: ' + contentControls.items[0].removeWhenEdited +
                        '   ----- title: ' + contentControls.items[0].title +
                        '   ----- text: ' + contentControls.items[0].text +
                        '   ----- type: ' + contentControls.items[0].type +
                        '   ----- style: ' + contentControls.items[0].style +
                        '   ----- tag: ' + contentControls.items[0].tag +
                        '   ----- font size: ' + contentControls.items[0].font.size +
                        '   ----- font name: ' + contentControls.items[0].font.name +
                        '   ----- font color: ' + contentControls.items[0].font.color);
            });
        }
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

## <a name="support-details"></a>サポートの詳細
実行時のチェックで[要件セット](../office-add-in-requirement-sets.md)を使用して、アプリケーションが Word のホスト バージョンによってサポートされていることを確かめます。Office ホスト アプリケーションとサーバーの要件の詳細については、「[Office アドインを実行するための要件](../../docs/overview/requirements-for-running-office-add-ins.md)」を参照してください。
