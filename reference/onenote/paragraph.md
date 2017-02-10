# <a name="paragraph-object-javascript-api-for-onenote"></a>Paragraph オブジェクト (JavaScript API for OneNote)

_適用対象:OneNote Online_  


ページに表示されるコンテンツのコンテナー。Paragraph に含めることができるのは、コンテンツの ParagraphType の種類のいずれか 1 つです。

## <a name="properties"></a>プロパティ

| プロパティ     | 型   |説明|フィードバック|
|:---------------|:--------|:----------|:-------|
|id|string|Paragraph オブジェクトの ID を取得します。読み取り専用です。|[移動](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-id)|
|type|string|Paragraph オブジェクトの種類を取得します。読み取り専用です。使用可能な値は次のとおりです。RichText、Image、Table、Other。|[移動](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-type)|

_プロパティのアクセスの[例を参照してください。](#property-access-examples)_

## <a name="relationships"></a>関係
| リレーションシップ | 型   |説明| フィードバック|
|:---------------|:--------|:----------|:-------|
|image|[Image](image.md)|Paragraph 内の Image オブジェクトを取得します。ParagraphType が Image ではない場合は例外をスローします。読み取り専用です。|[移動](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-image)|
|inkWords|[InkWordCollection](inkwordcollection.md)|Paragraph 内のインク コレクションを取得します。ParagraphType が Ink ではない場合は例外をスローします。読み取り専用です。|[移動](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-inkWords)|
|outline|[Outline](outline.md)|Paragraph を含む Outline オブジェクトを取得します。読み取り専用です。|[移動](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-outline)|
|paragraphs|[ParagraphCollection](paragraphcollection.md)|この段落の下にある段落のコレクション。読み取り専用です。|[移動](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-paragraphs)|
|parentParagraph|[Paragraph](paragraph.md)|親の Paragraph オブジェクトを取得します。親の Paragraph が存在しない場合はスローします。読み取り専用です。|[移動](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-parentParagraph)|
|parentParagraphOrNull|[Paragraph](paragraph.md)|親の Paragraph オブジェクトを取得します。親の Paragraph が存在しない場合は null を返します。読み取り専用です。|[移動](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-parentParagraphOrNull)|
|parentTableCell|[TableCell](tablecell.md)|Paragraph を含む TableCell オブジェクトを取得します (存在する場合)。親が TableCell でない場合は ItemNotFound をスローします。読み取り専用です。|[移動](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-parentTableCell)|
|parentTableCellOrNull|[TableCell](tablecell.md)|Paragraph を含む TableCell オブジェクトを取得します (存在する場合)。親が TableCell でない場合は null を返します。読み取り専用です。|[移動](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-parentTableCellOrNull)|
|richText|[RichText](richtext.md)|Paragraph 内の RichText オブジェクトを取得します。ParagraphType が RichText ではない場合は例外をスローします。読み取り専用です。|[移動](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-richText)|
|table|[Table](table.md)|Paragraph 内の Table オブジェクトを取得します。ParagraphType が Table ではない場合は例外をスローします。読み取り専用です。|[移動](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-table)|

## <a name="methods"></a>メソッド

| メソッド           | 戻り値の型    |説明| フィードバック|
|:---------------|:--------|:----------|:-------|
|[delete()](#delete)|void|段落を削除します。|[移動](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-delete)|
|[insertHtmlAsSibling(insertLocation: string, html: string)](#inserthtmlassiblinginsertlocation-string-html-string)|void|指定された HTML コンテンツを挿入します。|[移動](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-insertHtmlAsSibling)|
|[insertImageAsSibling(insertLocation: string, base64EncodedImage: string, width: double, height: double)](#insertimageassiblinginsertlocation-string-base64encodedimage-string-width-double-height-double)|[Image](image.md)|指定された挿入位置にイメージを挿入します。|[移動](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-insertImageAsSibling)|
|[insertRichTextAsSibling(insertLocation: string, paragraphText: string)](#insertrichtextassiblinginsertlocation-string-paragraphtext-string)|[RichText](richtext.md)|指定された挿入位置に段落テキストを挿入します。|[移動](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-insertRichTextAsSibling)|
|[insertTableAsSibling(insertLocation: string, rowCount: number, columnCount: number, values: string[][])](#inserttableassiblinginsertlocation-string-rowcount-number-columncount-number-values-string)|[Table](table.md)|指定された数の行と列を含む表を現在の段落の前または後に追加します。|[移動](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-insertTableAsSibling)|
|[load(param: object)](#loadparam-object)|(非推奨)|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|[移動](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-load)|

## <a name="method-details"></a>メソッドの詳細


### <a name="delete"></a>delete()
段落を削除します。

#### <a name="syntax"></a>構文
```js
paragraphObject.delete();
```

#### <a name="parameters"></a>パラメーター
なし

#### <a name="returns"></a>戻り値
void

#### <a name="examples"></a>例
```js
OneNote.run(function (context) {

    // Get the collection of pageContent items from the page.
    var pageContents = context.application.getActivePage().contents;

    // Get the first PageContent on the page
    // Assuming its an outline, get the outline's paragraphs.
    var pageContent = pageContents.getItemAt(0);
    
    var paragraphs = pageContent.outline.paragraphs;
    
    var firstParagraph = paragraphs.getItemAt(0);
    
    // Queue a command to load the id and type of the first paragraph
    firstParagraph.load("id,type");

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            
            // Queue a command to delete the first paragraph                 
            firstParagraph.delete();
            
            // Run the command to delete it
            return context.sync();
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```


### <a name="inserthtmlassiblinginsertlocation-string-html-string"></a>insertHtmlAsSibling(insertLocation: string, html: string)
指定された HTML コンテンツを挿入します。

#### <a name="syntax"></a>構文
```js
paragraphObject.insertHtmlAsSibling(insertLocation, html);
```

#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|insertLocation|string|現在の Paragraph を基準にした新しいコンテンツの場所です。使用可能な値は次のとおりです。Before、After。|
|Html|string|コンテンツの視覚表示を記述する HTML 文字列です。OneNote アドインの JavaScript API については、「[サポートされる HTML](../../docs/onenote/onenote-add-ins-page-content.md#supported-html)」を参照してください。|

#### <a name="returns"></a>戻り値
void

#### <a name="examples"></a>例
```js
OneNote.run(function (context) {

    // Get the collection of pageContent items from the page.
    var pageContents = context.application.getActivePage().contents;

    // Get the first PageContent on the page
    // Assuming its an outline, get the outline's paragraphs.
    var pageContent = pageContents.getItemAt(0);
    var paragraphs = pageContent.outline.paragraphs;
    var firstParagraph = paragraphs.getItemAt(0);

    // Queue a command to load the id and type of the first paragraph
    firstParagraph.load("id,type");

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {

            // Queue commands to insert before and after the first paragraph
            firstParagraph.insertHtmlAsSibling("Before", "<p>ContentBeforeFirstParagraph</p>");
            firstParagraph.insertHtmlAsSibling("After", "<p>ContentAfterFirstParagraph</p>");
            
            // Run the command to run inserts
            return context.sync();
        });
))
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```


### <a name="insertimageassiblinginsertlocation-string-base64encodedimage-string-width-double-height-double"></a>insertImageAsSibling(insertLocation: string, base64EncodedImage: string, width: double, height: double)
指定された挿入位置にイメージを挿入します。

#### <a name="syntax"></a>構文
```js
paragraphObject.insertImageAsSibling(insertLocation, base64EncodedImage, width, height);
```

#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|insertLocation|string|現在の Paragraph を基準にした表の位置。使用可能な値は次のとおりです。Before、After。|
|base64EncodedImage|string|追加する HTML 文字列。|
|width|double|省略可能。ポイント単位の幅。既定値は null で、イメージの幅が使用されます。|
|height|double|省略可能。ポイント単位の高さ。既定値は null で、イメージの高さが使用されます。|

#### <a name="returns"></a>戻り値
[Image](image.md)

#### <a name="examples"></a>例
```js
OneNote.run(function (context) {

    // Get the collection of pageContent items from the page.
    var pageContents = context.application.getActivePage().contents;

    // Get the first PageContent on the page
    // Assuming its an outline, get the outline's paragraphs.
    var pageContent = pageContents.getItemAt(0);
    var paragraphs = pageContent.outline.paragraphs;
    var firstParagraph = paragraphs.getItemAt(0);

    // Queue a command to load the id and type of the first paragraph
    firstParagraph.load("id,type");

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {

            // Queue commands to insert before and after the first paragraph
            firstParagraph.insertImageAsSibling("Before", "R0lGODlhDwAPAKECAAAAzMzM/////wAAACwAAAAADwAPAAACIISPeQHsrZ5ModrLlN48CXF8m2iQ3YmmKqVlRtW4MLwWACH+H09wdGltaXplZCBieSBVbGVhZCBTbWFydFNhdmVyIQAAOw==");
            firstParagraph.insertImageAsSibling("After", "R0lGODlhDwAPAKECAAAAzMzM/////wAAACwAAAAADwAPAAACIISPeQHsrZ5ModrLlN48CXF8m2iQ3YmmKqVlRtW4MLwWACH+H09wdGltaXplZCBieSBVbGVhZCBTbWFydFNhdmVyIQAAOw==");
            
            // Run the command to insert images
            return context.sync();
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```


### <a name="insertrichtextassiblinginsertlocation-string-paragraphtext-string"></a>insertRichTextAsSibling(insertLocation: string, paragraphText: string)
指定された挿入位置に段落テキストを挿入します。

#### <a name="syntax"></a>構文
```js
paragraphObject.insertRichTextAsSibling(insertLocation, paragraphText);
```

#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|insertLocation|string|現在の Paragraph を基準にした表の位置。使用可能な値は次のとおりです。Before、After。|
|paragraphText|string|追加する HTML 文字列です。|

#### <a name="returns"></a>戻り値
[RichText](richtext.md)

#### <a name="examples"></a>例
```js
OneNote.run(function (context) {

    // Get the collection of pageContent items from the page.
    var pageContents = context.application.getActivePage().contents;

    // Get the first PageContent on the page
    // Assuming its an outline, get the outline's paragraphs.
    var pageContent = pageContents.getItemAt(0);
    var paragraphs = pageContent.outline.paragraphs;
    var firstParagraph = paragraphs.getItemAt(0);

    // Queue a command to load the id and type of the first paragraph
    firstParagraph.load("id,type");

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {

            // Queue commands to insert before and after the first paragraph
            firstParagraph.insertRichTextAsSibling("Before", "Text Appears Before Paragraph");
            firstParagraph.insertRichTextAsSibling("After", "Text Appears After Paragraph");
            
            // Run the command to insert text contents
            return context.sync();
        });
})  
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
}); 
```


### <a name="inserttableassiblinginsertlocation-string-rowcount-number-columncount-number-values-string"></a>insertTableAsSibling(insertLocation: string, rowCount: number, columnCount: number, values: string[][])
指定された数の行と列を含む表を現在の段落の前または後に追加します。

#### <a name="syntax"></a>構文
```js
paragraphObject.insertTableAsSibling(insertLocation, rowCount, columnCount, values);
```

#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|insertLocation|string|現在の Paragraph を基準にした表の位置。使用可能な値は次のとおりです。Before、After。|
|rowCount|number|表の行数。|
|columnCount|number|表の列数を指定します。|
|values|string[][]|省略可能。省略可能な 2 次元配列。対応する文字列が配列で指定されている場合、セルに設定されます。|

#### <a name="returns"></a>戻り値
[Table](table.md)

### <a name="loadparam-object"></a>load(param: object)
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
### <a name="property-access-examples"></a>プロパティのアクセスの例

**id と type**
```js
OneNote.run(function (context) {

    // Get the collection of pageContent items from the page.
    var pageContents = context.application.getActivePage().contents;
    
    // Queue a command to load the outline property of each pageContent.
    pageContents.load("outline");
        
    // Get the first PageContent on the page, and then get its Outline.
    var pageContent = pageContents._GetItem(0);
    var paragraphs = pageContent.outline.paragraphs;
            
    // Queue a command to load the id and type of each paragraph.
    paragraphs.load("id,type");
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            
            // Write the text.                  
            $.each(paragraphs.items, function(index, paragraph) {
                console.log("Paragraph type: " + paragraph.type);
                console.log("Paragraph ID: " + paragraph.id);
            });
        });
})      
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
}); 
```

**paragraphs**
```js
OneNote.run(function(context) {
    var app = context.application;
    
    // Gets the active outline
    var outline = app.getActiveOutline();
    
    // load nested paragraphs and their types.
    outline.load("paragraphs/type");
    
    return context.sync().then(function () {
        var paragraphs = outline.paragraphs.items;
        
        var promise;
        // for each nested paragraphs, load tables only
        for (var i = 0; i < paragraphs.length; i++) {
            var paragraph = paragraphs[i];
            if (paragraph.type == "Table") {
                paragraph.load("table/id");
                promise =  context.sync().then(function() {
                    console.log(paragraph.table.id);
                });
            }
        }
        return promise;
    })
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

