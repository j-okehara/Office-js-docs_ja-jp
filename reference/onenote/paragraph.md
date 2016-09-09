# Paragraph オブジェクト (JavaScript API for OneNote)

_適用対象:OneNote Online_  


ページに表示されるコンテンツのコンテナー。 Paragraph に含めることができるのは、コンテンツの ParagraphType の種類のいずれか 1 つです。

## プロパティ

| プロパティ     | 型   |説明|フィードバック|
|:---------------|:--------|:----------|:-------|
|id|string|Paragraph オブジェクトの ID を取得します。 読み取り専用です。|[実行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-id)|
|type|string|Paragraph オブジェクトの種類を取得します。 読み取り専用です。 使用可能な値は次のとおりです。RichText、Image、Table、Other。|[実行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-type)|

_プロパティのアクセスの[例](#例)を参照してください。_

## リレーションシップ
| リレーションシップ | 型   |説明| フィードバック|
|:---------------|:--------|:----------|:-------|
|image|[Image](image.md)|Paragraph 内の Image オブジェクトを取得します。 ParagraphType が Image ではない場合は例外をスローします。 読み取り専用です。|[実行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-image)|
|inkWords|[InkWordCollection](inkwordcollection.md)|Paragraph 内のインク コレクションを取得します。 ParagraphType が Ink ではない場合は例外をスローします。 読み取り専用です。|[実行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-inkWords)|
|outline|[Outline](outline.md)|Paragraph を含む Outline オブジェクトを取得します。 読み取り専用です。|[移動](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-outline)|
|paragraphs|[ParagraphCollection](paragraphcollection.md)|この段落の下にある段落のコレクション。 読み取り専用です。|[実行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-paragraphs)|
|parentParagraph|[Paragraph](paragraph.md)|親の Paragraph オブジェクトを取得します。 親の Paragraph が存在しない場合はスローします。 読み取り専用です。|[実行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-parentParagraph)|
|parentParagraphOrNull|[Paragraph](paragraph.md)|親の Paragraph オブジェクトを取得します。 親の Paragraph が存在しない場合は null を返します。 読み取り専用です。|[実行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-parentParagraphOrNull)|
|parentTableCell|[TableCell](tablecell.md)|Paragraph を含む TableCell オブジェクトを取得します (存在する場合)。 親が TableCell でない場合は ItemNotFound をスローします。 読み取り専用です。|[実行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-parentTableCell)|
|parentTableCellOrNull|[TableCell](tablecell.md)|Paragraph を含む TableCell オブジェクトを取得します (存在する場合)。 親が TableCell でない場合は null を返します。 読み取り専用です。|[実行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-parentTableCellOrNull)|
|richText|[RichText](richtext.md)|Paragraph 内の RichText オブジェクトを取得します。 ParagraphType が RichText ではない場合は例外をスローします。 読み取り専用です。|[実行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-richText)|
|table|[Table](table.md)|Paragraph 内の Table オブジェクトを取得します。 ParagraphType が Table ではない場合は例外をスローします。 読み取り専用です。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-table)|

## メソッド

| メソッド           | 戻り値の型    |説明| フィードバック|
|:---------------|:--------|:----------|:-------|
|[delete()](#delete)|void|段落を削除します。|[実行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-delete)|
|[insertHtmlAsSibling(insertLocation: string, html: string)](#inserthtmlassiblinginsertlocation-string-html-string)|void|指定された HTML コンテンツを挿入します。|[実行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-insertHtmlAsSibling)|
|[insertImageAsSibling(insertLocation: string, base64EncodedImage: string, width: double, height: double)](#insertimageassiblinginsertlocation-string-base64encodedimage-string-width-double-height-double)|[Image](image.md)|指定された挿入位置にイメージを挿入します。|[実行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-insertImageAsSibling)|
|[insertRichTextAsSibling(insertLocation: string, paragraphText: string)](#insertrichtextassiblinginsertlocation-string-paragraphtext-string)|[RichText](richtext.md)|指定された挿入位置に段落テキストを挿入します。|[実行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-insertRichTextAsSibling)|
|[insertTableAsSibling(insertLocation: string, rowCount: number, columnCount: number, values: string[][])](#inserttableassiblinginsertlocation-string-rowcount-number-columncount-number-values-string)|[Table](table.md)|指定された数の行と列を含む表を現在の段落の前または後に追加します。|[実行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-insertTableAsSibling)|
|[load(param: object)](#loadparam-object)|void|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-load)|

## メソッドの詳細


### delete()
段落を削除します。

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


### insertHtmlAsSibling(insertLocation: string, html: string)
指定された HTML コンテンツを挿入します。

#### 構文
```js
paragraphObject.insertHtmlAsSibling(insertLocation, html);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|insertLocation|string|現在の Paragraph を基準にした新しいコンテンツの場所です。  使用可能な値は次のとおりです。Before、After。|
|html|string|コンテンツの視覚表示を記述する HTML 文字列です。 OneNote アドインの JavaScript API については、「[サポートされる HTML](../../docs/onenote/onenote-add-ins-page-content.md#supported-html)」を参照してください。|

#### 戻り値
void

#### 例
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


### insertImageAsSibling(insertLocation: string, base64EncodedImage: string, width: double, height: double)
指定された挿入位置にイメージを挿入します。

#### 構文
```js
paragraphObject.insertImageAsSibling(insertLocation, base64EncodedImage, width, height);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|insertLocation|string|現在の Paragraph を基準にした表の位置。  使用可能な値は次のとおりです。Before、After。|
|base64EncodedImage|string|追加する HTML 文字列。|
|width|double|省略可能。 ポイント単位の幅。 既定値は null で、イメージの幅が使用されます。|
|height|double|省略可能。 ポイント単位の高さ。 既定値は null で、イメージの高さが使用されます。|

#### 戻り値
[Image](image.md)

#### 例
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


### insertRichTextAsSibling(insertLocation: string, paragraphText: string)
指定された挿入位置に段落テキストを挿入します。

#### 構文
```js
paragraphObject.insertRichTextAsSibling(insertLocation, paragraphText);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|insertLocation|string|現在の Paragraph を基準にした表の位置。  使用可能な値は次のとおりです。Before、After。|
|paragraphText|string|追加する HTML 文字列です。|

#### 戻り値
[RichText](richtext.md)

#### 例
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


### insertTableAsSibling(insertLocation: string, rowCount: number, columnCount: number, values: string[][])
指定された数の行と列を含む表を現在の段落の前または後に追加します。

#### 構文
```js
paragraphObject.insertTableAsSibling(insertLocation, rowCount, columnCount, values);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|insertLocation|string|現在の Paragraph を基準にした表の位置。  使用可能な値は次のとおりです。Before、After。|
|rowCount|number|表の行数。|
|columnCount|number|表の列数を指定します。|
|values|string[][]|省略可能。 省略可能な 2 次元配列。 対応する文字列が配列で指定されている場合、セルに設定されます。|

#### 戻り値
[Table](table.md)

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
### プロパティのアクセスの例

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

