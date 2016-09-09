# Page オブジェクト (JavaScript API for OneNote)

_適用対象:OneNote Online_   


OneNote ページを表します。

## プロパティ

| プロパティ     | 型   |説明|フィードバック|
|:---------------|:--------|:----------|:-------|
|clientUrl|string|ページのクライアント URL。 読み取り専用です。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-page-clientUrl)|
|id|string|ページの ID を取得します。 読み取り専用です。|[実行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-page-id)|
|pageLevel|int|ページのインデント レベルを取得または設定します。|[実行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-page-pageLevel)|
|title|string|ページのタイトルを取得または設定します。|[実行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-page-title)|
|webUrl|string|ページの Web URL。 読み取り専用です。|[実行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-page-webUrl)|

_プロパティのアクセスの[例](#例)を参照してください。_

## リレーションシップ
| リレーションシップ | 型   |説明| フィードバック|
|:---------------|:--------|:----------|:-------|
|contents|[PageContentCollection](pagecontentcollection.md)|ページの PageContent オブジェクトのコレクション。 読み取り専用です。|[実行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-page-contents)|
|inkAnalysisOrNull|[InkAnalysis](inkanalysis.md)|ページ上のインクのテキスト解釈。 インクの解析情報がない場合は null を返します。 読み取り専用です。 読み取り専用です。|[実行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-page-inkAnalysisOrNull)|
|parentSection|[Section](section.md)|ページを含むセクションを取得します。 読み取り専用です。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-page-parentSection)|

## メソッド

| メソッド           | 戻り値の型    |説明| フィードバック|
|:---------------|:--------|:----------|:-------|
|[addOutline(left: double, top: double, html:String)](#addoutlineleft-double-top-double-html-string)|[アウトライン](outline.md)|Outline をページの指定した位置に追加します。|[実行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-page-addOutline)|
|[copyToSection(destinationSection:Section)](#copytosectiondestinationsection-section)|[Page](page.md)|このページを指定したセクションにコピーします。|[実行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-page-copyToSection)|
|[insertPageAsSibling(location: string, title: string)](#insertpageassiblinglocation-string-title-string)|[Page](page.md)|現在のページの前か後に、新しいページを挿入します。|[実行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-page-insertPageAsSibling)|
|[load(param: object)](#loadparam-object)|void|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-page-load)|

## メソッドの詳細


### addOutline(left: double, top: double, html:String)
Outline をページの指定した位置に追加します。

#### 構文
```js
pageObject.addOutline(left, top, html);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|left|double|Outline の左上隅の左の位置です。|
|top|double|Outline の左上隅の上の位置です。|
|html|String|Outline の視覚表示を記述する HTML 文字列。 OneNote アドインの JavaScript API については、「[サポートされる HTML](../../docs/onenote/onenote-add-ins-page-content.md#supported-html)」を参照してください。|

#### 戻り値
[アウトライン](outline.md)

#### 例
```js
OneNote.run(function (context) {

    // Gets the active page.
    var page = context.application.getActivePage();

    // Queue a command to add an outline with given html. 
    var outline = page.addOutline(200, 200,
"<p>Images and a table below:</p> \
 <img src=\"data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAUAAAAFCAYAAACNbyblAAAAHElEQVQI12P4//8/w38GIAXDIBKE0DHxgljNBAAO9TXL0Y4OHwAAAABJRU5ErkJggg==\"> \
 <img src=\"http://imagenes.es.sftcdn.net/es/scrn/6653000/6653659/microsoft-onenote-2013-01-535x535.png\"> \
 <table> \
   <tr> \
     <td>Jill</td> \
     <td>Smith</td> \
     <td>50</td> \
   </tr> \
   <tr> \
     <td>Eve</td> \
     <td>Jackson</td> \
     <td>94</td> \
   </tr> \
 </table>"     
        );

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .catch(function(error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
});
```


### copyToSection(destinationSection:Section)
このページを指定したセクションにコピーします。

#### 構文
```js
pageObject.copyToSection(destinationSection);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|destinationSection|Section|このページのコピー先のセクション。|

#### 戻り値
[Page](page.md)

#### 例
```js
OneNote.run(function(ctx) {
    var app = ctx.application;
    
    // Gets the active notebook.
    var notebook = app.getActiveNotebook();
    
    // Gets the active page.
    var page = app.getActivePage();
    
    // Queue a command to load sections under the notebook.
    notebook.load('sections');
    
    var newPage;
    
    // Run the queued commands, and return a promise to indicate task completion.
    return ctx.sync()
        .then(function() {
            var section = notebook.sections.items[0];
            
            // copy page to the section.
            newPage = page.copyToSection(section);
            newPage.load('id');
            return ctx.sync();
        })
        .then(function() {
            console.log(newPage.id);
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

### insertPageAsSibling(location: string, title: string)
現在のページの前か後に、新しいページを挿入します。

#### 構文
```js
pageObject.insertPageAsSibling(location, title);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|location|string|現在のページを基準にした新しいページの場所です。使用可能な値は次のとおりです。Before、After。|
|title|string|新しいページのタイトルです。|

#### 戻り値
[Page](page.md)

#### 例
```js
OneNote.run(function (context) {

    // Gets the active page.
    var activePage = context.application.getActivePage();

    // Queue a command to add a new page after the active page. 
    var newPage = activePage.insertPageAsSibling("After", "Next Page");

    // Queue a command to load the newPage to access its data.
    context.load(newPage);

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function() {
            console.log("page is created with title: " + newPage.title);
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
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
### プロパティのアクセスの例

**contents**
```js
OneNote.run(function (context) {

    // Gets the active page.
    var activePage = context.application.getActivePage();

    // Queue a command to add a new page after the active page. 
    var pageContents = activePage.contents;

    // Queue a command to load the pageContents to access its data.
    context.load(pageContents);

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function() {
            for(var i=0; i < pageContents.items.length; i++)
            {
                var pageContent = pageContents.items[i];
                if (pageContent.type == "Outline")
                {
                    console.log("Found an outline");
                }
                else if (pageContent.type == "Image")
                {
                    console.log("Found an image");
                }
                else if (pageContent.type == "Other")
                {
                    console.log("Found a type not supported yet.");
                }
            }
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

**webUrl**
```js
OneNote.run(function (context) {

    var app = context.application;
    
    // Gets the active page.
    var page = app.getActivePage();
    
    // Queue a command to load the webUrl of the page.
    page.load("webUrl");
    
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function() {
            console.log(page.webUrl);
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

**inkAnalysisOrNull**
```js
OneNote.run(function (ctx) {        
    var app = ctx.application;
    
    // Gets the active page.
    var page = app.getActivePage();
    
    // Load ink words
    page.load('inkAnalysisOrNull/paragraphs/lines/words');
    
    return ctx.sync()
        .then(function() {
            if (!page.inkAnalysisOrNull.isNull)
                console.log(page.inkAnalysisOrNull.paragraphs.length);
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

