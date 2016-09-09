# Outline オブジェクト (JavaScript API for OneNote)

_適用対象:OneNote Online_  


Paragraph オブジェクトのコンテナーを表します。

## プロパティ

| プロパティ     | 型   |説明|フィードバック|
|:---------------|:--------|:----------|:-------|
|id|string|Outline オブジェクトの ID を取得します。 読み取り専用です。|[移動](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-outline-id)|

## リレーションシップ
| リレーションシップ | 型   |説明| フィードバック|
|:---------------|:--------|:----------|:-------|
|pageContent|[PageContent](pagecontent.md)|Outline を含む PageContent オブジェクトを取得します。 このオブジェクトは、ページ上の Outline の位置を定義します。 読み取り専用です。|[移動](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-outline-pageContent)|
|paragraphs|[ParagraphCollection](paragraphcollection.md)|Outline に含まれる Paragraph オブジェクトのコレクションを取得します。 読み取り専用です。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-outline-paragraphs)|

## メソッド

| メソッド           | 戻り値の型    |説明| フィードバック|
|:---------------|:--------|:----------|:-------|
|[appendHtml(html: string)](#appendhtmlhtml-string)|void|指定された HTML を Outline の一番下に追加します。|[移動](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-outline-appendHtml)|
|[appendImage(base64EncodedImage: string, width: double, height: double)](#appendimagebase64encodedimage-string-width-double-height-double)|[Image](image.md)|指定されたイメージを Outline の一番下に追加します。|[移動](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-outline-appendImage)|
|[appendRichText(paragraphText: string)](#appendrichtextparagraphtext-string)|[RichText](richtext.md)|指定されたテキストを Outline の一番下に追加します。|[移動](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-outline-appendRichText)|
|[appendTable(rowCount: number, columnCount: number, values: string[][])](#appendtablerowcount-number-columncount-number-values-string)|[Table](table.md)|指定された数の行と列を含むテーブルを Outline の一番下に追加します。|[移動](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-outline-appendTable)|
|[load(param: object)](#loadparam-object)|void|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-outline-load)|

## メソッドの詳細


### appendHtml(html: string)
指定された HTML を Outline の一番下に追加します。

#### 構文
```js
outlineObject.appendHtml(html);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|Html|string|追加する HTML 文字列です。 OneNote アドインの JavaScript API については、「[サポートされる HTML](../../docs/onenote/onenote-add-ins-page-content.md#supported-html)」を参照してください。|

#### 戻り値
void

#### 例
```js
OneNote.run(function (context) {

    // Gets the active page.
    var activePage = context.application.getActivePage();

    // Get pageContents of the activePage. 
    var pageContents = activePage.contents;

    // Queue a command to load the pageContents to access its data.
    context.load(pageContents);

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function() {
            if (pageContents.items.length != 0 && pageContents.items[0].type == "Outline")
            {
                // First item is an outline.
                outline = pageContents.items[0].outline;

                // Queue a command to append a paragraph to the outline.
                outline.appendHtml("<p>new paragraph</p>");

                // Run the queued commands.
                return context.sync();
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


### appendImage(base64EncodedImage: string, width: double, height: double)
指定されたイメージを Outline の一番下に追加します。

#### 構文
```js
outlineObject.appendImage(base64EncodedImage, width, height);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|base64EncodedImage|string|追加する HTML 文字列。|
|width|double|省略可能。 ポイント単位の幅。 既定値は null で、イメージの幅が使用されます。|
|height|double|省略可能。 ポイント単位の高さ。 既定値は null で、イメージの高さが使用されます。|

#### 戻り値
[Image](image.md)

### appendRichText(paragraphText: string)
指定されたテキストを Outline の一番下に追加します。

#### 構文
```js
outlineObject.appendRichText(paragraphText);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|paragraphText|string|追加する HTML 文字列です。|

#### 戻り値
[RichText](richtext.md)

### appendTable(rowCount: number, columnCount: number, values: string[][])
指定された数の行と列を含むテーブルを Outline の一番下に追加します。

#### 構文
```js
outlineObject.appendTable(rowCount, columnCount, values);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|rowCount|number|必須。 表の行数。|
|columnCount|number|必須。 表の列数。|
|values|string[][]|省略可能。 省略可能な 2 次元配列。 対応する文字列が配列で指定されている場合、セルに設定されます。|

#### 戻り値
[テーブル](table.md)

#### 例
```js
OneNote.run(function (context) {

    // Gets the active page.
    var activePage = context.application.getActivePage();

    // Get pageContents of the activePage. 
    var pageContents = activePage.contents;

    // Queue a command to load the pageContents to access its data.
    context.load(pageContents);

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function() {
            if (pageContents.items.length != 0 && pageContents.items[0].type == "Outline") {
                // First item is an outline.
                var outline = pageContents.items[0].outline;

                // Queue a command to append a paragraph to the outline.
                outline.appendTable(2, 2, [[1, 2],[3, 4]]);

                // Run the queued commands.
                return context.sync();
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
