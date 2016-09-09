# Image オブジェクト (JavaScript API for OneNote)

_適用対象:OneNote Online_  


イメージを表します。 Image は、PageContent オブジェクトまたは Paragraph オブジェクトの直接の子にすることができます。

## プロパティ

| プロパティ     | 型   |説明|フィードバック|
|:---------------|:--------|:----------|:-------|
|description|string|Image の説明を取得または設定します。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-image-description)|
|height|double|Image レイアウトの高さを取得または設定します。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-image-height)|
|hyperlink|string|Image のハイパーリンクを取得または設定します。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-image-hyperlink)|
|id|string|Image オブジェクトの ID を取得します。 読み取り専用です。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-image-id)|
|width|double|Image レイアウトの幅を取得または設定します。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-image-width)|

_プロパティのアクセスの[例](#例)を参照してください。_

## リレーションシップ
| リレーションシップ | 型   |説明| フィードバック|
|:---------------|:--------|:----------|:-------|
|ocrData|[ImageOcrData](imageocrdata.md)|OCR テキストや言語など、OCR (光学式文字認識) で取得されたこの画像のデータを取得します。 読み取り専用です。|[実行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-image-ocrData)|
|pageContent|[ページ コンテンツ](pagecontent.md)|Image を含む PageContent オブジェクトを取得します。 Image が PageContent の直接の子ではない場合はスローします。 このオブジェクトは、ページの Image の位置を定義します。 読み取り専用です。|[実行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-image-pageContent)|
|paragraph|[段落](paragraph.md)|Image を含む Paragraph オブジェクトを取得します。 Image が Paragraph の直接の子ではない場合はスローします。 読み取り専用です。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-image-paragraph)|

## メソッド

| メソッド           | 戻り値の型    |説明| フィードバック|
|:---------------|:--------|:----------|:-------|
|[getBase64Image()](#getbase64image)|string|Image の Base64 エンコードのバイナリ形式を取得します。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-image-getBase64Image)|
|[load(param: object)](#loadparam-object)|void|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-image-load)|

## メソッドの詳細


### getBase64Image()
Image の Base64 エンコードのバイナリ形式を取得します。

#### 構文
```js
imageObject.getBase64Image();
```

#### パラメーター
なし

#### 戻り値
string

#### 例
```js

var image = null;
var imageString;

OneNote.run(function(ctx){
    // Get the current outline.         
    var outline = ctx.application.getActiveOutline();
    
    // Queue a command to load paragraphs and their types. 
    outline.load("paragraphs/type")
    return ctx.sync().
        then(function(){
            for (var i=0; i < outline.paragraphs.items.length; i++)
            {
                var paragraph = outline.paragraphs.items[i];
                if (paragraph.type == "Image")
                {
                    image = paragraph.image;
                }
            }
        })
        .then(function(){
            if (image != null)
            {
                imageString = image.getBase64Image();
                return ctx.sync();
            }
        })
        .then(function(){
            console.log(imageString);
        });
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
**id、width、height、description、hyperlink**
```js
OneNote.run(function(ctx){
    // Get the current outline.         
    var outline = ctx.application.getActiveOutline();
    var image = null;
    
    // Queue a command to load paragraphs and their types. 
    outline.load("paragraphs/type")
    return ctx.sync().
        then(function(){
            for (var i=0; i < outline.paragraphs.items.length; i++)
            {
                var paragraph = outline.paragraphs.items[i];
                if (paragraph.type == "Image")
                {
                    image = paragraph.image;
                }
            }
        })
        .then(function(){
            if (image != null)
            {
                // load every properties and relationships
                ctx.load(image);
                return ctx.sync();
            }
        })
        .then(function(){
            if (image != null)
            {                   
                console.log("image " + image.id + " width is " + image.width + " height is " + image.height);
                console.log("description: " + image.description);                   
                console.log("hyperlink: " + image.hyperlink);
            }
        });
});
```

**ocrData**
```js
var image = null;

OneNote.run(function(ctx){
    // Get the current outline.
    var outline = ctx.application.getActiveOutline();

    // Queue a command to load paragraphs and their types.
    outline.load("paragraphs")
    return ctx.sync().
        then(function(){
            for (var i=0; i < outline.paragraphs.items.length; i++)
            {
                var paragraph = outline.paragraphs.items[i];
                if (paragraph.type == "Image")
                {
                    image = paragraph.image;
                }
            }
            if (image != null)
            {
               image.load("ocrData");
            }
            return ctx.sync();
        })
        .then(function(){
            console.log(image.ocrData);
        });
}).catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

**paragraph**
```js
OneNote.run(function(ctx){
    // Get the current outline.         
    var outline = ctx.application.getActiveOutline();
    var searchedParagraph = null;
    
    // Queue a command to load paragraphs and their types. 
    outline.load("paragraphs/type")
    return ctx.sync().
        then(function() {
            for (var i=0; i < outline.paragraphs.items.length; i++)
            {
                var paragraph = outline.paragraphs.items[i];
                if (paragraph.type == "Image")
                {
                    searchedParagraph = paragraph;
                    break;
                }
            }
        })
        .then(function() {
            if (searchedParagraph != null)
            {
                // load every properties and relationships
                searchedParagraph.image.load('paragraph');
                return ctx.sync();
            }
        })
        .then(function() {
            if (searchedParagraph != null)
            {                   
                if (searchedParagraph.id != searchedParagraph.image.paragraph.id)
                {
                    console.log("id must match");
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

