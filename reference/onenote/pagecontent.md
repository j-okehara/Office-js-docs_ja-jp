# PageContent オブジェクト (JavaScript API for OneNote)

_適用対象:OneNote Online_  


Outline や Image などの最上位のコンテンツの種類を含むページの領域を表します。PageContent オブジェクトは、XY の位置を指定できます。

## プロパティ

| プロパティ     | 型   |説明|フィードバック|
|:---------------|:--------|:----------|:-------|
|id|string|PageContent オブジェクトの ID を取得します。 読み取り専用です。|[実行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-id)|
|left|double|PageContent オブジェクトの左 (X 軸) の位置を取得するか設定します。|[実行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-left)|
|top|double|PageContent オブジェクトの上 (Y 軸) の位置を取得するか設定します。|[実行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-top)|
|type|string|PageContent オブジェクトの種類を取得します。 読み取り専用です。 使用可能な値は次のとおりです。Outline、Image、Other。|[実行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-type)|

## リレーションシップ
| リレーションシップ | 型   |説明| フィードバック|
|:---------------|:--------|:----------|:-------|
|image|[Image](image.md)|PageContent オブジェクト内の Image を取得します。 PageContentType が Image ではない場合は例外をスローします。 読み取り専用です。|[実行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-image)|
|Ink|[FloatingInk](floatingink.md)|PageContent オブジェクト内の Ink を取得します。 PageContentType が Ink ではない場合は例外をスローします。 読み取り専用です。|[実行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-ink)|
|outline|[アウトライン](outline.md)|PageContent オブジェクト内の Outline を取得します。 PageContentType が Outline ではない場合は例外をスローします。 読み取り専用です。|[実行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-outline)|
|parentPage|[Page](page.md)|PageContent オブジェクトを含むページを取得します。 読み取り専用です。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-parentPage)|

## メソッド

| メソッド           | 戻り値の型    |説明| フィードバック|
|:---------------|:--------|:----------|:-------|
|[delete()](#delete)|void|PageContent オブジェクトを削除します。|[実行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-delete)|
|[load(param: object)](#loadparam-object)|void|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-load)|

## メソッドの詳細


### delete()
PageContent オブジェクトを削除します。

#### 構文
```js
pageContentObject.delete();
```

#### パラメーター
なし

#### 戻り値
void

#### 例
```js
OneNote.run(function (context) {

    var page = context.application.getActivePage();
    var pageContents = page.contents;

    var firstPageContent = pageContents.getItemAt(0);
    firstPageContent.load('type');

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            if(firstPageContent.isNull === false) {
                firstPageContent.delete();
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
