# <a name="pagecontent-object-(javascript-api-for-onenote)"></a>PageContent オブジェクト (JavaScript API for OneNote)

_適用対象:OneNote Online_  


Outline や Image などの最上位のコンテンツの種類を含むページの領域を表します。PageContent オブジェクトは、XY の位置を指定できます。

## <a name="properties"></a>プロパティ

| プロパティ     | 型   |説明|フィードバック|
|:---------------|:--------|:----------|:-------|
|id|string|PageContent オブジェクトの ID を取得します。読み取り専用です。|[移動](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-id)|
|left|double|PageContent オブジェクトの左 (X 軸) の位置を取得するか設定します。|[移動](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-left)|
|top|double|PageContent オブジェクトの上 (Y 軸) の位置を取得するか設定します。|[移動](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-top)|
|type|string|PageContent オブジェクトの種類を取得します。読み取り専用です。使用可能な値は次のとおりです。Outline、Image、Other。|[移動](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-type)|

## <a name="relationships"></a>関係
| リレーションシップ | 型   |説明| フィードバック|
|:---------------|:--------|:----------|:-------|
|image|[Image](image.md)|PageContent オブジェクト内の Image を取得します。PageContentType が Image ではない場合は例外をスローします。読み取り専用です。|[移動](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-image)|
|ink|[FloatingInk](floatingink.md)|PageContent オブジェクト内の Ink を取得します。PageContentType が Ink ではない場合は例外をスローします。読み取り専用です。|[移動](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-ink)|
|outline|[Outline](outline.md)|PageContent オブジェクト内の Outline を取得します。PageContentType が Outline ではない場合は例外をスローします。読み取り専用です。|[移動](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-outline)|
|parentPage|[Page](page.md)|PageContent オブジェクトを含むページを取得します。読み取り専用です。|[移動](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-parentPage)|

## <a name="methods"></a>メソッド

| メソッド           | 戻り値の型    |説明| フィードバック|
|:---------------|:--------|:----------|:-------|
|[delete()](#delete)|void|PageContent オブジェクトを削除します。|[移動](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-delete)|
|[load(param: object)](#loadparam-object)|(非推奨)|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|[移動](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-load)|

## <a name="method-details"></a>メソッドの詳細


### <a name="delete()"></a>delete()
PageContent オブジェクトを削除します。

#### <a name="syntax"></a>構文
```js
pageContentObject.delete();
```

#### <a name="parameters"></a>パラメーター
なし

#### <a name="returns"></a>戻り値
void

#### <a name="examples"></a>例
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
