# InkAnalysisWord オブジェクト (JavaScript API for OneNote)

_適用対象:OneNote Online_  


インク ストロークによって形成された識別済み単語のインクの解析データを表します。

## プロパティ

| プロパティ     | 型   |説明|フィードバック|
|:---------------|:--------|:----------|:-------|
|id|string|InkAnalysisWord オブジェクトの ID を取得します。 読み取り専用です。|[実行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisWord-id)|
|languageId|string|この inkAnalysisWord で認識された言語の id。 読み取り専用です。|[実行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisWord-languageId)|
|wordAlternates|string|このインクの単語で認識された単語を可能性の高い順に並べたもの。 読み取り専用です。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisWord-wordAlternates)|

_プロパティのアクセスの[例](#例)を参照してください。_

## リレーションシップ
| リレーションシップ | 型   |説明| フィードバック|
|:---------------|:--------|:----------|:-------|
|line|[InkAnalysisLine](inkanalysisline.md)|親の InkAnalysisLine を参照します。 読み取り専用です。|[実行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisWord-line)|
|strokePointers|[InkStrokePointer](inkstrokepointer.md)|このインクの解析単語の一部として認識されたインク ストロークへの弱い参照。 読み取り専用です。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisWord-strokePointers)|

## メソッド

| メソッド           | 戻り値の型    |説明| フィードバック|
|:---------------|:--------|:----------|:-------|
|[load(param: object)](#loadparam-object)|void|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisWord-load)|

## メソッドの詳細


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

**wordAlternates と languageId**
```js
OneNote.run(function (ctx) {        
    var app = ctx.application;
    
    // Gets the active page.
    var page = app.getActivePage();
    
    page.load('inkAnalysisOrNull/paragraphs/lines/words');
    return ctx.sync()
        .then(function() {
            var inkParagraphs = page.inkAnalysisOrNull.paragraphs;
            $.each(inkParagraphs.items, function(i, inkParagraph) {
                var inkLines = inkParagraph.lines;
                $.each(inkLines.items, function(j, inkLine) {
                    var inkWords = inkLine.words;
                    $.each(inkWords.items, function(k, inkWord) {
                    
                        // Log language Id of the word
                        console.log(inkWord.languageId);
                        
                        // Log every ink analyzed words.
                        $.each(inkWord.wordAlternates, function(l, word) {
                            console.log(word);                                  
                        })
                    })
                })
            })
        })
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
}); 
```