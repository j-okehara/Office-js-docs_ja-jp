# InkAnalysisParagraph オブジェクト (JavaScript API for OneNote)

_適用対象:OneNote Online_  


インク ストロークによって形成された識別済み段落のインクの解析データを表します。

## プロパティ

| プロパティ     | 型   |説明|フィードバック|
|:---------------|:--------|:----------|:-------|
|id|string|InkAnalysisParagraph オブジェクトの ID を取得します。 読み取り専用です。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisParagraph-id)|

_プロパティのアクセスの[例](#例)を参照してください。_

## リレーションシップ
| リレーションシップ | 型   |説明| フィードバック|
|:---------------|:--------|:----------|:-------|
|inkAnalysis|[InkAnalysis](inkanalysis.md)|親の InkAnalysisPage を参照します。 読み取り専用です。|[実行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisParagraph-inkAnalysis)|
|lines|[InkAnalysisLineCollection](inkanalysislinecollection.md)|このインクの解析段落におけるインクの解析行を取得します。 読み取り専用です。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisParagraph-lines)|

## メソッド

| メソッド           | 戻り値の型    |説明| フィードバック|
|:---------------|:--------|:----------|:-------|
|[load(param: object)](#loadparam-object)|void|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisParagraph-load)|

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

**lines**
```js
OneNote.run(function (ctx) {        
    var app = ctx.application;
    
    // Gets the active page.
    var page = app.getActivePage();
    
    // Load a line of ink words.
    page.load('inkAnalysisOrNull/paragraphs/lines');
    
    return ctx.sync()
        .then(function() {
            var inkParagraphs = page.inkAnalysisOrNull.paragraphs;
            
            // Log id of each line in ink paragraphs.
            $.each(inkParagraphs.items, function(i, inkParagraph){
                var inkLines = inkParagraph.lines;
                $.each(inkLines.items, function (j, inkLine) {
                    console.log(inkLine.id);
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