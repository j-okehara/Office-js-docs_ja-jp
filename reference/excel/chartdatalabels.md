# ChartDataLabels オブジェクト (JavaScript API for Excel)

グラフのポイントにあるすべてのデータ ラベルのコレクションを表します。

## プロパティ

| プロパティ     | 型   |説明
|:---------------|:--------|:----------|
|position|string|データ ラベルの位置を表すDataLabelPosition 値。使用可能な値は次のとおりです。None、Center、InsideEnd、InsideBase、OutsideEnd、Left、Right、Top、Bottom、BestFit、Callout。値の設定のみ可能です。|
|Separator|string|グラフのデータ ラベルに使用される区切り文字を表す文字列を設定します。値の設定のみ可能です。|
|showBubbleSize|bool|データ ラベルのバブルのサイズを表示または非表示にするかを表すブール型の値。値の設定のみ可能です。|
|showCategoryName|bool|データ ラベルのカテゴリ名を表示するか非表示にするかを表すブール型の値。値の設定のみ可能です。|
|showLegendKey|bool|データ ラベルの凡例マーカーを表示するか非表示にするかを表すブール型の値。値の設定のみ可能です。|
|showPercentage|bool|データ ラベルのパーセンテージを表示するか非表示にするかを表すブール型の値。値の設定のみ可能です。|
|showSeriesName|bool|データ ラベルの系列名を表示するか非表示にするかを表すブール型の値。値の設定のみ可能です。|
|showValue|bool|データ ラベルの値を表示するか非表示にするかを表すブール型の値。値の設定のみ可能です。|

_プロパティのアクセスの[例](#例)を参照してください。_

## リレーションシップ
| リレーションシップ | 型   |説明|
|:---------------|:--------|:----------|
|オプション パラメーターを適用する|[ChartDataLabelFormat](chartdatalabelformat.md)|グラフのデータ ラベルの書式 (塗りつぶしとフォントの書式設定を含む) を表します。値の取得のみ可能です。|

## メソッド

| メソッド           | 戻り値の型    |説明|
|:---------------|:--------|:----------|
|[load(param: object)](#loadparam-object)|void|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|

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
|param|object|省略可能。パラメーター名とリレーションシップ名を、区切られた文字列または 1 つの配列として受け入れます。あるいは、[loadOption](loadoption.md) オブジェクトを提供します。|

#### 戻り値
void
### プロパティのアクセスの例

データ ラベルに系列名を表示し、データ ラベルの `position` を "top" に設定します。

```js
Excel.run(function (ctx) { 
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1"); 
    chart.datalabels.visible = true;
    chart.datalabels.position = "top";
    chart.datalabels.ShowSeriesName = true;
    return ctx.sync().then(function() {
            console.log("Datalabels Shown");
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
