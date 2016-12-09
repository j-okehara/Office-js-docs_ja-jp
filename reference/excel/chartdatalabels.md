# <a name="chartdatalabels-object-javascript-api-for-excel"></a>ChartDataLabels オブジェクト (JavaScript API for Excel)

グラフのポイントにあるすべてのデータ ラベルのコレクションを表します。

## <a name="properties"></a>プロパティ

| プロパティ     | 型   |説明| 要件セット|
|:---------------|:--------|:----------|:----|
|position|string|データ ラベルの位置を表すDataLabelPosition 値。使用可能な値は次のとおりです。None、Center、InsideEnd、InsideBase、OutsideEnd、Left、Right、Top、Bottom、BestFit、Callout。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|separator|string|グラフのデータ ラベルに使用される区切り文字を表す文字列を設定します。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|showBubbleSize|bool|データ ラベルのバブルのサイズを表示または非表示にするかを表すブール型の値。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|showCategoryName|bool|データ ラベルのカテゴリ名を表示するか非表示にするかを表すブール型の値。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|showLegendKey|bool|データ ラベルの凡例マーカーを表示するか非表示にするかを表すブール型の値。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|showPercentage|bool|データ ラベルのパーセンテージを表示するか非表示にするかを表すブール型の値。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|showSeriesName|bool|データ ラベルの系列名を表示するか非表示にするかを表すブール型の値。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|showValue|bool|データ ラベルの値を表示するか非表示にするかを表すブール型の値。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_プロパティのアクセスの[例を参照してください。](#property-access-examples)_

## <a name="relationships"></a>関係
| リレーションシップ | 型   |説明| 要件セット|
|:---------------|:--------|:----------|:----|
|format|[ChartDataLabelFormat](chartdatalabelformat.md)|グラフのデータ ラベルの書式 (塗りつぶしとフォントの書式設定を含む) を表します。読み取り専用です。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>メソッド

| メソッド           | 戻り値の型    |説明| 要件セット|
|:---------------|:--------|:----------|:----|
|[load(param: object)](#loadparam-object)|(非推奨)|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>メソッドの詳細


### <a name="loadparam-object"></a>load(param: object)
JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。

#### <a name="syntax"></a>構文
```js
object.load(param);
```

#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|:---|
|param|object|省略可能。パラメーターとリレーションシップ名を、区切られた文字列または 1 つの配列として受け入れます。あるいは、[loadOption](loadoption.md) オブジェクトを提供します。|

#### <a name="returns"></a>戻り値
void
### <a name="property-access-examples"></a>プロパティのアクセスの例

データラベルに系列名を表示し、データラベルの `position` を "top" に設定します。

```js
Excel.run(function (ctx) { 
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1"); 
    chart.datalabels.showValue = true;
    chart.datalabels.position = "top";
    chart.datalabels.showSeriesName = true;
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
