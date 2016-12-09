# <a name="chartlegend-object-javascript-api-for-excel"></a>ChartLegend オブジェクト (JavaScript API for Excel)

グラフに凡例を表します。

## <a name="properties"></a>プロパティ

| プロパティ     | 型   |説明| 要件セット|
|:---------------|:--------|:----------|:----|
|overlay|bool|グラフの凡例をグラフの本体に重ねるかどうかを指定するブール型の値です。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|position|string|グラフの凡例の位置を表します。使用可能な値は次のとおりです。Top、Bottom、Left、Right、Corner、Custom。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|visible|bool|ChartLegend オブジェクトを表示または非表示にするかを表すブール型の値。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_プロパティのアクセスの[例を参照してください。](#property-access-examples)_

## <a name="relationships"></a>関係
| リレーションシップ | 型   |説明| 要件セット|
|:---------------|:--------|:----------|:----|
|format|[ChartLegendFormat](chartlegendformat.md)|グラフの凡例の書式設定を表します。これには塗りつぶしとフォントの書式設定などがあります。読み取り専用です。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

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

Chart1 からグラフの凡例の `position` を取得します。

```js
Excel.run(function (ctx) { 
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1"); 
    var legend = chart.legend;
    legend.load('position');
    return ctx.sync().then(function() {
            console.log(legend.position);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

Chart1 の凡例を表示させ、グラフの一番上になるように設定します。

```js
Excel.run(function (ctx) { 
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1"); 
    chart.legend.visible = true;
    chart.legend.position = "top"; 
    chart.legend.overlay = false; 
    return ctx.sync().then(function() {
            console.log("Legend Shown ");
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
``` 
