# <a name="chartfill-object-javascript-api-for-excel"></a>ChartFill オブジェクト (JavaScript API for Excel)

グラフ要素の塗りつぶしの書式設定を表します。

## <a name="properties"></a>プロパティ

なし

## <a name="relationships"></a>関係
なし


## <a name="methods"></a>メソッド

| メソッド           | 戻り値の型    |説明| 要件セット|
|:---------------|:--------|:----------|:----|
|[clear()](#clear)|void|グラフ要素の塗りつぶしの色をクリアします。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[setSolidColor(color: string)](#setsolidcolorcolor-string)|void|グラフ要素の塗りつぶしの書式設定を均一な色に設定します。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>メソッドの詳細


### <a name="clear"></a>clear()
グラフ要素の塗りつぶしの色をクリアします。

#### <a name="syntax"></a>構文
```js
chartFillObject.clear();
```

#### <a name="parameters"></a>パラメーター
なし

#### <a name="returns"></a>戻り値
void

#### <a name="examples"></a>例

"Chart1" という名前のグラフの数値軸の目盛線の線の書式をクリアします。

```js
Excel.run(function (ctx) { 
    var gridlines = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").axes.valueAxis.majorGridlines;   
    gridlines.format.line.clear();
    return ctx.sync().then(function() {
            console.log("Chart Major Gridlines Format Cleared");
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### <a name="setsolidcolorcolor-string"></a>setSolidColor(color: 文字列)
グラフ要素の塗りつぶしの書式設定を均一な色に設定します。

#### <a name="syntax"></a>構文
```js
chartFillObject.setSolidColor(color);
```

#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|:---|
|color|文字列|枠線の色を表す HTML カラー コード。形式は #RRGGBB (例: "FFA500")、または名前付きの HTML 色 (例: "オレンジ") です。|

#### <a name="returns"></a>戻り値
void

#### <a name="examples"></a>例

Chart1 の背景色を赤に設定します。

```js
Excel.run(function (ctx) { 
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1"); 

    chart.format.fill.setSolidColor("#FF0000");

    return ctx.sync().then(function() {
            console.log("Chart1 Background Color Changed.");
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
