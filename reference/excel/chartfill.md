# ChartFill オブジェクト (JavaScript API for Excel)

グラフ要素の塗りつぶしの書式設定を表します。

## プロパティ

なし

## 関係
なし


## メソッド

| メソッド           | 戻り値の型    |説明|
|:---------------|:--------|:----------|
|[clear()](#clear)|void|グラフ要素の塗りつぶしの色をクリアします。|
|[setSolidColor(color: string)](#setsolidcolorcolor-string)|void|グラフ要素の塗りつぶしの書式設定を均一な色に設定します。|

## メソッドの詳細


### clear()
グラフ要素の塗りつぶしの色をクリアします。

#### 構文
```js
chartFillObject.clear();
```

#### パラメーター
なし

#### 戻り値
void

#### 例

"Chart1" という名前のグラフの数値軸の目盛線の線の書式をクリアします。

```js
Excel.run(function (ctx) { 
    var gridlines = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").axes.valueaxis.majorGridlines;   
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

### setSolidColor(color: string)
グラフ要素の塗りつぶしの書式設定を均一な色に設定します。

#### 構文
```js
chartFillObject.setSolidColor(color);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|color|string|枠線の色を表す HTML カラー コード。形式は #RRGGBB (例:"FFA500")、または名前付きの HTML 色 (例: "orange") です。|

#### 戻り値
void

#### 例

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
