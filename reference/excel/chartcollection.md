# <a name="chartcollection-object-javascript-api-for-excel"></a>ChartCollection オブジェクト (JavaScript API for Excel)

ワークシート上のすべてのグラフ オブジェクトのコレクション。

## <a name="properties"></a>プロパティ

| プロパティ     | 型   |説明| 要件セット|
|:---------------|:--------|:----------|:----|
|count|int|ワークシート上のグラフの数を返します。値の取得のみ可能です。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|items|[Chart[]](chart.md)|グラフ オブジェクトのコレクション。値の取得のみ可能です。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_プロパティのアクセスの[例を参照してください。](#property-access-examples)_

## <a name="relationships"></a>関係
なし


## <a name="methods"></a>メソッド

| メソッド           | 戻り値の型    |説明| 要件セット|
|:---------------|:--------|:----------|:----|
|[add(type: string, sourceData:Range, seriesBy: string)](#addtype-string-sourcedata-range-seriesby-string)|[Chart](chart.md)|新しいグラフを作成します。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItem(name: string)](#getitemname-string)|[Chart](chart.md)|グラフ名を使用してグラフを取得します。同じ名前の複数のグラフがある場合は、最初の 1 つが返されます。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemAt(index: number)](#getitematindex-number)|[Chart](chart.md)|コレクション内の位置に基づいて、グラフを取得します。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemOrNull(name: string)](#getitemornullname-string)|[Chart](chart.md)|グラフ名を使用してグラフを取得します。同じ名前の複数のグラフがある場合は、最初の 1 つが返されます。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|[load(param: object)](#loadparam-object)|(非推奨)|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>メソッドの詳細


### <a name="addtype-string-sourcedata-range-seriesby-string"></a>add(type: 文字列、sourceData:範囲、seriesBy: 文字列)
新しいグラフを作成します。

#### <a name="syntax"></a>構文
```js
chartCollectionObject.add(type, sourceData, seriesBy);
```

#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|:---|
|type|string|グラフの種類を表します。可能なグラフの種類を以下に示します。|
|sourceData|Range|データ ソースに対応する Range オブジェクトです。|
|seriesBy|string|省略可能。列や行がグラフのデータ系列として使用される方法を指定します。使用可能な値は次のとおりです。自動、列、行|

**有効なグラフの種類を次に示します:**

`ColumnClustered`, `ColumnStacked`, `ColumnStacked100`, `_3DColumnClustered`, `_3DColumnStacked`, `_3DColumnStacked100`, `BarClustered`, `BarStacked`, `BarStacked100`, `_3DBarClustered`, `_3DBarStacked`, `_3DBarStacked100`, `LineStacked`, `LineStacked100`, `LineMarkers`, `LineMarkersStacked`, `LineMarkersStacked100`, `PieOfPie`, `PieExploded`, `_3DPieExploded`, `BarOfPie`, `XYScatterSmooth`, `XYScatterSmoothNoMarkers`, `XYScatterLines`, `XYScatterLinesNoMarkers`, `AreaStacked`, `AreaStacked100`, `_3DAreaStacked`, `_3DAreaStacked100`, `DoughnutExploded`, `RadarMarkers`, `RadarFilled`, `Surface`, `SurfaceWireframe`, `SurfaceTopView`, `SurfaceTopViewWireframe`, `Bubble`, `Bubble3DEffect`, `StockHLC`, `StockOHLC`, `StockVHLC`, `StockVOHLC`, `CylinderColClustered`, `CylinderColStacked`, `CylinderColStacked100`, `CylinderBarClustered`, `CylinderBarStacked`, `CylinderBarStacked100`, `CylinderCol`, `ConeColClustered`, `ConeColStacked`, `ConeColStacked100`, `ConeBarClustered`, `ConeBarStacked`, `ConeBarStacked100`, `ConeCol`, `PyramidColClustered`, `PyramidColStacked`, `PyramidColStacked100`, `PyramidBarClustered`, `PyramidBarStacked`, `PyramidBarStacked100`, `PyramidCol`, `_3DColumn`, `Line`, `_3DLine`, `_3DPie`, `Pie`, `XYScatter`, `_3DArea`, `Area`, `Doughnut`, `Radar`


#### <a name="returns"></a>戻り値
[Chart](chart.md)

#### <a name="examples"></a>例

`sourceData` が "A1:B4" の範囲で、`seriresBy` が "auto" に設定された、`chartType` が "ColumnClustered" であるグラフをワークシート "Charts" に追加します。

```js
Excel.run(function (ctx) { 
    var rangeSelection = "A1:B4";
    var range = ctx.workbook.worksheets.getItem(sheetName)
        .getRange(rangeSelection);
    var chart = ctx.workbook.worksheets.getItem(sheetName)
        .charts.add("ColumnClustered", range, "auto");  return ctx.sync().then(function() {
            console.log("New Chart Added");
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getitemname-string"></a>getItem(name: string)
グラフ名を使用してグラフを取得します。同じ名前の複数のグラフがある場合は、最初の 1 つが返されます。

#### <a name="syntax"></a>構文
```js
chartCollectionObject.getItem(name);
```

#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|:---|
|name|string|取得するグラフの名前。|

#### <a name="returns"></a>戻り値
[Chart](chart.md)

#### <a name="examples"></a>例

```js
Excel.run(function (ctx) { 
    var chartname = 'Chart1';
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem(chartname);
    return ctx.sync().then(function() {
            console.log(chart.height);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


#### <a name="examples"></a>例

```js
Excel.run(function (ctx) { 
    var chartId = 'SamplChartId';
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem(chartId);
    return ctx.sync().then(function() {
            console.log(chart.height);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```



#### <a name="examples"></a>例

```js
Excel.run(function (ctx) { 
    var lastPosition = ctx.workbook.worksheets.getItem("Sheet1").charts.count - 1;
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItemAt(lastPosition);
    return ctx.sync().then(function() {
            console.log(chart.name);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getitematindex-number"></a>getItemAt(index: number)
コレクション内での位置を基にグラフを取得します。

#### <a name="syntax"></a>構文
```js
chartCollectionObject.getItemAt(index);
```

#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|:---|
|index|number|取得するオブジェクトのインデックス値。0 を起点とする番号になります。|

#### <a name="returns"></a>戻り値
[Chart](chart.md)

#### <a name="examples"></a>例

```js
Excel.run(function (ctx) { 
    var lastPosition = ctx.workbook.worksheets.getItem("Sheet1").charts.count - 1;
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItemAt(lastPosition);
    return ctx.sync().then(function() {
            console.log(chart.name);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getitemornullname-string"></a>getItemOrNull(name: string)
グラフ名を使用してグラフを取得します。同じ名前の複数のグラフがある場合は、最初の 1 つが返されます。

#### <a name="syntax"></a>構文
```js
chartCollectionObject.getItemOrNull(name);
```

#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|:---|
|name|string|取得するグラフの名前。|

#### <a name="returns"></a>戻り値
[Chart](chart.md)

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

```js
Excel.run(function (ctx) { 
    var charts = ctx.workbook.worksheets.getItem("Sheet1").charts;
    charts.load('items');
    return ctx.sync().then(function() {
        for (var i = 0; i < charts.items.length; i++)
        {
            console.log(charts.items[i].name);
        }
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

グラフの数を取得します。

```js
Excel.run(function (ctx) { 
    var charts = ctx.workbook.worksheets.getItem("Sheet1").charts;
    charts.load('count');
    return ctx.sync().then(function() {
        console.log("charts: Count= " + charts.count);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

