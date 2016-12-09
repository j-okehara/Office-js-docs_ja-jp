# <a name="chart-object-javascript-api-for-excel"></a>Chart オブジェクト (JavaScript API for Excel)

ブック内のグラフ オブジェクトを表します。

## <a name="properties"></a>プロパティ

| プロパティ     | 型   |説明| 要件セット|
|:---------------|:--------|:----------|:----|
|height|double|グラフ オブジェクトの高さをポイント単位で表します。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|id|文字列|コレクション内の位置に基づいて、グラフを取得します。読み取り専用です。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|left|double|グラフの左側からワークシートの原点までの距離 (ポイント単位)。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|name|string|グラフ オブジェクトの名前を表します。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|top|double|オブジェクトの上端から (ワークシートの) 1 行目の上部または (グラフの) グラフ領域の上部までの距離をポイント単位で表します。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|width|double|グラフ オブジェクトの幅をポイント単位で表します。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_プロパティのアクセスの[例を参照してください。](#property-access-examples)_

## <a name="relationships"></a>関係
| リレーションシップ | 型   |説明| 要件セット|
|:---------------|:--------|:----------|:----|
|axes|[ChartAxes](chartaxes.md)|グラフの軸を表します。読み取り専用です。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|dataLabels|[ChartDataLabels](chartdatalabels.md)|グラフのデータラベルを表します。読み取り専用です。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|format|[ChartAreaFormat](chartareaformat.md)|グラフ領域の書式設定プロパティをカプセル化します。読み取り専用です。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|legend|[ChartLegend](chartlegend.md)|グラフの凡例を表します。読み取り専用です。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|series|[ChartSeriesCollection](chartseriescollection.md)|グラフの 1 つのデータ系列またはデータ系列のコレクションを表します。読み取り専用です。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|title|[ChartTitle](charttitle.md)|指定したグラフのタイトル (タイトルのテキスト、表示/非表示、位置、書式設定など) を表します。読み取り専用です。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|worksheet|[Worksheet](worksheet.md)|現在のグラフを含んでいるワークシート。読み取り専用です。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>メソッド

| メソッド           | 戻り値の型    |説明| 要件セット|
|:---------------|:--------|:----------|:----|
|[delete()](#delete)|void|グラフ オブジェクトを削除します。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getImage(height: number, width: number, fittingMode: string)](#getimageheight-number-width-number-fittingmode-string)|[System.IO.Stream](system.io.stream.md)|指定したサイズに合わせてグラフを拡大、縮小することで、グラフを Base64 でエンコードされた画像としてレンダリングします。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[load(param: object)](#loadparam-object)|(非推奨)|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[setData(sourceData:Range, seriesBy: string)](#setdatasourcedata-range-seriesby-string)|void|グラフの元データをリセットします。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[setPosition(startCell:Range or string, endCell:Range or string)](#setpositionstartcell-range-or-string-endcell-range-or-string)|void|ワークシート上のセルを基準にしてグラフを配置します。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>メソッドの詳細


### <a name="delete"></a>delete()
グラフ オブジェクトを削除します。

#### <a name="syntax"></a>構文
```js
chartObject.delete();
```

#### <a name="parameters"></a>パラメーター
なし

#### <a name="returns"></a>戻り値
void

#### <a name="examples"></a>例
```js
Excel.run(function (ctx) { 
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1"); 
    chart.delete();
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### <a name="getimageheight-number-width-number-fittingmode-string"></a>getImage(height: number, width: number, fittingMode: string)
指定したサイズに合わせてグラフを拡大・縮小することで、グラフを Base64 でエンコードされた画像としてレンダリングします。

#### <a name="syntax"></a>構文
```js
chartObject.getImage(height, width, fittingMode);
```

#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|:---|
|height|number|省略可能。(省略可能) 結果の画像の希望する高さ。|
|width|number|省略可能。(省略可能) 結果の画像の希望する幅。|
|fittingMode|string|省略可能。(省略可能) 指定したディメンションに合わせてグラフを拡大または縮小するために使用するメソッド (高さと幅の両方が設定されている場合)。使用可能な値は次のとおりです。Fit、FitAndCenter、Fill|

#### <a name="returns"></a>戻り値
[System.IO.Stream](system.io.stream.md)

#### <a name="examples"></a>例
```js
Excel.run(function (ctx) { 
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1"); 
    var image = chart.getImage();
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```





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

### <a name="setdatasourcedata-range-seriesby-string"></a>setData(sourceData: Range, seriesBy: string)
グラフの元データをリセットします。

#### <a name="syntax"></a>構文
```js
chartObject.setData(sourceData, seriesBy);
```

#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|:---|
|sourceData|Range|データ ソースに対応する Range オブジェクトです。|
|seriesBy|文字列|省略可能。列や行がグラフのデータ系列として使用される方法を指定します。次のいずれかを指定できます。自動 (既定)、行、列。使用可能な値は次のとおりです。自動、列、行|

#### <a name="returns"></a>戻り値
void

#### <a name="examples"></a>例

`sourceData` を "A1:B4"、`seriesBy` を "列" と設定します

```js
Excel.run(function (ctx) { 
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1"); 
    var sourceData = "A1:B4";
    chart.setData(sourceData, "Columns");
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="setpositionstartcell-range-or-string-endcell-range-or-string"></a>setPosition(startCell:範囲または文字列、endCell:Range or string)
ワークシート上のセルを基準にしてグラフを配置します。

#### <a name="syntax"></a>構文
```js
chartObject.setPosition(startCell, endCell);
```

#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|:---|
|startCell|Range または string|開始セル。これは、グラフの移動先です。開始セルは、ユーザーの右から左への表示の設定に応じて、左上のセルか、右上のセルとなります。|
|endCell|Range または string|省略可能。(省略可能) 最後のセル。指定されている場合、グラフの幅と高さは、このセルまたは範囲を完全にカバーするように設定されます。|

#### <a name="returns"></a>戻り値
void

#### <a name="examples"></a>例


```js
Excel.run(function (ctx) { 
    var sheetName = "Charts";
    var rangeSelection = "A1:B4";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeSelection);
    var sourceData = sheetName + "!" + "A1:B4";
    var chart = ctx.workbook.worksheets.getItem(sheetName).charts.add("pie", range, "auto");
    chart.width = 500;
    chart.height = 300;
    chart.setPosition("C2", null);
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### <a name="property-access-examples"></a>プロパティのアクセスの例

"Chart1" という名前のグラフを取得します。

```js
Excel.run(function (ctx) { 
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1"); 
    chart.load('name');
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

名前、配置、サイズなどを変更してグラフを更新します。

```js
Excel.run(function (ctx) { 
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1"); 
    chart.name="New Name";
    chart.top = 100;
    chart.left = 100;
    chart.height = 200;
    chart.width = 200;
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

グラフの名前を新しい名前に変更し、グラフの高さと幅の両方を 200 ポイントにサイズ変更します。Chart1 を上と左に 100 ポイントずつ移動します。 

```js
Excel.run(function (ctx) { 
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");
    chart.name="New Name";  
    chart.top = 100;
    chart.left = 100;
    chart.height =200;
    chart.width =200;
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

