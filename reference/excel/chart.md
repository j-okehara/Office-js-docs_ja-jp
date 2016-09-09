# Chart オブジェクト (JavaScript API for Excel)

ブック内のグラフ オブジェクトを表します。

## プロパティ

| プロパティ     | 型   |説明
|:---------------|:--------|:----------|
|height|double|グラフ オブジェクトの高さをポイント単位で表します。|
|id|string|コレクション内での位置を基にグラフを取得します。読み取り専用です。|
|left|double|グラフの左側からワークシートの原点までの距離 (ポイント単位)。|
|name|string| グラフ オブジェクトの名前を表します。|
|top|double|オブジェクトの上端から (ワークシートの) 1 行目の上部または (グラフの) グラフ領域の上部までの距離をポイント単位で表します。|
|width|double|グラフ オブジェクトの幅をポイント単位で表します。|

_プロパティのアクセスの[例](#例)を参照してください。_

## リレーションシップ
| リレーションシップ | 型   |説明|
|:---------------|:--------|:----------|
|axes|[ChartAxes](chartaxes.md)|グラフの軸を表します。値の取得のみ可能です。|
|dataLabels|[ChartDataLabels](chartdatalabels.md)|グラフのデータ ラベルを表します。値の取得のみ可能です。|
|オプション パラメーターを適用する|[ChartAreaFormat](chartareaformat.md)|グラフ領域の書式設定プロパティをカプセル化します。値の取得のみ可能です。|
|legend|[ChartLegend](chartlegend.md)|グラフの凡例を表します。値の取得のみ可能です。|
|データ系列|[ChartSeriesCollection](chartseriescollection.md)|グラフの 1 つの系列または系列のコレクションを表します。値の取得のみ可能です。|
|title|[ChartTitle](charttitle.md)|指定したグラフのタイトル (タイトルのテキスト、表示/非表示、位置、書式設定など) を表します。値の取得のみ可能です。|

## メソッド

| メソッド           | 戻り値の型    |説明|
|:---------------|:--------|:----------|
|[delete()](#delete)|void|グラフ オブジェクトを削除します。|
|[getImage(height: number, width: number, fittingMode: string)](#getimageheight-number-width-number-fittingmode-string)|System.IO.Stream|指定したサイズに合わせてグラフを拡大・縮小することで、グラフを Base64 でエンコードされた画像としてレンダリングします。|
|[load(param: object)](#loadparam-object)|void|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|
|[setData(sourceData: Range, seriesBy: string)](#setdatasourcedata-range-seriesby-string)|void|グラフの元データをリセットします。|
|[setPosition(startCell:Range or string, endCell:Range or string)](#setpositionstartcell-range-or-string-endcell-range-or-string)|void|ワークシート上のセルを基準にしてグラフを配置します。|

## メソッドの詳細


### delete()
グラフ オブジェクトを削除します。

#### 構文
```js
chartObject.delete();
```

#### パラメーター
なし

#### 戻り値
void

#### 例
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

### getImage(height: number, width: number, fittingMode: string)
指定したサイズに合わせてグラフを拡大・縮小することで、グラフを Base64 でエンコードされた画像としてレンダリングします。

#### 構文
```js
chartObject.getImage(height, width, fittingMode);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|height|number|省略可能。(省略可能) 結果の画像の希望する高さ。|
|width|number|省略可能。(省略可能) 結果の画像の希望する幅。|
|fittingMode|string|省略可能。(省略可能) 指定したディメンションに合わせてグラフを拡大または縮小するために使用するメソッド (高さと幅の両方が設定されている場合)。使用可能な値は次のとおりです。Fit、FitAndCenter、Fill|

#### 戻り値
System.IO.Stream

#### 例
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

### setData(sourceData: Range, seriesBy: string)
グラフの元データをリセットします。

#### 構文
```js
chartObject.setData(sourceData, seriesBy);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|sourceData|Range|データ ソースに対応する Range オブジェクトです。|
|seriesBy|string|省略可能。列や行がグラフのデータ系列として使用される方法を指定します。使用可能な値は次のとおりです。Auto、Columns、Rows。Desktop では、"auto" オプションを使用するとソース データの図形が検査され、そのデータの行または列のどちらを使用するかが自動的に推測されます。Excel Online では、"auto" は単純に既定の "columns" が使用されます。|

#### 戻り値
void

#### 例

`sourceData` を "A1:B4"、`seriesBy` を "列" と設定します。

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


### setPosition(startCell:Range or string, endCell:Range or string)
ワークシート上のセルを基準にしてグラフを配置します。

#### 構文
```js
chartObject.setPosition(startCell, endCell);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|startCell|Range または string|開始セル。これは、グラフの移動先です。開始セルは、ユーザーの左から右への表示の設定に応じて、左上のセルか、右上のセルとなります。|
|endCell|Range または string|省略可能。最終セル。指定されている場合、グラフの幅と高さは、このセルまたは範囲までを完全にカバーするように設定されます。|

#### 戻り値
void

#### 例


```js
Excel.run(function (ctx) { 
    var sheetName = "Charts";
    var sourceData = sheetName + "!" + "A1:B4";
    var chart = ctx.workbook.worksheets.getItem(sheetName).charts.add("pie", sourceData, "auto");
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

### プロパティのアクセスの例

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
    chart.weight = 200;
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

グラフに新しい名前を付け、グラフの高さと幅を両方とも 200 ポイントにサイズ変更します。Chart1 を上にそして左に 100 ポイントずつ移動します。 

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

