# ChartCollection オブジェクト (JavaScript API for Excel)

_適用対象: Excel 2016、Excel Online、Office 2016_

ワークシート上のすべてのグラフ オブジェクトのコレクション。

## プロパティ

| プロパティ   | 型|説明
|:---------------|:--------|:----------|
|count|int|ワークシート上のグラフの数を返します。値の取得のみ可能です。|
|Items|[Chart[]](chart.md)|グラフ オブジェクトのコレクション。値の取得のみ可能です。|

_プロパティのアクセスの[例](#property-access-examples)を参照してください。_

## 関係
なし


## メソッド

| メソッド   | 戻り値の型|説明|
|:---------------|:--------|:----------|
|[add(type: string, sourceData:Range, seriesBy: string)](#addtype-string-sourcedata-range-seriesby-string)|[Chart](chart.md)|新しいグラフを作成します。|
|[getItem(name: string)](#getitemname-string)|[Chart](chart.md)|グラフ名を使用してグラフを取得します。同じ名前の複数のグラフがある場合は、最初の 1 つが返されます。|
|[getItemAt(index: number)](#getitematindex-number)|[Chart](chart.md)|コレクション内での位置を基にグラフを取得します。|
|[load(param: object)](#loadparam-object)|void|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|

## メソッドの詳細

### add(type: 文字列、sourceData:範囲、seriesBy: 文字列)
新しいグラフを作成します。

#### 構文
```js
chartCollectionObject.add(type, sourceData, seriesBy);
```

#### パラメータ
| パラメーター   | 型|説明|
|:---------------|:--------|:----------|
|type|string|グラフの種類を表します。使用可能な値は次のとおりです。ColumnClustered、ColumnStacked、ColumnStacked100、BarClustered、BarStacked、BarStacked100、LineStacked、LineStacked100、LineMarkers、LineMarkersStacked、LineMarkersStacked100、PieOfPie など。|
|sourceData|Range|元データを含む range オブジェクト。|
|seriesBy|string|省略可能。列や行がグラフのデータ系列として使用される方法を指定します。使用可能な値は次のとおりです。自動、列、行|

#### 戻り値
[Chart](chart.md)

#### 例

`sourceData` が "A1:B4" の範囲で、`seriesBy` が "auto" に設定されたワークシート "Charts" で、`chartType` "ColumnClustered" のグラフを追加します。

```js
Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var sourceData = sheetName + "!" + "A1:B4";
	var chart = ctx.workbook.worksheets.getItem(sheetName).charts.add("ColumnClustered", sourceData, "auto");
	return ctx.sync().then(function() {
			console.log("New Chart Added");
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### getItem(name: string)
グラフ名を使用してグラフを取得します。同じ名前の複数のグラフがある場合は、最初の 1 つが返されます。

#### 構文
```js
chartCollectionObject.getItem(name);
```

#### パラメータ
| パラメーター   | 型|説明|
|:---------------|:--------|:----------|
|name|string|取得するグラフの名前。|

#### 戻り値
[Chart](chart.md)

#### 例

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


#### 例

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



#### 例

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

### getItemAt(index: number)
コレクション内での位置を基にグラフを取得します。

#### 構文
```js
chartCollectionObject.getItemAt(index);
```

#### パラメータ
| パラメーター   | 型|説明|
|:---------------|:--------|:----------|
|index|number|取得するオブジェクトのインデックス値。0 を起点とする番号になります。|

#### 戻り値
[Chart](chart.md)

#### 例

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

### load(param: object)
JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。

#### 構文
```js
object.load(param);
```

#### パラメータ
| パラメーター   | 型|説明|
|:---------------|:--------|:----------|
|param|object|省略可能。パラメーター名とリレーションシップ名を、区切られた文字列または 1 つの配列として受け入れます。あるいは、[loadOption](loadoption.md) オブジェクトを提供します。|

#### 戻り値
void
### プロパティのアクセスの例

```js
Excel.run(function (ctx) { 
	var charts = ctx.workbook.worksheets.getItem("Sheet1").charts;
	charts.load('items');
	return ctx.sync().then(function() {
		for (var i = 0; i < charts.items.length; i++)
		{
			console.log(charts.items[i].name);
			console.log(charts.items[i].index);
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


