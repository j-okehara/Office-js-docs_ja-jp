# ChartAxis オブジェクト (JavaScript API for Excel)

_適用対象: Excel 2016、Excel Online、Office 2016_

グラフの 1 つの軸を表します。

## プロパティ

| プロパティ   | 型|説明
|:---------------|:--------|:----------|
|majorUnit|object|2 つの大きい目盛の間隔を表します。数値の値または空の文字列を設定できます。戻り値は常に数値です。|
|maximum|object|数値軸の最大値を表します。数値の値または空の文字列を設定できます (軸の値が自動の場合)。戻り値は常に数値です。|
|minimum|object|数値軸の最大値を表します。数値の値または空の文字列を設定できます (軸の値が自動の場合)。戻り値は常に数値です。|
|minorUnit|object|2 つの小さい目盛の間隔を表します。"数値の値または空の文字列を設定できます (軸の値が自動の場合)。戻り値は常に数値です。|

_プロパティのアクセスの[例](#property-access-examples)をご覧ください。_

## 関係
| リレーションシップ | 型|説明|
|:---------------|:--------|:----------|
|オプション パラメーターを適用する|[ChartAxisFormat](chartaxisformat.md)|グラフ オブジェクトの書式設定を表します。これには線とフォントの書式設定などがあります。値の取得のみ可能です。|
|majorGridlines|[ChartGridlines](chartgridlines.md)|指定された軸の大きい目盛線を表す Gridlines オブジェクトを返します。値の取得のみ可能です。|
|minorGridlines|[ChartGridlines](chartgridlines.md)|指定された軸の小さい目盛線を表す gridlines オブジェクトを返します。値の取得のみ可能です。|
|title|[ChartAxisTitle](chartaxistitle.md)|軸タイトルを表します。値の取得のみ可能です。|

## メソッド

| メソッド   | 戻り値の型|説明|
|:---------------|:--------|:----------|
|[load(param: object)](#loadparam-object)|void|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|

## メソッドの詳細

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
Chart1 のグラフ軸の `maximum` を取得します。

```js
Excel.run(function (ctx) { 
	var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	
	var axis = chart.axes.valueaxis;
	axis.load('maximum');
	return ctx.sync().then(function() {
			console.log(axis.maximum);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

数値軸の `maximum`、`minimum`、`majorunit`、または `minorunit` を設定します。 

```js
Excel.run(function (ctx) { 
	var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	
	chart.axes.valueaxis.maximum = 5;
	chart.axes.valueaxis.minimum = 0;
	chart.axes.valueaxis.majorunit = 1;
	chart.axes.valueaxis.minorunit = 0.2;
	return ctx.sync().then(function() {
			console.log("Axis Settings Changed");
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

