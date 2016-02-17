# ChartGridlines オブジェクト (JavaScript API for Excel)

_適用対象: Excel 2016、Excel Online、Office 2016_

グラフの軸の目盛線または補助目盛線を表します。

## プロパティ

| プロパティ   | 型|説明
|:---------------|:--------|:----------|
|visible|bool|軸の目盛線を表示するか非表示にするかを表すブール型の値。|

_プロパティのアクセスの[例](#property-access-examples)を参照してください。_

## 関係
| リレーションシップ | 型|説明|
|:---------------|:--------|:----------|
|オプション パラメーターを適用する|[ChartGridlinesFormat](chartgridlinesformat.md)|グラフの目盛線の書式設定を表します。値の取得のみ可能です。|

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

Chart1 の数値軸の大きい目盛線の `visible` プロパティを取得します。

```js
Excel.run(function (ctx) { 
	var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	
	var majGridlines = chart.axes.valueaxis.majorGridlines;
	majGridlines.load('visible');
	return ctx.sync().then(function() {
			console.log(majGridlines.visible);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

Chart1 の数値軸の大きい目盛線を表示するように設定します。

```js
Excel.run(function (ctx) { 
	var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	
	chart.axes.valueaxis.majorgridlines.visible = true;
	return ctx.sync().then(function() {
			console.log("Axis Gridlines Added ");
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

