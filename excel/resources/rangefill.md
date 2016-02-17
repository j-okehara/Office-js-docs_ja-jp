# RangeFill オブジェクト (JavaScript API for Excel)

_適用対象: Excel 2016、Excel Online、Office 2016_

Range オブジェクトの背景を表します。

## プロパティ

| プロパティ   | 型|説明
|:---------------|:--------|:----------|
|color|string|枠線の色を表す HTML カラー コード。形式は #RRGGBB (例:"FFA500")、または名前付きの HTML 色 (例: 「オレンジ」) です。|

_プロパティのアクセスの[例](#property-access-examples)を参照してください。_

## 関係
なし


## メソッド

| メソッド   | 戻り値の型|説明|
|:---------------|:--------|:----------|
|[clear()](#clear)|void|範囲の背景をリセットします。|
|[load(param: object)](#loadparam-object)|void|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|

## メソッドの詳細

### clear()
範囲の背景をリセットします。

#### 構文
```js
rangeFillObject.clear();
```

#### パラメーター
なし

#### 戻り値
(非推奨)

#### 例

この例では、範囲の背景をリセットします。

```js
Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "F:G";
	var worksheet = ctx.workbook.worksheets.getItem(sheetName);
	var range = worksheet.getRange(rangeAddress);
	var rangeFill = range.format.fill;
	rangeFill.clear();
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

#### パラメータ
| パラメーター   | 型|説明|
|:---------------|:--------|:----------|
|param|object|省略可能。パラメーター名とリレーションシップ名を、区切られた文字列または 1 つの配列として受け入れます。あるいは、[loadOption](loadoption.md) オブジェクトを提供します。|

#### 戻り値
void
### プロパティのアクセスの例
```js
Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "F:G";
	var worksheet = ctx.workbook.worksheets.getItem(sheetName);
	var range = worksheet.getRange(rangeAddress);
	var rangeFill = range.format.fill;
	rangeFill.load('color');
	return ctx.sync().then(function() {
		console.log(rangeFill.color);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
次の例では、塗りつぶしの色を設定します。 

```js
Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "F:G";
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
	range.format.fill.color = '0000FF';
	return ctx.sync(); 
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
