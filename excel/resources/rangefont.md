# RangeFont オブジェクト (JavaScript API for Excel)

_適用対象: Excel 2016、Excel Online、Office 2016_

このオブジェクトは、オブジェクトのフォントの属性 (フォント名、フォント サイズ、色など) を表します。

## プロパティ

| プロパティ   | 型|説明
|:---------------|:--------|:----------|
|bold|bool|フォントの太字設定を表します。|
|color|string|テキストの色の HTML カラー コード表記。たとえば、#FF0000 は赤を表します。|
|italic|bool|フォントの斜体設定を表します。|
|name|string|フォント名 (例:"Calibri")|
|size|double|フォント サイズ|
|underline|string|フォントに適用する下線の種類。使用可能な値は次のとおりです。None、Single、Double、SingleAccountant、DoubleAccountant。|

_プロパティのアクセスの[例](#property-access-examples)を参照してください。_

## 関係
なし


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

```js
Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "F:G";
	var worksheet = ctx.workbook.worksheets.getItem(sheetName);
	var range = worksheet.getRange(rangeAddress);
	var rangeFont = range.format.font;
	rangeFont.load('name');
	return ctx.sync().then(function() {
		console.log(rangeFont.name);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
次の例では、フォント名を設定します。 

```js
Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "F:G";
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
	range.format.font.name = 'Times New Roman';
	return ctx.sync(); 
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
