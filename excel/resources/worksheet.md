# Worksheet オブジェクト (JavaScript API for Excel)

_適用対象: Excel 2016、Excel Online、Office 2016_

Excel のワークシートは、セルのグリッドになっています。そこに、データ、表、グラフなどを含めることができます。

## プロパティ

| プロパティ   | 型|説明
|:---------------|:--------|:----------|
|id|string|指定されたブックのワークシートを一意に識別する値を返します。この識別子の値は、ワークシートの名前を変更したり移動したりしても同じままです。値の取得のみ可能です。|
|name|string|ワークシートの表示名。|
|position|int|0 を起点とした、ブック内のワークシートの位置。|
|visibility|string|ワークシートの可視性。使用可能な値は次のとおりです: Visible、Hidden、VeryHidden。読み取り専用です。|

_プロパティのアクセスの[例](#property-access-examples)をご覧ください。_

## 関係
| リレーションシップ | 型|説明|
|:---------------|:--------|:----------|
|グラフ|[ChartCollection](chartcollection.md)|ワークシートの一部になっているグラフのコレクションを返します。値の取得のみ可能です。|
|テーブル|[TableCollection](tablecollection.md)|ワークシートの一部になっているグラフのコレクション。値の取得のみ可能です。|

## メソッド

| メソッド   | 戻り値の型|説明|
|:---------------|:--------|:----------|
|[activate()](#activate)|void|Excel UI でワークシートをアクティブにします。|
|[delete()](#delete)|void|ブックからワークシートを削除します。|
|[getCell(row: number, column: number)](#getcellrow-number-column-number)|[Range](range.md)|行と列の番号に基づいて、1 つのセルを含んだ範囲オブジェクトを取得します。このセルは、このワークシートのグリッド内であれば、親の範囲の境界の外のセルであってもかまいません。|
|[getRange(address: string)](#getrangeaddress-string)|[Range](range.md)|アドレスまたは名前で指定された範囲 オブジェクトを取得します。|
|[getUsedRange()](#getusedrange)|[Range](range.md)|使用範囲とは、値または書式設定が割り当たっているすべてのセルを包含する最小の範囲です。ワークシートが空白の場合、この関数は左上のセルを返します。|
|[load(param: object)](#loadparam-object)|void|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|

## メソッドの詳細

### activate()
Excel UI でワークシートをアクティブにします。

#### 構文
```js
worksheetObject.activate();
```

#### パラメーター
なし

#### 戻り値
(非推奨)

#### 例

```js
Excel.run(function (ctx) { 
	var wSheetName = 'Sheet1';
	var worksheet = ctx.workbook.worksheets.getItem(wSheetName);
	worksheet.activate();
	return ctx.sync(); 
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### delete()
ブックからワークシートを削除します。

#### 構文
```js
worksheetObject.delete();
```

#### パラメーター
なし

#### 戻り値
(非推奨)

#### 例

```js
Excel.run(function (ctx) { 
	var wSheetName = 'Sheet1';
	var worksheet = ctx.workbook.worksheets.getItem(wSheetName);
	worksheet.delete();
	return ctx.sync(); 
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### getCell(row: number, column: number)
行と列の番号に基づいて、1 つのセルを含んだ範囲オブジェクトを取得します。このセルは、このワークシートのグリッド内であれば、親の範囲の境界の外のセルであってもかまいません。

#### 構文
```js
worksheetObject.getCell(row, column);
```

#### パラメータ
| パラメーター   | 型|説明|
|:---------------|:--------|:----------|
|row|number|取得するセルの行番号。0 を起点とする番号になります。|
|column|number|取得するセルの列番号。0 を起点とする番号になります。|

#### 戻り値
[Range](range.md)

#### 例

```js
Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "A1:F8";
	var worksheet = ctx.workbook.worksheets.getItem(sheetName);
	var cell = worksheet.getCell(0,0);
	cell.load('address');
	return ctx.sync().then(function() {
		console.log(cell.address);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### getRange(address: string)
アドレスまたは名前で指定された範囲 オブジェクトを取得します。

#### 構文
```js
worksheetObject.getRange(address);
```

#### パラメータ
| パラメーター   | 型|説明|
|:---------------|:--------|:----------|
|address|string|省略可能。範囲のアドレスまたは名前。指定されていない場合は、ワークシート全体の範囲が返されます。|

#### 戻り値
[Range](range.md)

#### 例
次の例では、範囲アドレスを使用して、範囲オブジェクトを取得しています。

```js
Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "A1:F8";
	var worksheet = ctx.workbook.worksheets.getItem(sheetName);
	var range = worksheet.getRange(rangeAddress);
	range.load('cellCount');
	return ctx.sync().then(function() {
		console.log(range.cellCount);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

次の例では、名前付き範囲を使用して、範囲オブジェクトを取得しています。

```js

Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeName = 'MyRange';
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeName);
	range.load('address');
	return ctx.sync().then(function() {
		console.log(range.address);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
### getUsedRange()
使用範囲とは、値または書式設定が割り当たっているすべてのセルを包含する最小の範囲です。ワークシートが空白の場合、この関数は左上のセルを返します。

#### 構文
```js
worksheetObject.getUsedRange();
```

#### パラメーター
なし

#### 戻り値
[Range](range.md)

#### 例

```js
Excel.run(function (ctx) { 
	var wSheetName = 'Sheet1';
	var worksheet = ctx.workbook.worksheets.getItem(wSheetName);
	var usedRange = worksheet.getUsedRange();
	usedRange.load('address');
	return ctx.sync().then(function() {
			console.log(usedRange.address);
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

シート名に基づいて、ワークシートのプロパティを取得します。

```js
Excel.run(function (ctx) { 
	var wSheetName = 'Sheet1';
	var worksheet = ctx.workbook.worksheets.getItem(wSheetName);
	worksheet.load('position')
	return ctx.sync().then(function() {
			console.log(worksheet.position);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

ワークシートの位置を設定します。 

```js
Excel.run(function (ctx) { 
	var wSheetName = 'Sheet1';
	var worksheet = ctx.workbook.worksheets.getItem(wSheetName);
	worksheet.position = 2;
	return ctx.sync(); 
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```


