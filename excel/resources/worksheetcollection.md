# WorksheetCollection オブジェクト (JavaScript API for Excel)

_適用対象: Excel 2016、Excel Online、Office 2016_

ブックの一部であるワークシート オブジェクトのコレクションを表します。

## プロパティ

| プロパティ   | 型|説明
|:---------------|:--------|:----------|
|Items|[Worksheet[]](worksheet.md)|ワークシート オブジェクトのコレクション。値の取得のみ可能です。|

_プロパティのアクセスの[例](#property-access-examples)をご覧ください。_

## 関係
なし


## メソッド

| メソッド   | 戻り値の型|説明|
|:---------------|:--------|:----------|
|[add(name: string)](#addname-string)|[Worksheet](worksheet.md)|新しいワークシートをブックに追加します。ワークシートは、既存のワークシートの末尾に追加されます。新しく追加したワークシートをアクティブにする場合は、そのワークシートに対して ".activate() を呼び出します。|
|[getActiveWorksheet()](#getactiveworksheet)|[Worksheet](worksheet.md)|ブックの、現在作業中のワークシートを取得します。|
|[getItem(key: string)](#getitemkey-string)|[Worksheet](worksheet.md)|名前または ID を使用して、ワークシート オブジェクトを取得します。|
|[load(param: object)](#loadparam-object)|void|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|

## メソッドの詳細

### add(name: string)
新しいワークシートをブックに追加します。ワークシートは、既存のワークシートの末尾に追加されます。新しく追加したワークシートをアクティブにする場合は、そのワークシートに対して ".activate() を呼び出します。

#### 構文
```js
worksheetCollectionObject.add(name);
```

#### パラメータ
| パラメーター   | 型|説明|
|:---------------|:--------|:----------|
|name|string|省略可能。追加するワークシートの名前。指定する場合、名前は一意である必要があります。指定されていない場合は、Excel が新しいワークシートの名前を決定します。|

#### 戻り値
[Worksheet](worksheet.md)

#### 例

```js
Excel.run(function (ctx) { 
	var wSheetName = 'Sample Name';
	var worksheet = ctx.workbook.worksheets.add(wSheetName);
	worksheet.load('name');
	return ctx.sync().then(function() {
		console.log(worksheet.name);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### getActiveWorksheet()
ブックの、現在作業中のワークシートを取得します。

#### 構文
```js
worksheetCollectionObject.getActiveWorksheet();
```

#### パラメーター
なし

#### 戻り値
[Worksheet](worksheet.md)

#### 例

```js
Excel.run(function (ctx) {  
	var activeWorksheet = ctx.workbook.worksheets.getActiveWorksheet();
	activeWorksheet.load('name');
	return ctx.sync().then(function() {
			console.log(activeWorksheet.name);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### getItem(key: string)
名前または ID を使用して、ワークシート オブジェクトを取得します。

#### 構文
```js
worksheetCollectionObject.getItem(key);
```

#### パラメータ
| パラメーター   | 型|説明|
|:---------------|:--------|:----------|
|key|string|ワークシートの名前または ID。|

#### 戻り値
[Worksheet](worksheet.md)
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
	var worksheets = ctx.workbook.worksheets;
	worksheets.load('items');
	return ctx.sync().then(function() {
		for (var i = 0; i < worksheets.items.length; i++)
		{
			console.log(worksheets.items[i].name);
			console.log(worksheets.items[i].index);
		}
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

