# Binding オブジェクト (JavaScript API for Excel)

_適用対象: Excel 2016、Excel Online、Office 2016_

ブックで定義されている Office.js のバインディングを表します。

## プロパティ

| プロパティ   | 型|説明
|:---------------|:--------|:----------|
|id|string|バインド識別子を表します。読み取り専用。|
|type|string|バインドの型を返します。読み取り専用。使用可能な値は次のとおりです。Range, Table, Text。|

プロパティのアクセスの[例](#property-access-examples)をご覧ください。

## 関係
なし


## メソッド

| メソッド   | 戻り値の型|説明|
|:---------------|:--------|:----------|
|[getRange()](#getrange)|[Range](range.md)|バインディングによって表される範囲を返します。バインドが正しい型ではない場合、エラーがスローされます。|
|[getTable()](#gettable)|[Table](table.md)|バインドによって表されるテーブルを返します。バインドが正しい型ではない場合、エラーがスローされます。|
|[getText()](#gettext)|string|バインドによって表されるテキストを返します。バインドが正しい型ではない場合、エラーがスローされます。|
|[load(param: object)](#loadparam-object)|void|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|

## メソッドの詳細

### getRange()
バインディングによって表される範囲を返します。バインドが正しい型ではない場合、エラーがスローされます。

#### 構文
```js
bindingObject.getRange();
```

#### パラメーター
なし

#### 戻り値
[Range](range.md)

#### 例
以下の例では、バインド オブジェクトを使用して、関連付けられている範囲を取得しています。

```js
Excel.run(function (ctx) { 
	var binding = ctx.workbook.bindings.getItemAt(0);
	var range = binding.getRange();
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

### getTable()
バインドによって表されるテーブルを返します。バインドが正しい型ではない場合、エラーがスローされます。

#### 構文
```js
bindingObject.getTable();
```

#### パラメーター
なし

#### 戻り値
[Table](table.md)

#### 例
```js
Excel.run(function (ctx) { 
	var binding = ctx.workbook.bindings.getItemAt(0);
	var table = binding.getTable();
	table.load('name');
	return ctx.sync().then(function() {
			console.log(table.name);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### getText()
バインドによって表されるテキストを返します。バインドが正しい型ではない場合、エラーがスローされます。

#### 構文
```js
bindingObject.getText();
```

#### パラメーター
なし

#### 戻り値
文字列

#### 例

```js
Excel.run(function (ctx) { 
	var binding = ctx.workbook.bindings.getItemAt(0);
	var text = binding.getText();
	ctx.load('text');
	return ctx.sync().then(function() {
		console.log(text);
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
|param|object|省略可能。パラメーター名とリレーションシップ名を、区切られた文字列または 1 つの配列として受け入れます。あるいは、[loadOption](loadoption.md) オブジェクトを受け入れます。|

#### 戻り値
void
### プロパティのアクセスの例

```js
Excel.run(function (ctx) { 
	var binding = ctx.workbook.bindings.getItemAt(0);
	binding.load('type');
	return ctx.sync().then(function() {
		console.log(binding.type);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

