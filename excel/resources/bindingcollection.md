# BindingCollection オブジェクト (JavaScript API for Excel)

_適用対象: Excel 2016、Excel Online、Office 2016_

ブックの一部であるすべてのバインド オブジェクトのコレクションを表します。

## プロパティ

| プロパティ   | 型|説明
|:---------------|:--------|:----------|
|count|int|コレクション内にあるバインドの数を取得します。値の取得のみ可能です。|
|Items|[Binding[]](binding.md)|バインド オブジェクトのコレクション。読み取り専用。|

_プロパティのアクセスの[例](#property-access-examples)を参照してください。_

## 関係
なし


## メソッド

| メソッド   | 戻り値の型|説明|
|:---------------|:--------|:----------|
|[getItem(id: string)](#getitemid-string)|[binding](binding.md)|ID によってバインド オブジェクトを取得します。|
|[getItemAt(index: number)](#getitematindex-number)|[binding](binding.md)|項目の配列内の位置に基づいて、バインド オブジェクトを取得します。|
|[load(param: object)](#loadparam-object)|void|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|

## メソッドの詳細

### getItem(id: string)
ID によってバインド オブジェクトを取得します。

#### 構文
```js
bindingCollectionObject.getItem(id);
```

#### パラメータ
| パラメーター   | 型|説明|
|:---------------|:--------|:----------|
|id|string|取得するバインド オブジェクトの ID。|

#### 戻り値
[binding](binding.md)

#### 例

表のデータ変更をモニターするための表バインドを作成します。データが変更されると、表の背景色がオレンジ色に変更されます。

```js
function addEventHandler() {
	//Create Table1
Excel.run(function (ctx) { 
	ctx.workbook.tables.add("Sheet1!A1:C4", true);
	return ctx.sync().then(function() {
			 console.log("My Diet Data Inserted!");
	})
	.catch(function (error) {
			 console.log(JSON.stringify(error));
	});
});
	//Create a new table binding for Table1
Office.context.document.bindings.addFromNamedItemAsync("Table1", Office.CoercionType.Table, { id: "myBinding" }, function (asyncResult) {
	if (asyncResult.status == "failed") {
		console.log("Action failed with error: " + asyncResult.error.message);
	}
	else {
		// If successful, add the event handler to the table binding.
		Office.select("bindings#myBinding").addHandlerAsync(Office.EventType.BindingDataChanged, onBindingDataChanged);
	}
});
}
	
// When data in the table is changed, this event is triggered.
function onBindingDataChanged(eventArgs) {
Excel.run(function (ctx) { 
	// Highlight the table in orange to indicate data changed.
	ctx.workbook.bindings.getItem(eventArgs.binding.id).getTable().getDataBodyRange().format.fill.color = "Orange";
	return ctx.sync().then(function() {
			console.log("The value in this table got changed!");
	})
	.catch(function (error) {
			console.log(JSON.stringify(error));
	});
});
}

```



#### 例
```js
Excel.run(function (ctx) { 
	var lastPosition = ctx.workbook.bindings.count - 1;
	var binding = ctx.workbook.bindings.getItemAt(lastPosition);
	binding.load('type')
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

### getItemAt(index: number)
項目の配列内の位置に基づいて、バインド オブジェクトを取得します。

#### 構文
```js
bindingCollectionObject.getItemAt(index);
```

#### パラメータ
| パラメーター   | 型|説明|
|:---------------|:--------|:----------|
|index|number|取得するオブジェクトのインデックス値。0 を起点とする番号になります。|

#### 戻り値
[binding](binding.md)

#### 例
```js
Excel.run(function (ctx) { 
	var lastPosition = ctx.workbook.bindings.count - 1;
	var binding = ctx.workbook.bindings.getItemAt(lastPosition);
	binding.load('type')
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
	var bindings = ctx.workbook.bindings;
	bindings.load('items');
	return ctx.sync().then(function() {
		for (var i = 0; i < bindings.items.length; i++)
		{
			console.log(bindings.items[i].id);
			console.log(bindings.items[i].index);
		}
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
バインドの数を取得します。

```js
Excel.run(function (ctx) { 
	var bindings = ctx.workbook.bindings;
	bindings.load('count');
	return ctx.sync().then(function() {
		console.log("Bindings: Count= " + bindings.count);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

