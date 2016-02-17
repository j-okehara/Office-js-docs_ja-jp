# TableRowCollection オブジェクト (JavaScript API for Excel)

_適用対象: Excel 2016、Excel Online、Office 2016_

表の一部であるすべての行のコレクションを表します。

## プロパティ

| プロパティ   | 型|説明
|:---------------|:--------|:----------|
|count|int|テーブルの行数を返します。読み取り専用。|
|Items|[TableRow[]](tablerow.md)|tableRow オブジェクトのコレクション。読み取り専用。|

_プロパティのアクセスの[例](#property-access-examples)をご覧ください。_

## 関係
なし


## メソッド

| メソッド   | 戻り値の型|説明|
|:---------------|:--------|:----------|
|[add(index: number, values: (boolean、string または number)[][])](#addindex-number-values-boolean-or-string-or-number)|[TableRow](tablerow.md)|新しい行をテーブルに追加します。|
|[getItemAt(index: number)](#getitematindex-number)|[TableRow](tablerow.md)|コレクション内の位置を基に行を取得します。|
|[load(param: object)](#loadparam-object)|void|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|

## メソッドの詳細

### add(index: number, values: (boolean、string または number)[][])
新しい行をテーブルに追加します。

#### 構文
```js
tableRowCollectionObject.add(index, values);
```

#### パラメータ
| パラメーター   | 型|説明|
|:---------------|:--------|:----------|
|index|number|省略可能。新しい行の相対位置を指定します。null の場合、末尾に追加されます。挿入した行の下のすべての行が下方向にシフトします。0 を起点とする番号になります。|
|values|(boolean、string または number)[][]|省略可能。テーブルの行の書式設定されていない値の 2 次元の配列。|

#### 戻り値
[TableRow](tablerow.md)

#### 例

```js
Excel.run(function (ctx) { 
	var tables = ctx.workbook.tables;
	var values = [["Sample", "Values", "For", "New", "Row"]];
	var row = tables.getItem("Table1").rows.add(null, values);
	row.load('index');
	return ctx.sync().then(function() {
		console.log(row.index);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
### getItemAt(index: number)
コレクション内の位置を基に行を取得します。

#### 構文
```js
tableRowCollectionObject.getItemAt(index);
```

#### パラメータ
| パラメーター   | 型|説明|
|:---------------|:--------|:----------|
|index|number|取得するオブジェクトのインデックス値。0 を起点とする番号になります。|

#### 戻り値
[TableRow](tablerow.md)

#### 例

```js
Excel.run(function (ctx) { 
	var tablerow = ctx.workbook.tables.getItem('Table1').rows.getItemAt(0);
	tablerow.load('name');
	return ctx.sync().then(function() {
			console.log(tablerow.name);
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
	var tablerows = ctx.workbook.tables.getItem('Table1').rows;
	tablerows.load('items');
	return ctx.sync().then(function() {
		console.log("tablerows Count: " + tablerows.count);
		for (var i = 0; i < tablerows.items.length; i++)
		{
			console.log(tablerows.items[i].index);
		}
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
