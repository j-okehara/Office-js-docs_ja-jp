# TableColumnCollection オブジェクト (JavaScript API for Excel)

表の一部であるすべての列のコレクションを表します。

## プロパティ

| プロパティ     | 型   |説明
|:---------------|:--------|:----------|
|count|int|テーブルの列数を返します。読み取り専用です。|
|Items|[TableColumn[]](tablecolumn.md)|tableColumn オブジェクトのコレクション。読み取り専用です。|

_プロパティのアクセスの[例](#例)をご覧ください。_

## リレーションシップ
なし


## メソッド

| メソッド           | 戻り値の型    |説明|
|:---------------|:--------|:----------|
|[add(index: number, values: (boolean、string または number)[][])](#addindex-number-values-boolean-string-または-number)|[TableColumn](tablecolumn.md)|テーブルに新しい列を追加します。|
|[getItem(key: number またはstring)](#getitemkey-number-またはstring)|[TableColumn](tablecolumn.md)|名前または ID によって、列オブジェクトを取得します。|
|[getItemAt(index: number)](#getitematindex-number)|[TableColumn](tablecolumn.md)|コレクション内の位置に基づいて列を取得します。|
|[load(param: object)](#loadparam-object)|void|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|

## メソッドの詳細


### add(index: number, values: (boolean、string または number)[][])
テーブルに新しい列を追加します。

#### 構文
```js
tableColumnCollectionObject.add(index, values);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|index|number|新しい列の相対位置を指定します。この位置の前の列は右にシフトされます。インデックス値は、最後の列のインデックス値と等しいか、小さくなります。そのため、表の末尾に列を追加するためには使用できません。0 を起点とする番号になります。|
|values|(boolean、string または number)[][]|省略可能。テーブルの列の、書式設定されていない値の 2 次元の配列。|

#### 戻り値
[TableColumn](tablecolumn.md)

#### 例

```js
Excel.run(function (ctx) { 
    var tables = ctx.workbook.tables;
    var values = [["Sample"], ["Values"], ["For"], ["New"], ["Column"]];
    var column = tables.getItem("Table1").columns.add(null, values);
    column.load('name');
    return ctx.sync().then(function() {
        console.log(column.name);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### getItem(key: number またはstring)
名前または ID によって、列オブジェクトを取得します。

#### 構文
```js
tableColumnCollectionObject.getItem(key);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|Key|number または string| 列名または ID。|

#### 戻り値
[TableColumn](tablecolumn.md)

#### 例

```js
Excel.run(function (ctx) { 
    var tablecolumn = ctx.workbook.tables.getItem['Table1'].columns.getItem(0);
    tablecolumn.load('name');
    return ctx.sync().then(function() {
            console.log(tablecolumn.name);
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
    var tablecolumn = ctx.workbook.tables.getItem['Table1'].columns.getItemAt(0);
    tablecolumn.load('name');
    return ctx.sync().then(function() {
            console.log(tablecolumn.name);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### getItemAt(index: number)
コレクション内の位置に基づいて列を取得します。

#### 構文
```js
tableColumnCollectionObject.getItemAt(index);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|index|number|取得するオブジェクトのインデックス値。0 を起点とする番号になります。|

#### 戻り値
[TableColumn](tablecolumn.md)

#### 例
```js
Excel.run(function (ctx) { 
    var tablecolumn = ctx.workbook.tables.getItem['Table1'].columns.getItemAt(0);
    tablecolumn.load('name');
    return ctx.sync().then(function() {
            console.log(tablecolumn.name);
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

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|param|object|省略可能。パラメーター名とリレーションシップ名を、区切られた文字列または 1 つの配列として受け入れます。あるいは、[loadOption](loadoption.md) オブジェクトを提供します。|

#### 戻り値
void
### プロパティのアクセスの例

```js
Excel.run(function (ctx) { 
    var tablecolumns = ctx.workbook.tables.getItem['Table1'].columns;
    tablecolumns.load('items');
    return ctx.sync().then(function() {
        console.log("tablecolumns Count: " + tablecolumns.count);
        for (var i = 0; i < tablecolumns.items.length; i++)
        {
            console.log(tablecolumns.items[i].name);
        }
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
