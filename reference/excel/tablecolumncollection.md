# <a name="tablecolumncollection-object-javascript-api-for-excel"></a>TableColumnCollection オブジェクト (JavaScript API for Excel)

表の一部であるすべての列のコレクションを表します。

## <a name="properties"></a>プロパティ

| プロパティ       | 型    |説明| 要件セット|
|:---------------|:--------|:----------|:----|
|count|int|テーブルの列数を返します。読み取り専用です。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|items|[TableColumn[]](tablecolumn.md)|tableColumn オブジェクトのコレクション。読み取り専用です。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_プロパティのアクセスの[例を参照してください。](#property-access-examples)_

## <a name="relationships"></a>関係
なし


## <a name="methods"></a>メソッド

| メソッド           | 戻り値の型    |説明| 要件セット|
|:---------------|:--------|:----------|:----|
|[add(index: number, values: (boolean または string または number)[][], name: string)](#addindex-number-values-boolean-or-string-or-number-name-string)|[TableColumn](tablecolumn.md)|テーブルに新しい列を追加します。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getCount()](#getcount)|int|表の列数を取得します。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getItem(key: number または string)](#getitemkey-number-or-string)|[TableColumn](tablecolumn.md)|名前または ID を使用して列オブジェクトを取得します。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemAt(index: number)](#getitematindex-number)|[TableColumn](tablecolumn.md)|コレクション内の位置に基づいて列を取得します。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemOrNullObject(key: number または string)](#getitemornullobjectkey-number-or-string)|[TableColumn](tablecolumn.md)|名前または ID によって、列オブジェクトを取得します。列が存在しない場合は null オブジェクトを返します。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>メソッドの詳細


### <a name="addindex-number-values-boolean-or-string-or-number-name-string"></a>add(index: number, values: (boolean または string または number)[][], name: string)
テーブルに新しい列を追加します。

#### <a name="syntax"></a>構文
```js
tableColumnCollectionObject.add(index, values, name);
```

#### <a name="parameters"></a>パラメーター
| パラメーター       | 型    |説明|
|:---------------|:--------|:----------|:---|
|index|number|省略可能。新しい列の相対位置を指定します。null または -1 の場合、最後に追加が行われます。上位のインデックスを持つ列は横にシフトされます。0 を起点とする番号になります。|
|values|(boolean、string または number)[][]|省略可能。テーブルの列の、書式設定されていない値の 2 次元の配列。|
|name|string|省略可能。新しい列の名前を指定します。null の場合は既定の名前が使用されます。|

#### <a name="returns"></a>戻り値
[TableColumn](tablecolumn.md)

#### <a name="examples"></a>例

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


### <a name="getcount"></a>getCount()
表の列数を取得します。

#### <a name="syntax"></a>構文
```js
tableColumnCollectionObject.getCount();
```

#### <a name="parameters"></a>パラメーター
なし

#### <a name="returns"></a>戻り値
int

### <a name="getitemkey-number-or-string"></a>getItem(key: number または string)
名前または ID によって、列オブジェクトを取得します。

#### <a name="syntax"></a>構文
```js
tableColumnCollectionObject.getItem(key);
```

#### <a name="parameters"></a>パラメーター
| パラメーター       | 型    |説明|
|:---------------|:--------|:----------|:---|
|Key|number または string| 列名または ID。|

#### <a name="returns"></a>戻り値
[TableColumn](tablecolumn.md)

#### <a name="examples"></a>例

```js
Excel.run(function (ctx) { 
    var tablecolumn = ctx.workbook.tables.getItem('Table1').columns.getItem(0);
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


#### <a name="examples"></a>例
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

### <a name="getitematindex-number"></a>getItemAt(index: number)
コレクション内の位置に基づいて列を取得します。

#### <a name="syntax"></a>構文
```js
tableColumnCollectionObject.getItemAt(index);
```

#### <a name="parameters"></a>パラメーター
| パラメーター       | 型    |説明|
|:---------------|:--------|:----------|:---|
|index|number|取得するオブジェクトのインデックス値。0 を起点とする番号になります。|

#### <a name="returns"></a>戻り値
[TableColumn](tablecolumn.md)

#### <a name="examples"></a>例
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

### <a name="getitemornullobjectkey-number-or-string"></a>getItemOrNullObject(key: number または string)
名前または ID によって、列オブジェクトを取得します。列が存在しない場合は null オブジェクトを返します。

#### <a name="syntax"></a>構文
```js
tableColumnCollectionObject.getItemOrNullObject(key);
```

#### <a name="parameters"></a>パラメーター
| パラメーター       | 型    |説明|
|:---------------|:--------|:----------|:---|
|Key|number または string| 列名または ID。|

#### <a name="returns"></a>戻り値
[TableColumn](tablecolumn.md)
### <a name="property-access-examples"></a>プロパティのアクセスの例

```js
Excel.run(function (ctx) { 
    var tablecolumns = ctx.workbook.tables.getItem('Table1').columns;
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