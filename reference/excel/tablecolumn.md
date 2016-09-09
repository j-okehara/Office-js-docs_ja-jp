# TableColumn オブジェクト (JavaScript API for Excel)

テーブル内にある 1 つの列を表します。

## プロパティ

| プロパティ     | 型   |説明
|:---------------|:--------|:----------|
|id|int|表内の列を識別する一意のキーを返します。読み取り専用です。|
|index|int|表の列コレクション内の列のインデックス番号を返します。0 を起点とする番号になります。読み取り専用です。|
|name|string|テーブル列の名前を取得します。読み取り専用です。|
|values|object[][]|指定した範囲の Raw 値を表します。返されるデータの型は、文字列、数値、またはブール値のいずれかになります。エラーが含まれているセルは、エラー文字列を返します。|

_プロパティのアクセスの[例](#例)をご覧ください。_

## リレーションシップ
| リレーションシップ | 型   |説明|
|:---------------|:--------|:----------|
|フィルター|[フィルター](filter.md)|列に適用されるフィルターを取得します。読み取り専用です。|

## メソッド

| メソッド           | 戻り値の型    |説明|
|:---------------|:--------|:----------|
|[delete()](#delete)|void|テーブルから列を削除します。|
|[getDataBodyRange()](#getdatabodyrange)|[範囲](range.md)|列のデータ本体に関連付けられた範囲オブジェクトを取得します。|
|[getHeaderRowRange()](#getheaderrowrange)|[範囲](range.md)|列のヘッダー行に関連付けられた範囲オブジェクトを取得します。|
|[getRange()](#getrange)|[範囲](range.md)|列全体に関連付けられた範囲オブジェクトを取得します。|
|[getTotalRowRange()](#gettotalrowrange)|[範囲](range.md)|列の集計行に関連付けられた範囲オブジェクトを取得します。|
|[load(param: object)](#loadparam-object)|void|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|

## メソッドの詳細


### delete()
テーブルから列を削除します。

#### 構文
```js
tableColumnObject.delete();
```

#### パラメーター
なし

#### 戻り値
void

#### 例

```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var column = ctx.workbook.tables.getItem(tableName).tableColumns.getItemAt(2);
    column.delete();
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### getDataBodyRange()
列のデータ本体に関連付けられた範囲オブジェクトを取得します。

#### 構文
```js
tableColumnObject.getDataBodyRange();
```

#### パラメーター
なし

#### 戻り値
[範囲](range.md)

#### 例

```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var column = ctx.workbook.tables.getItem(tableName).tableColumns.getItemAt(0);
    var dataBodyRange = column.getDataBodyRange();
    dataBodyRange.load('address');
    return ctx.sync().then(function() {
        console.log(dataBodyRange.address);
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### getHeaderRowRange()
列のヘッダー行に関連付けられた範囲オブジェクトを取得します。

#### 構文
```js
tableColumnObject.getHeaderRowRange();
```

#### パラメーター
なし

#### 戻り値
[範囲](range.md)

#### 例

```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var columns = ctx.workbook.tables.getItem(tableName).tableColumns.getItemAt(0);
    var headerRowRange = columns.getHeaderRowRange();
    headerRowRange.load('address');
    return ctx.sync().then(function() {
        console.log(headerRowRange.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### getRange()
列全体に関連付けられた範囲オブジェクトを取得します。

#### 構文
```js
tableColumnObject.getRange();
```

#### パラメーター
なし

#### 戻り値
[範囲](range.md)

#### 例

```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var columns = ctx.workbook.tables.getItem(tableName).tableColumns.getItemAt(0);
    var columnRange = columns.getRange();
    columnRange.load('address');
    return ctx.sync().then(function() {
        console.log(columnRange.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### getTotalRowRange()
列の集計行に関連付けられた範囲オブジェクトを取得します。

#### 構文
```js
tableColumnObject.getTotalRowRange();
```

#### パラメーター
なし

#### 戻り値
[範囲](range.md)

#### 例

```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var columns = ctx.workbook.tables.getItem(tableName).tableColumns.getItemAt(0);
    var totalRowRange = columns.getTotalRowRange();
    totalRowRange.load('address');
    return ctx.sync().then(function() {
        console.log(totalRowRange.address);
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
    var tableName = 'Table1';
    var column = ctx.workbook.tables.getItem(tableName).tableColumns.getItem(0);
    column.load('index');
    return ctx.sync().then(function() {
        console.log(column.index);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

```js
Excel.run(function (ctx) { 
    var tables = ctx.workbook.tables;
    var newValues = [["New"], ["Values"], ["For"], ["New"], ["Column"]];
    var column = ctx.workbook.tables.getItem(tableName).tableColumns.getItemAt(2);
    column.values = newValues;
    column.load('values');
    return ctx.sync().then(function() {
        console.log(column.values);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
