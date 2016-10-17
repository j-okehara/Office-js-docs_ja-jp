# <a name="tablerowcollection-object-(javascript-api-for-excel)"></a>TableRowCollection オブジェクト (JavaScript API for Excel)

表の一部であるすべての行のコレクションを表します。

## <a name="properties"></a>プロパティ

| プロパティ     | 型   |説明
|:---------------|:--------|:----------|
|count|int|テーブルの行数を返します。読み取り専用。|
|items|[TableRow[]](tablerow.md)|TableRow オブジェクトのコレクション。読み取り専用です。|

_プロパティのアクセスの[例を参照してください。](#property-access-examples)_

## <a name="relationships"></a>関係
なし


## <a name="methods"></a>メソッド

| メソッド           | 戻り値の型    |説明|
|:---------------|:--------|:----------|
|[add(index: number, values: (boolean、string、または number)[][])](#addindex-number-values-boolean-or-string-or-number)|[TableRow](tablerow.md)|新しい行をテーブルに追加します。|
|[getItemAt(index: number)](#getitematindex-number)|[TableRow](tablerow.md)|コレクション内の位置を基に行を取得します。|
|[load(param: object)](#loadparam-object)|void|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|

## <a name="method-details"></a>メソッドの詳細


### <a name="add(index:-number,-values:-(boolean-or-string-or-number)[][])"></a>add(index: number, values: (boolean、string または number)[][])
新しい行をテーブルに追加します。

#### <a name="syntax"></a>構文
```js
tableRowCollectionObject.add(index, values);
```

#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|index|number|省略可能。新しい行の相対位置を指定します。null の場合、末尾に追加されます。挿入した行の下のすべての行が下方向にシフトします。0 を起点とする番号になります。|
|values|(boolean、string または number)[][]|省略可能。テーブルの行の書式設定されていない値の 2 次元の配列。|

#### <a name="returns"></a>戻り値
[TableRow](tablerow.md)

#### <a name="examples"></a>例

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

### <a name="getitemat(index:-number)"></a>getItemAt(index: number)
コレクション内の位置を基に行を取得します。

#### <a name="syntax"></a>構文
```js
tableRowCollectionObject.getItemAt(index);
```

#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|index|number|取得するオブジェクトのインデックス値。0 を起点とする番号になります。|

#### <a name="returns"></a>戻り値
[TableRow](tablerow.md)

#### <a name="examples"></a>例

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

### <a name="load(param:-object)"></a>load(param: object)
JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。

#### <a name="syntax"></a>構文
```js
object.load(param);
```

#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|param|object|省略可能。パラメーター名とリレーションシップ名を、区切られた文字列または 1 つの配列として受け入れます。あるいは、[loadOption](loadoption.md) オブジェクトを提供します。|

#### <a name="returns"></a>戻り値
void
### <a name="property-access-examples"></a>プロパティのアクセスの例

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
