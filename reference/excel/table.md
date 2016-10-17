# <a name="table-object-(javascript-api-for-excel)"></a>Table オブジェクト (JavaScript API for Excel)

_適用対象:Excel 2016、Excel Online、Excel for iOS、Office 2016_

Excel の表を表します。

## <a name="properties"></a>プロパティ

| プロパティ     | 型   |説明
|:---------------|:--------|:----------|
|id|int|指定されたブックのテーブルを一意に識別する値を返します。識別子の値は、テーブルの名前が変更された場合も変わりません。読み取り専用です。|
|name|string|テーブルの名前。|
|showHeaders|bool|ヘッダー行を表示するかどうかを示します。この値によって、ヘッダー行の表示または削除を設定できます。|
|showTotals|bool|集計行を表示するかどうかを示します。この値によって、集計行の表示または削除を設定できます。|
|style|string|表スタイルを表す定数値。使用可能な値は次のとおりです: TableStyleLight1 から TableStyleLight21、TableStyleMedium1 から TableStyleMedium28、TableStyleStyleDark1 から TableStyleStyleDark11。ブックに存在するカスタムのユーザー定義スタイルも指定できます。|

_プロパティのアクセスの[例を参照してください。](#property-access-examples)_

## <a name="relationships"></a>関係
| リレーションシップ | 型   |説明|
|:---------------|:--------|:----------|
|columns|[TableColumnCollection](tablecolumncollection.md)|テーブルに含まれるすべての列のコレクションを表します。読み取り専用です。|
|rows|[TableRowCollection](tablerowcollection.md)|テーブルに含まれるすべての行のコレクションを表します。読み取り専用です。|
|sort|[TableSort](tablesort.md)|テーブル内のソート順の構成を表します。読み取り専用です。|
|worksheet|[Worksheet](worksheet.md)|現在の表を含んでいるワークシート。読み取り専用です。|

## <a name="methods"></a>メソッド

| メソッド           | 戻り値の型    |説明|
|:---------------|:--------|:----------|
|[clearFilters()](#clearfilters)|void|現在テーブルに適用されているすべてのフィルターをクリアします。|
|[convertToRange()](#converttorange)|[Range](range.md)|テーブルを通常の範囲のセルに変換します。すべてのデータが保持されます。|
|[delete()](#delete)|void|テーブルを削除します。|
|[getDataBodyRange()](#getdatabodyrange)|[Range](range.md)|テーブルのデータ本体に関連付けられた範囲オブジェクトを取得します。|
|[getHeaderRowRange()](#getheaderrowrange)|[Range](range.md)|テーブルのヘッダー行に関連付けられた範囲オブジェクトを取得します。|
|[getRange()](#getrange)|[Range](range.md)|テーブル全体に関連付けられた範囲オブジェクトを取得します。|
|[getTotalRowRange()](#gettotalrowrange)|[Range](range.md)|表の集計行に関連付けられた範囲オブジェクトを取得します。|
|[load(param: object)](#loadparam-object)|void|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|
|[reapplyFilters()](#reapplyfilters)|void|現在テーブルにあるすべてのフィルターを再適用します。|

## <a name="method-details"></a>メソッドの詳細


### <a name="clearfilters()"></a>clearFilters()
現在テーブルに適用されているすべてのフィルターをクリアします。

#### <a name="syntax"></a>構文
```js
tableObject.clearFilters();
```

#### <a name="parameters"></a>パラメーター
なし

#### <a name="returns"></a>戻り値
void

### <a name="converttorange()"></a>convertToRange()
テーブルを通常の範囲のセルに変換します。すべてのデータが保持されます。

#### <a name="syntax"></a>構文
```js
tableObject.convertToRange();
```

#### <a name="parameters"></a>パラメーター
なし

#### <a name="returns"></a>戻り値
[Range](range.md)

#### <a name="examples"></a>例
```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    table.convertToRange();
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### <a name="delete()"></a>delete()
テーブルを削除します。

#### <a name="syntax"></a>構文
```js
tableObject.delete();
```

#### <a name="parameters"></a>パラメーター
なし

#### <a name="returns"></a>戻り値
void

#### <a name="examples"></a>例
```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    table.delete();
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getdatabodyrange()"></a>getDataBodyRange()
テーブルのデータ本体に関連付けられた範囲オブジェクトを取得します。

#### <a name="syntax"></a>構文
```js
tableObject.getDataBodyRange();
```

#### <a name="parameters"></a>パラメーター
なし

#### <a name="returns"></a>戻り値
[Range](range.md)

#### <a name="examples"></a>例
```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    var tableDataRange = table.getDataBodyRange();
    tableDataRange.load('address')
    return ctx.sync().then(function() {
            console.log(tableDataRange.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### <a name="getheaderrowrange()"></a>getHeaderRowRange()
表のヘッダー行に関連付けられた範囲オブジェクトを取得します。

#### <a name="syntax"></a>構文
```js
tableObject.getHeaderRowRange();
```

#### <a name="parameters"></a>パラメーター
なし

#### <a name="returns"></a>戻り値
[Range](range.md)

#### <a name="examples"></a>例
```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    var tableHeaderRange = table.getHeaderRowRange();
    tableHeaderRange.load('address');
    return ctx.sync().then(function() {
        console.log(tableHeaderRange.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getrange()"></a>getRange()
テーブル全体に関連付けられた範囲オブジェクトを取得します。

#### <a name="syntax"></a>構文
```js
tableObject.getRange();
```

#### <a name="parameters"></a>パラメーター
なし

#### <a name="returns"></a>戻り値
[Range](range.md)

#### <a name="examples"></a>例
```js
Excel.run(function (ctx) { 
    var table = ctx.workbook.tables.getItem(tableName);
    var tableRange = table.getRange();
    tableRange.load('address'); 
    return ctx.sync().then(function() {
            console.log(tableRange.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="gettotalrowrange()"></a>getTotalRowRange()
表の集計行に関連付けられた範囲オブジェクトを取得します。

#### <a name="syntax"></a>構文
```js
tableObject.getTotalRowRange();
```

#### <a name="parameters"></a>パラメーター
なし

#### <a name="returns"></a>戻り値
[Range](range.md)

#### <a name="examples"></a>例
```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    var tableTotalsRange = table.getTotalRowRange();
    tableTotalsRange.load('address');   
    return ctx.sync().then(function() {
            console.log(tableTotalsRange.address);
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

名前でテーブルを取得します。 

```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    table.load('index')
    return ctx.sync().then(function() {
            console.log(table.index);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

インデックスでテーブルを取得します。

```js
Excel.run(function (ctx) { 
    var index = 0;
    var table = ctx.workbook.tables.getItemAt(0);
    table.name('name')
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

テーブル スタイルを設定します。 

```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    table.name = 'Table1-Renamed';
    table.showTotals = false;
    table.tableStyle = 'TableStyleMedium2';
    table.load('tableStyle');
    return ctx.sync().then(function() {
            console.log(table.tableStyle);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
