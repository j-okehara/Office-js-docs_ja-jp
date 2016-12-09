# <a name="tablecollection-object-javascript-api-for-excel"></a>TableCollection オブジェクト (JavaScript API for Excel)

ブックの一部として含まれている、すべての表のコレクションを表します。

## <a name="properties"></a>プロパティ

| プロパティ     | 型   |説明| 要件セット|
|:---------------|:--------|:----------|:----|
|count|int|ブックに含まれるテーブルの数を返します。読み取り専用です。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|items|[Table[]](table.md)|Table オブジェクトのコレクション。読み取り専用です。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_プロパティのアクセスの[例を参照してください。](#property-access-examples)_

## <a name="relationships"></a>関係
なし


## <a name="methods"></a>メソッド

| メソッド           | 戻り値の型    |説明| 要件セット|
|:---------------|:--------|:----------|:----|
|[add(address:Range or string, hasHeaders: bool)](#addaddress-range-or-string-hasheaders-bool)|[Table](table.md)|新しいテーブルを作成します。範囲オブジェクトまたはソース アドレスにより、テーブルが追加されるワークシートが判断されます。テーブルが追加できない場合 (たとえば、アドレスが無効な場合や、テーブルが別のテーブルと重複している場合) は、エラーがスローされます。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItem(key: number or string)](#getitemkey-number-or-string)|[Table](table.md)|名前または ID を使用してテーブルを取得します。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemAt(index: number)](#getitematindex-number)|[Table](table.md)|コレクション内の位置に基づいてテーブルを取得します。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemOrNull(key: number or string)](#getitemornullkey-number-or-string)|[Table](table.md)|名前または ID を使用してテーブルを取得します。テーブルが存在しない場合、戻りオブジェクトの isNull プロパティは true になります。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|[load(param: object)](#loadparam-object)|(非推奨)|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>メソッドの詳細


### <a name="addaddress-range-or-string-hasheaders-bool"></a>add(address:Range or string, hasHeaders: bool)
新しいテーブルを作成します。範囲オブジェクトまたはソース アドレスにより、テーブルが追加されるワークシートが判断されます。テーブルが追加できない場合 (たとえば、アドレスが無効な場合や、テーブルが別のテーブルと重複している場合) は、エラーがスローされます。

#### <a name="syntax"></a>構文
```js
tableCollectionObject.add(address, hasHeaders);
```

#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|:---|
|address|Range または string|Range オブジェクト、あるいはデータ ソースを表す範囲の文字列アドレスまたは名前。アドレスにシート名が含まれていない場合は、現在作業中のシートが使用されます。文字列パラメーターには要件セット 1.1、Range オブジェクトの受け入れには 1.3 が必要です。|
|hasHeaders|bool|インポートされたデータに列ラベルがあるかどうかを示すブール値。ソースにヘッダーが含まれていない場合 (このプロパティが false に設定されている場合)、Excel はデータを下方向に 1 行シフトして、自動的にヘッダーを生成します。|

#### <a name="returns"></a>戻り値
[Table](table.md)

#### <a name="examples"></a>例

```js
Excel.run(function (ctx) { 
    var table = ctx.workbook.tables.add('Sheet1!A1:E7', true);
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

### <a name="getitemkey-number-or-string"></a>getItem(key: number またはstring)
名前または ID でテーブルを取得します。

#### <a name="syntax"></a>構文
```js
tableCollectionObject.getItem(key);
```

#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|:---|
|Key|number または string|取得するテーブルの名前または ID。|

#### <a name="returns"></a>戻り値
[Table](table.md)

#### <a name="examples"></a>例

```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
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


#### <a name="examples"></a>例

```js
Excel.run(function (ctx) { 
    var table = ctx.workbook.tables.getItemAt(0);
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


### <a name="getitematindex-number"></a>getItemAt(index: number)
コレクション内の位置に基づいてテーブルを取得します。

#### <a name="syntax"></a>構文
```js
tableCollectionObject.getItemAt(index);
```

#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|:---|
|index|number|取得するオブジェクトのインデックス値。0 を起点とする番号になります。|

#### <a name="returns"></a>戻り値
[Table](table.md)

#### <a name="examples"></a>例

```js
Excel.run(function (ctx) { 
    var table = ctx.workbook.tables.getItemAt(0);
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


### <a name="getitemornullkey-number-or-string"></a>getItemOrNull(key: number or string)
名前または ID を使用してテーブルを取得します。テーブルが存在しない場合、戻りオブジェクトの isNull プロパティは true になります。

#### <a name="syntax"></a>構文
```js
tableCollectionObject.getItemOrNull(key);
```

#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|:---|
|Key|number または string|取得するテーブルの名前または ID。|

#### <a name="returns"></a>戻り値
[Table](table.md)

### <a name="loadparam-object"></a>load(param: object)
JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。

#### <a name="syntax"></a>構文
```js
object.load(param);
```

#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|:---|
|param|object|省略可能。パラメーターとリレーションシップ名を、区切られた文字列または 1 つの配列として受け入れます。あるいは、[loadOption](loadoption.md) オブジェクトを提供します。|

#### <a name="returns"></a>戻り値
void
### <a name="property-access-examples"></a>プロパティのアクセスの例

```js
Excel.run(function (ctx) { 
    var tables = ctx.workbook.tables;
    tables.load();
    return ctx.sync().then(function() {
        console.log("tables Count: " + tables.count);
        for (var i = 0; i < tables.items.length; i++)
        {
            console.log(tables.items[i].name);
        }
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

テーブルの数を取得します

```js
Excel.run(function (ctx) { 
    var tables = ctx.workbook.tables;
    tables.load('count');
    return ctx.sync().then(function() {
        console.log(tables.count);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```