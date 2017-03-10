# <a name="tablecollection-object-javascript-api-for-excel"></a>TableCollection オブジェクト (JavaScript API for Excel)

ブックまたはワークシートの一部として含まれる、すべてのテーブルのコレクションを、到達方法に応じて表します。

## <a name="properties"></a>プロパティ

| プロパティ       | 型    |説明| 要件セット|
|:---------------|:--------|:----------|:----|
|count|int|ブックに含まれるテーブルの数を返します。読み取り専用です。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|items|[Table[]](table.md)|Table オブジェクトのコレクション。読み取り専用です。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_プロパティのアクセスの[例を参照してください。](#property-access-examples)_

## <a name="relationships"></a>関係
なし


## <a name="methods"></a>メソッド

| メソッド           | 戻り値の型    |説明| 要件セット|
|:---------------|:--------|:----------|:----|
|[add(address: [object, hasHeaders: bool)](#addaddress-object-hasheaders-bool)|[Table](table.md)|新しいテーブルを作成します。範囲オブジェクトまたはソース アドレスにより、テーブルが追加されるワークシートが判断されます。テーブルが追加できない場合 (たとえば、アドレスが無効な場合や、テーブルが別のテーブルと重複している場合) は、エラーがスローされます。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getCount()](#getcount)|int|コレクション内のテーブルの数を取得します。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getItem(key: number または string)](#getitemkey-number-or-string)|[Table](table.md)|名前または ID を使用してテーブルを取得します。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemAt(index: number)](#getitematindex-number)|[Table](table.md)|コレクション内の位置に基づいてテーブルを取得します。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemOrNullObject(key: number または string)](#getitemornullobjectkey-number-or-string)|[Table](table.md)|名前または ID でテーブルを取得します。テーブルが存在しない場合は null オブジェクトを返します。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>メソッドの詳細


### <a name="addaddress-object-hasheaders-bool"></a>add(address: [object, hasHeaders: bool)
新しいテーブルを作成します。範囲オブジェクトまたはソース アドレスにより、テーブルが追加されるワークシートが判断されます。テーブルが追加できない場合 (たとえば、アドレスが無効な場合や、テーブルが別のテーブルと重複している場合) は、エラーがスローされます。

#### <a name="syntax"></a>構文
```js
tableCollectionObject.add(address, hasHeaders);
```

#### <a name="parameters"></a>パラメーター
| パラメーター       | 型    |説明|
|:---------------|:--------|:----------|:---|
|address|[object|Range オブジェクト、あるいはデータ ソースを表す範囲の文字列アドレスまたは名前。アドレスにシート名が含まれていない場合は、現在作業中のシートが使用されます。1.1 では、文字列パラメーターを使用します。1.3 では Range オブジェクトも受け入れることができます。|
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

### <a name="getcount"></a>getCount()
コレクション内のテーブルの数を取得します。

#### <a name="syntax"></a>構文
```js
tableCollectionObject.getCount();
```

#### <a name="parameters"></a>パラメーター
なし

#### <a name="returns"></a>戻り値
int

### <a name="getitemkey-number-or-string"></a>getItem(key: number または string)
名前または ID でテーブルを取得します。

#### <a name="syntax"></a>構文
```js
tableCollectionObject.getItem(key);
```

#### <a name="parameters"></a>パラメーター
| パラメーター       | 型    |説明|
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
| パラメーター       | 型    |説明|
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


### <a name="getitemornullobjectkey-number-or-string"></a>getItemOrNullObject(key: number または string)
名前または ID でテーブルを取得します。テーブルが存在しない場合は null オブジェクトを返します。

#### <a name="syntax"></a>構文
```js
tableCollectionObject.getItemOrNullObject(key);
```

#### <a name="parameters"></a>パラメーター
| パラメーター       | 型    |説明|
|:---------------|:--------|:----------|:---|
|Key|number または string|取得するテーブルの名前または ID。|

#### <a name="returns"></a>戻り値
[Table](table.md)
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