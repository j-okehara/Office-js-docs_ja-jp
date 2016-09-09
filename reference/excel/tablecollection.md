# TableCollection オブジェクト (JavaScript API for Excel)

ブックの一部として含まれている、すべての表のコレクションを表します。

## プロパティ

| プロパティ     | 型   |説明
|:---------------|:--------|:----------|
|count|int|ブックに含まれるテーブルの数を返します。読み取り専用です。|
|Items|[Table[]](table.md)|Table オブジェクトのコレクション。読み取り専用です。|

_プロパティのアクセスの[例](#例)をご覧ください。_

## リレーションシップ
なし


## メソッド

| メソッド           | 戻り値の型    |説明|
|:---------------|:--------|:----------|
|[add(address: string, hasHeaders: bool)](#addaddress-string-hasheaders-bool)|[テーブル](table.md)|新しい表を作成します。範囲のソース アドレスにより、表が追加されるワークシートが決まります。表が追加できない場合 (たとえば、アドレスが無効な場合や、表が別の表と重なっている場合) は、エラーがスローされます。|
|[getItem(key: number またはstring)](#getitemkey-number-またはstring)|[テーブル](table.md)|名前または ID でテーブルを取得します。|
|[getItemAt(index: number)](#getitematindex-number)|[テーブル](table.md)|コレクション内の位置に基づいてテーブルを取得します。|
|[load(param: object)](#loadparam-object)|void|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|

## メソッドの詳細


### add(address: string, hasHeaders: bool)
新しい表を作成します。範囲のソース アドレスにより、表が追加されるワークシートが決まります。表が追加できない場合 (たとえば、アドレスが無効な場合や、表が別の表と重なっている場合) は、エラーがスローされます。

#### 構文
```js
tableCollectionObject.add(address, hasHeaders);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|address|string|データ ソースを表す Range オブジェクトのアドレスまたは名前。アドレスにシート名が含まれていない場合は、現在作業中のシートが使用されます。|
|hasHeaders|bool|インポートされたデータに列ラベルがあるかどうかを示すブール値。ソースにヘッダーが含まれていない場合 (このプロパティが false に設定されている場合)、Excel はデータを下方向に 1 行シフトして、自動的にヘッダーを生成します。|

#### 戻り値
[テーブル](table.md)

#### 例

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

### getItem(key: number またはstring)
名前または ID でテーブルを取得します。

#### 構文
```js
tableCollectionObject.getItem(key);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|Key|number または string|取得するテーブルの名前または ID。|

#### 戻り値
[テーブル](table.md)

#### 例

```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
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


#### 例

```js
Excel.run(function (ctx) { 
    var table = ctx.workbook.tables.getItemAt(0);
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


### getItemAt(index: number)
コレクション内の位置に基づいてテーブルを取得します。

#### 構文
```js
tableCollectionObject.getItemAt(index);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|index|number|取得するオブジェクトのインデックス値。0 を起点とする番号になります。|

#### 戻り値
[テーブル](table.md)

#### 例

```js
Excel.run(function (ctx) { 
    var table = ctx.workbook.tables.getItemAt(0);
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
    var tables = ctx.workbook.tables;
    tables.load('items');
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

表の数を取得します

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
