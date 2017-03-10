# <a name="worksheetcollection-object-javascript-api-for-excel"></a>WorksheetCollection オブジェクト (JavaScript API for Excel)

ブックの一部であるワークシート オブジェクトのコレクションを表します。

## <a name="properties"></a>プロパティ

| プロパティ       | 型    |説明| 要件セット|
|:---------------|:--------|:----------|:----|
|items|[Worksheet[]](worksheet.md)|ワークシート オブジェクトのコレクション。値の取得のみ可能です。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_プロパティのアクセスの[例を参照してください。](#property-access-examples)_

## <a name="relationships"></a>関係
なし


## <a name="methods"></a>メソッド

| メソッド           | 戻り値の型    |説明| 要件セット|
|:---------------|:--------|:----------|:----|
|[add(name: string)](#addname-string)|[Worksheet](worksheet.md)|新しいワークシートをブックに追加します。ワークシートは、既存のワークシートの末尾に追加されます。新しく追加したワークシートをアクティブにする場合は、そのワークシートに対して ".activate() を呼び出します。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getActiveWorksheet()](#getactiveworksheet)|[Worksheet](worksheet.md)|ブックの、現在作業中のワークシートを取得します。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getCount(visibleOnly: bool)](#getcountvisibleonly-bool)|int|コレクション内のワークシートの数を取得します。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getItem(key: string)](#getitemkey-string)|[Worksheet](worksheet.md)|名前または ID を使用して、ワークシート オブジェクトを取得します。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemOrNullObject(key: string)](#getitemornullobjectkey-string)|[Worksheet](worksheet.md)|名前または ID を使用して、ワークシート オブジェクトを取得します。ワークシートが存在しない場合は null オブジェクトを返します。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>メソッドの詳細


### <a name="addname-string"></a>add(name: string)
新しいワークシートをブックに追加します。ワークシートは、既存のワークシートの末尾に追加されます。新しく追加したワークシートをアクティブにする場合は、そのワークシートに対して ".activate() を呼び出します。

#### <a name="syntax"></a>構文
```js
worksheetCollectionObject.add(name);
```

#### <a name="parameters"></a>パラメーター
| パラメーター       | 型    |説明|
|:---------------|:--------|:----------|:---|
|name|string|省略可能。追加するワークシートの名前。指定する場合、名前は一意である必要があります。指定されていない場合は、Excel が新しいワークシートの名前を決定します。|

#### <a name="returns"></a>戻り値
[Worksheet](worksheet.md)

#### <a name="examples"></a>例

```js
Excel.run(function (ctx) { 
    var wSheetName = 'Sample Name';
    var worksheet = ctx.workbook.worksheets.add(wSheetName);
    worksheet.load('name');
    return ctx.sync().then(function() {
        console.log(worksheet.name);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getactiveworksheet"></a>getActiveWorksheet()
ブックの、現在作業中のワークシートを取得します。

#### <a name="syntax"></a>構文
```js
worksheetCollectionObject.getActiveWorksheet();
```

#### <a name="parameters"></a>パラメーター
なし

#### <a name="returns"></a>戻り値
[Worksheet](worksheet.md)

#### <a name="examples"></a>例

```js
Excel.run(function (ctx) {  
    var activeWorksheet = ctx.workbook.worksheets.getActiveWorksheet();
    activeWorksheet.load('name');
    return ctx.sync().then(function() {
            console.log(activeWorksheet.name);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getcountvisibleonly-bool"></a>getCount(visibleOnly: bool)
コレクション内のワークシートの数を取得します。

#### <a name="syntax"></a>構文
```js
worksheetCollectionObject.getCount(visibleOnly);
```

#### <a name="parameters"></a>パラメーター
| パラメーター       | 型    |説明|
|:---------------|:--------|:----------|:---|
|visibleOnly|bool|省略可能。true に設定されている場合は、表示されているワークシートのみを返します。 |

#### <a name="returns"></a>戻り値
int

### <a name="getitemkey-string"></a>getItem(key: string)
名前または ID を使用して、ワークシート オブジェクトを取得します。

#### <a name="syntax"></a>構文
```js
worksheetCollectionObject.getItem(key);
```

#### <a name="parameters"></a>パラメーター
| パラメーター       | 型    |説明|
|:---------------|:--------|:----------|:---|
|Key|string|ワークシートの名前または ID。|

#### <a name="returns"></a>戻り値
[Worksheet](worksheet.md)

### <a name="getitemornullobjectkey-string"></a>getItemOrNullObject(key: string)
名前または ID を使用して、ワークシート オブジェクトを取得します。ワークシートが存在しない場合は null オブジェクトを返します。

#### <a name="syntax"></a>構文
```js
worksheetCollectionObject.getItemOrNullObject(key);
```

#### <a name="parameters"></a>パラメーター
| パラメーター       | 型    |説明|
|:---------------|:--------|:----------|:---|
|Key|string|ワークシートの名前または ID。|

#### <a name="returns"></a>戻り値
[Worksheet](worksheet.md)
### <a name="property-access-examples"></a>プロパティのアクセスの例
```js
Excel.run(function (ctx) { 
    var worksheets = ctx.workbook.worksheets;
    worksheets.load('items');
    return ctx.sync().then(function() {
        for (var i = 0; i < worksheets.items.length; i++)
        {
            console.log(worksheets.items[i].name);
            console.log(worksheets.items[i].index);
        }
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
