# <a name="worksheet-object-javascript-api-for-excel"></a>Worksheet オブジェクト (JavaScript API for Excel)

Excel のワークシートは、セルのグリッドです。データ、表、グラフなどを含めることができます。

## <a name="properties"></a>プロパティ

| プロパティ       | 型    |説明| 要件セット|
|:---------------|:--------|:----------|:----|
|id|string|指定されたブックのワークシートを一意に識別する値を返します。この識別子の値は、ワークシートの名前を変更したり移動したりしても同じままです。値の取得のみ可能です。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|name|string|ワークシートの表示名。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|position|int|0 を起点とした、ブック内のワークシートの位置。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|visibility|string|ワークシートの可視性。使用可能な値は次のとおりです。Visible、Hidden、VeryHidden。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_プロパティのアクセスの[例を参照してください。](#property-access-examples)_

## <a name="relationships"></a>関係
| リレーションシップ | 型    |説明| 要件セット|
|:---------------|:--------|:----------|:----|
|charts|[ChartCollection](chartcollection.md)|ワークシートの一部になっているグラフのコレクションを返します。値の取得のみ可能です。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|names|[NamedItemCollection](nameditemcollection.md)|現在のワークシートにスコープされている名前のコレクション。読み取り専用です。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|pivotTables|[PivotTableCollection](pivottablecollection.md)|ワークシートの一部になっているピボットテーブルのコレクション。読み取り専用。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|protection|[WorksheetProtection](worksheetprotection.md)|ワークシートのシート保護オブジェクトを返します。読み取り専用です。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|tables|[TableCollection](tablecollection.md)|ワークシートの一部になっているグラフのコレクション。値の取得のみ可能です。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>メソッド

| メソッド           | 戻り値の型    |説明| 要件セット|
|:---------------|:--------|:----------|:----|
|[activate()](#activate)|void|Excel UI でワークシートをアクティブにします。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[delete()](#delete)|void|ブックからワークシートを削除します。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getCell(row: number, column: number)](#getcellrow-number-column-number)|[Range](range.md)|行と列の番号に基づいて、1 つのセルを含んだ範囲オブジェクトを取得します。このセルは、ワークシートのグリッド内であれば、親の範囲の境界の外のセルであってもかまいません。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getRange(address: string)](#getrangeaddress-string)|[Range](range.md)|アドレスまたは名前で指定された範囲オブジェクトを取得します。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getUsedRange(valuesOnly: [ApiSet(Version)](#getusedrangevaluesonly-apisetversion)|[Range](range.md)|使用範囲とは、値または書式設定が割り当たっているすべてのセルを包含する最小の範囲です。ワークシート全体が空白の場合、この関数は左上のセルを返します (つまり、エラーは*発生しません*)。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getUsedRangeOrNullObject(valuesOnly: bool)](#getusedrangeornullobjectvaluesonly-bool)|[Range](range.md)|使用範囲とは、値または書式設定が割り当たっているすべてのセルを包含する最小の範囲です。ワークシート全体が空白の場合、この関数は null オブジェクトを返します。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>メソッドの詳細


### <a name="activate"></a>activate()
Excel UI でワークシートをアクティブにします。

#### <a name="syntax"></a>構文
```js
worksheetObject.activate();
```

#### <a name="parameters"></a>パラメーター
なし

#### <a name="returns"></a>戻り値
void

#### <a name="examples"></a>例

```js
Excel.run(function (ctx) { 
    var wSheetName = 'Sheet1';
    var worksheet = ctx.workbook.worksheets.getItem(wSheetName);
    worksheet.activate();
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="delete"></a>delete()
ブックからワークシートを削除します。

#### <a name="syntax"></a>構文
```js
worksheetObject.delete();
```

#### <a name="parameters"></a>パラメーター
なし

#### <a name="returns"></a>戻り値
void

#### <a name="examples"></a>例

```js
Excel.run(function (ctx) { 
    var wSheetName = 'Sheet1';
    var worksheet = ctx.workbook.worksheets.getItem(wSheetName);
    worksheet.delete();
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getcellrow-number-column-number"></a>getCell(row: number, column: number)
行と列の番号に基づいて、1 つのセルを含んだ範囲オブジェクトを取得します。このセルは、ワークシートのグリッド内であれば、親の範囲の境界の外のセルであってもかまいません。

#### <a name="syntax"></a>構文
```js
worksheetObject.getCell(row, column);
```

#### <a name="parameters"></a>パラメーター
| パラメーター       | 型    |説明|
|:---------------|:--------|:----------|:---|
|row|number|取得するセルの行番号。0 を起点とする番号になります。|
|column|number|取得するセルの列番号。0 を起点とする番号になります。|

#### <a name="returns"></a>戻り値
[Range](range.md)

#### <a name="examples"></a>例

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:F8";
    var worksheet = ctx.workbook.worksheets.getItem(sheetName);
    var cell = worksheet.getCell(0,0);
    cell.load('address');
    return ctx.sync().then(function() {
        console.log(cell.address);
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getrangeaddress-string"></a>getRange(address: string)
アドレスまたは名前で指定された範囲 オブジェクトを取得します。

#### <a name="syntax"></a>構文
```js
worksheetObject.getRange(address);
```

#### <a name="parameters"></a>パラメーター
| パラメーター       | 型    |説明|
|:---------------|:--------|:----------|:---|
|address|string|省略可能。範囲のアドレスまたは名前。指定されていない場合は、ワークシート全体の範囲が返されます。|

#### <a name="returns"></a>戻り値
[Range](range.md)

#### <a name="examples"></a>例
次の例では、範囲アドレスを使用して、範囲オブジェクトを取得しています。

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:F8";
    var worksheet = ctx.workbook.worksheets.getItem(sheetName);
    var range = worksheet.getRange(rangeAddress);
    range.load('cellCount');
    return ctx.sync().then(function() {
        console.log(range.cellCount);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

次の例では、名前付き範囲を使用して、範囲オブジェクトを取得しています。

```js

Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeName = 'MyRange';
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeName);
    range.load('address');
    return ctx.sync().then(function() {
        console.log(range.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### <a name="getusedrangevaluesonly-apisetversion"></a>getUsedRange(valuesOnly: [ApiSet(Version)
使用範囲とは、値または書式設定が割り当たっているすべてのセルを包含する最小の範囲です。ワークシート全体が空白の場合、この関数は左上のセルを返します (つまり、エラーは*発生しません*)。

#### <a name="syntax"></a>構文
```js
worksheetObject.getUsedRange(valuesOnly);
```

#### <a name="parameters"></a>パラメーター
| パラメーター       | 型    |説明|
|:---------------|:--------|:----------|:---|
|valuesOnly|[ApiSet(Version|値の入っているセルのみを使用セルと見なします (書式設定は無視されます)。|

#### <a name="returns"></a>戻り値
[Range](range.md)

#### <a name="examples"></a>例

```js
Excel.run(function (ctx) { 
    var wSheetName = 'Sheet1';
    var worksheet = ctx.workbook.worksheets.getItem(wSheetName);
    var usedRange = worksheet.getUsedRange();
    usedRange.load('address');
    return ctx.sync().then(function() {
            console.log(usedRange.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getusedrangeornullobjectvaluesonly-bool"></a>getUsedRangeOrNullObject(valuesOnly: bool)
使用範囲とは、値または書式設定が割り当たっているすべてのセルを包含する最小の範囲です。ワークシート全体が空白の場合、この関数は null オブジェクトを返します。

#### <a name="syntax"></a>構文
```js
worksheetObject.getUsedRangeOrNullObject(valuesOnly);
```

#### <a name="parameters"></a>パラメーター
| パラメーター       | 型    |説明|
|:---------------|:--------|:----------|:---|
|valuesOnly|bool|省略可能。値の入っているセルのみを使用セルと見なします。|

#### <a name="returns"></a>戻り値
[Range](range.md)
### <a name="property-access-examples"></a>プロパティのアクセスの例

シート名に基づいて、ワークシートのプロパティを取得します。

```js
Excel.run(function (ctx) { 
    var wSheetName = 'Sheet1';
    var worksheet = ctx.workbook.worksheets.getItem(wSheetName);
    worksheet.load('position')
    return ctx.sync().then(function() {
            console.log(worksheet.position);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

ワークシートの位置を設定します。 

```js
Excel.run(function (ctx) { 
    var wSheetName = 'Sheet1';
    var worksheet = ctx.workbook.worksheets.getItem(wSheetName);
    worksheet.position = 2;
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
