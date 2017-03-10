# <a name="range-object-javascript-api-for-excel"></a>Range オブジェクト (JavaScript API for Excel)

範囲は、1 つ以上の隣接するセル (セル、行、列、セルのブロックなど) のセットを表します。

## <a name="properties"></a>プロパティ

| プロパティ       | 型    |説明| 要件セット|
|:---------------|:--------|:----------|:----|
|アドレス|string|A1 スタイルの範囲参照を表します。アドレス値には、シート参照が格納されます (例: Sheet1!A1:B4)。読み取り専用です。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|addressLocal|string|ユーザーの言語で指定された範囲の範囲参照を表します。読み取り専用です。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|cellCount|int|範囲に含まれるセルの数。セルの数が 2^31-1 (2,147,483,647) を超えると、この API は -1 を返します。読み取り専用です。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|columnCount|int|範囲に含まれる列の合計数を表します。読み取り専用です。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|columnHidden|bool|現在の範囲のすべての列が非表示になっているかどうかを表します。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|columnIndex|int|範囲に含まれる最初のセルの列番号を表します。0 を起点とする番号になります。読み取り専用です。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|formulas|object[][]|A1 スタイル表記の数式を表します。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|formulasLocal|object[][]|ユーザーの言語と数値書式ロケールで、A1 スタイル表記の数式を表します。たとえば、英語の数式 "=SUM(A1, 1.5)" は、ドイツ語では "=SUMME(A1; 1,5)" になります。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|formulasR1C1|object[][]|R1C1 スタイル表記の数式を表します。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|hidden|bool|現在の範囲のすべてのセルが非表示になっているかどうかを表します。読み取り専用です。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|numberFormat|オブジェクト型 (Object)|指定したセルの Excel の数値書式コードを表します。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|rowCount|int|範囲に含まれる行の合計数を返します。読み取り専用です。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|rowHidden|bool|現在の範囲のすべての行が非表示になっているかどうかを表します。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|rowIndex|int|範囲に含まれる最初のセルの行番号を返します。0 を起点とする番号になります。読み取り専用です。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|text|object[][]|指定した範囲のテキスト値。テキスト値は、セルの幅には依存しません。Excel UI で発生する # 記号による置換は、この API から返されるテキスト値には影響しません。読み取り専用です。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|valueTypes|string|各セルのデータの種類を表します。読み取り専用です。使用可能な値は次のとおりです。Unknown、Empty、String、Integer、Double、Boolean、Error。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|values|object[][]|指定した範囲の Raw 値を表します。返されるデータの型は、文字列、数値、またはブール値のいずれかになります。エラーが含まれているセルは、エラー文字列を返します。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_プロパティのアクセスの[例を参照してください。](#property-access-examples)_

## <a name="relationships"></a>関係
| リレーションシップ | 型    |説明| 要件セット|
|:---------------|:--------|:----------|:----|
|format|[RangeFormat](rangeformat.md)|Format オブジェクト (範囲のフォント、塗りつぶし、罫線、配置などのプロパティをカプセル化するオブジェクト) を返します。読み取り専用です。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|sort|[RangeSort](rangesort.md)|現在の範囲について、範囲の並べ替えを表します。読み取り専用。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|worksheet|[Worksheet](worksheet.md)|現在の範囲を含んでいるワークシート。読み取り専用です。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>メソッド

| メソッド           | 戻り値の型    |説明| 要件セット|
|:---------------|:--------|:----------|:----|
|[clear(applyTo: string)](#clearapplyto-string)|(非推奨)|範囲の値、書式、塗りつぶし、罫線などをクリアします。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[delete(shift: string)](#deleteshift-string)|(非推奨)|範囲に関連付けられているセルを削除します。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getBoundingRect(anotherRange:Range or string)](#getboundingrectanotherrange-range-or-string)|[Range](range.md)|指定した範囲を包含する、最小の Range オブジェクトを取得します。たとえば、"B2:C5" と "D10:E15" の GetBoundingRect は、"B2:E16" になります。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getCell(row: number, column: number)](#getcellrow-number-column-number)|[Range](range.md)|行と列の番号に基づいて、1 つのセルを含んだ範囲オブジェクトを取得します。以外このセルは、ワークシートのグリッド内であれば、親の範囲の境界の外のセルであってもかまいません。返されるセルは、範囲の左上のセルを基準に配置されます。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getColumn(column: number)](#getcolumncolumn-number)|[Range](range.md)|範囲に含まれる列を 1 つ取得します。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getColumnsAfter(count: number)](#getcolumnsaftercount-number)|[Range](range.md)|現在の Range オブジェクトの右にある特定の列数を取得します。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getColumnsBefore(count: number)](#getcolumnsbeforecount-number)|[Range](range.md)|現在の Range オブジェクトの左にある特定の列数を取得します。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getEntireColumn()](#getentirecolumn)|[Range](range.md)|範囲の列全体を表すオブジェクトを取得します (たとえば、現在の範囲がセル "B4:E11" を表す場合、その `getEntireColumn` は列 "B:E" を表す範囲になります)。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getEntireRow()](#getentirerow)|[Range](range.md)|範囲の行全体を表すオブジェクトを取得します (たとえば、現在の範囲がセル "B4:E11" を表す場合、その `GetEntireRow` は行 "4:11" を表す範囲になります)。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getIntersection(anotherRange:Range or string)](#getintersectionanotherrange-range-or-string)|[Range](range.md)|指定した範囲の長方形の交差部分を表す範囲オブジェクトを取得します。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getIntersectionOrNullObject(anotherRange:Range または string)](#getintersectionornullobjectanotherrange-range-or-string)|[Range](range.md)|指定した範囲の長方形の交差部分を表す範囲オブジェクトを取得します。交差部分が見つからない場合は、null オブジェクトを返します。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getLastCell()](#getlastcell)|[Range](range.md)|範囲内の最後のセルを取得します。たとえば、"B2:D5" の最後のセルは "D5" になります。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getLastColumn()](#getlastcolumn)|[Range](range.md)|範囲内の最後の列を取得します。たとえば、"B2:D5" の最後の列は "D2:D5" になります。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getLastRow()](#getlastrow)|[Range](range.md)|範囲内の最後の行を取得します。たとえば、"B2:D5" の最後の行は "B5:D5" になります。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getOffsetRange(rowOffset: number, columnOffset: number)](#getoffsetrangerowoffset-number-columnoffset-number)|[Range](range.md)|指定した範囲からのオフセットで範囲を表すオブジェクトを取得します。返される範囲のディメンションは、この範囲と一致します。結果の範囲がワークシートのグリッドの境界線の外にはみ出る場合は、エラーがスローされます。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getResizedRange(deltaRows: number, deltaColumns: number)](#getresizedrangedeltarows-number-deltacolumns-number)|[Range](range.md)|現在の Range オブジェクトに似た (ただし、右下隅がいくつかの行と列で拡張 (または縮小) されている) Range オブジェクトを取得します。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getRow(row: number)](#getrowrow-number)|[Range](range.md)|範囲に含まれている行を 1 つ取得します。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getRowsAbove(count: number)](#getrowsabovecount-number)|[Range](range.md)|現在の Range オブジェクトの上にある特定の行数を取得します。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getRowsBelow(count: number)](#getrowsbelowcount-number)|[Range](range.md)|現在の Range オブジェクトの下にある特定の行数を取得します。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getUsedRange(valuesOnly: [ApiSet(Version)](#getusedrangevaluesonly-apisetversion)|[Range](range.md)|指定した範囲オブジェクトのうち使用されている範囲を返します。範囲内に使用済みのセルがない場合、この関数は ItemNotFound エラーをスローします。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getUsedRangeOrNullObject(valuesOnly: bool)](#getusedrangeornullobjectvaluesonly-bool)|[Range](range.md)|指定した範囲オブジェクトのうち使用されている範囲を返します。範囲内に使用済みのセルがない場合、この関数は null オブジェクトを返します。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getVisibleView()](#getvisibleview)|[RangeView](rangeview.md)|現在の範囲の表示されている行を表します。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|[insert(shift: string)](#insertshift-string)|[Range](range.md)|この範囲を占めるセルまたはセルの範囲をワークシートに挿入し、領域を空けるために他のセルをシフトします。この時点で空き領域に位置する、新しい Range オブジェクトが返されます。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[merge(across: bool)](#mergeacross-bool)|void|範囲内のセルをワークシートの 1 つの領域に結合します。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[select()](#select)|void|Excel UI で指定した範囲を選択します。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[unmerge()](#unmerge)|void|範囲内のセルを結合解除して別々のセルにします。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>メソッドの詳細


### <a name="clearapplyto-string"></a>clear(applyTo: string)
範囲の値、書式、塗りつぶし、罫線などをクリアします。

#### <a name="syntax"></a>構文
```js
rangeObject.clear(applyTo);
```

#### <a name="parameters"></a>パラメーター
| パラメーター       | 型    |説明|
|:---------------|:--------|:----------|:---|
|applyTo|文字列|省略可能。クリア操作の種類を決定します。使用可能な値は次のとおりです。`All`、Default-option、`Formats`、`Contents` |

#### <a name="returns"></a>戻り値
void

#### <a name="examples"></a>例

次の例では、範囲の書式と内容をクリアします。 

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "D:F";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.clear();
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="deleteshift-string"></a>delete(shift: string)
範囲に関連付けられているセルを削除します。

#### <a name="syntax"></a>構文
```js
rangeObject.delete(shift);
```

#### <a name="parameters"></a>パラメーター
| パラメーター       | 型    |説明|
|:---------------|:--------|:----------|:---|
|<legacyBold>Shift</legacyBold>|文字列|セルをシフトする方向を指定します。使用可能な値は次のとおりです。Up、Left|

#### <a name="returns"></a>戻り値
void

#### <a name="examples"></a>例

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "D:F";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.delete();
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getboundingrectanotherrange-range-or-string"></a>getBoundingRect(anotherRange:Range or string)
指定した範囲を包含する、最小の Range オブジェクトを取得します。たとえば、"B2:C5" と "D10:E15" の GetBoundingRect は、"B2:E16" になります。

#### <a name="syntax"></a>構文
```js
rangeObject.getBoundingRect(anotherRange);
```

#### <a name="parameters"></a>パラメーター
| パラメーター       | 型    |説明|
|:---------------|:--------|:----------|:---|
|anotherRange|Range または string|Range オブジェクト、アドレスまたは範囲名。|

#### <a name="returns"></a>戻り値
[Range](range.md)

#### <a name="examples"></a>例

```js

Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "D4:G6";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var range = range.getBoundingRect("G4:H8");
    range.load('address');
    return ctx.sync().then(function() {
        console.log(range.address); // Prints Sheet1!D4:H8
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getcellrow-number-column-number"></a>getCell(row: number, column: number)
行と列の番号に基づいて、1 つのセルを含んだ範囲オブジェクトを取得します。以外このセルは、ワークシートのグリッド内であれば、親の範囲の境界の外のセルであってもかまいません。返されるセルは、範囲の左上のセルを基準に配置されます。

#### <a name="syntax"></a>構文
```js
rangeObject.getCell(row, column);
```

#### <a name="parameters"></a>パラメーター
| パラメーター       | 型    |説明|
|:---------------|:--------|:----------|:---|
|row|number|取得するセルの行番号。0 を起点とする番号になります。|
|column|number|取得セルの列番号。0 を起点とする番号になります。|

#### <a name="returns"></a>戻り値
[Range](range.md)

#### <a name="examples"></a>例

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:F8";
    var worksheet = ctx.workbook.worksheets.getItem(sheetName);
    var range = worksheet.getRange(rangeAddress);
    var cell = range.cell(0,0);
    cell.load('address');
    return ctx.sync().then(function() {
        console.log(cell.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getcolumncolumn-number"></a>getColumn(column: number)
範囲に含まれる列を 1 つ取得します。

#### <a name="syntax"></a>構文
```js
rangeObject.getColumn(column);
```

#### <a name="parameters"></a>パラメーター
| パラメーター       | 型    |説明|
|:---------------|:--------|:----------|:---|
|column|number|取得する範囲の列番号。0 を起点とする番号になります。|

#### <a name="returns"></a>戻り値
[Range](range.md)

#### <a name="examples"></a>例

```js

Excel.run(function (ctx) { 
    var sheetName = "Sheet19";
    var rangeAddress = "A1:F8";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress).getColumn(1);
    range.load('address');
    return ctx.sync().then(function() {
        console.log(range.address); // prints Sheet1!B1:B8
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getcolumnsaftercount-number"></a>getColumnsAfter(count: number)
現在の Range オブジェクトの右にある特定の列数を取得します。

#### <a name="syntax"></a>構文
```js
rangeObject.getColumnsAfter(count);
```

#### <a name="parameters"></a>パラメーター
| パラメーター       | 型    |説明|
|:---------------|:--------|:----------|:---|
|count|number|省略可能。結果の範囲に含める列の数です。通常、正の数値を使用して現在の範囲外に範囲を作成します。負の数値を使用して、現在の範囲内に範囲を作成することもできます。既定値は 1 です。|

#### <a name="returns"></a>戻り値
[Range](range.md)

### <a name="getcolumnsbeforecount-number"></a>getColumnsBefore(count: number)
現在の Range オブジェクトの左にある特定の列数を取得します。

#### <a name="syntax"></a>構文
```js
rangeObject.getColumnsBefore(count);
```

#### <a name="parameters"></a>パラメーター
| パラメーター       | 型    |説明|
|:---------------|:--------|:----------|:---|
|count|number|省略可能。結果の範囲に含める列の数です。通常、正の数値を使用して現在の範囲外に範囲を作成します。負の数値を使用して、現在の範囲内に範囲を作成することもできます。既定値は 1 です。|

#### <a name="returns"></a>戻り値
[Range](range.md)

### <a name="getentirecolumn"></a>getEntireColumn()
範囲の列全体を表すオブジェクトを取得します (たとえば、現在の範囲がセル "B4:E11" を表す場合、その `getEntireColumn` は列 "B:E" を表す範囲になります)。

#### <a name="syntax"></a>構文
```js
rangeObject.getEntireColumn();
```

#### <a name="parameters"></a>パラメーター
なし

#### <a name="returns"></a>戻り値
[Range](range.md)

#### <a name="examples"></a>例

注: Range のグリッド プロパティ (values、numberFormat、formulas) には、当該の範囲に境界がないため、`null` が格納されます。

```js

Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "D:F";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var rangeEC = range.getEntireColumn();
    rangeEC.load('address');
    return ctx.sync().then(function() {
        console.log(rangeEC.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### <a name="getentirerow"></a>getEntireRow()
範囲の行全体を表すオブジェクトを取得します (たとえば、現在の範囲がセル "B4:E11" を表す場合、その `GetEntireRow` は行 "4:11" を表す範囲になります)。

#### <a name="syntax"></a>構文
```js
rangeObject.getEntireRow();
```

#### <a name="parameters"></a>パラメーター
なし

#### <a name="returns"></a>戻り値
[Range](range.md)

#### <a name="examples"></a>例
```js

Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "D:F"; 
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var rangeER = range.getEntireRow();
    rangeER.load('address');
    return ctx.sync().then(function() {
        console.log(rangeER.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
Range のグリッド プロパティ (values、numberFormat、formulas) には、当該の範囲に制限がないため、`null` が格納されます。


### <a name="getintersectionanotherrange-range-or-string"></a>getIntersection(anotherRange:Range or string)
指定した範囲の長方形の交差を表す Range オブジェクトを取得します。

#### <a name="syntax"></a>構文
```js
rangeObject.getIntersection(anotherRange);
```

#### <a name="parameters"></a>パラメーター
| パラメーター       | 型    |説明|
|:---------------|:--------|:----------|:---|
|anotherRange|Range または string|範囲の交差を判断するために使用される、Range オブジェクトまたは Range アドレス。|

#### <a name="returns"></a>戻り値
[Range](range.md)

#### <a name="examples"></a>例

```js

Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:F8";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress).getIntersection("D4:G6");
    range.load('address');
    return ctx.sync().then(function() {
        console.log(range.address); // prints Sheet1!D4:F6
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getintersectionornullobjectanotherrange-range-or-string"></a>getIntersectionOrNullObject(anotherRange:Range または string)
指定した範囲の長方形の交差部分を表す範囲オブジェクトを取得します。交差部分が見つからない場合は、null オブジェクトを返します。

#### <a name="syntax"></a>構文
```js
rangeObject.getIntersectionOrNullObject(anotherRange);
```

#### <a name="parameters"></a>パラメーター
| パラメーター       | 型    |説明|
|:---------------|:--------|:----------|:---|
|anotherRange|Range または string|範囲の交差を判断するために使用される、Range オブジェクトまたは Range アドレス。|

#### <a name="returns"></a>戻り値
[Range](range.md)

### <a name="getlastcell"></a>getLastCell()
範囲内の最後のセルを取得します。たとえば、"B2:D5" の最後のセルは "D5" になります。

#### <a name="syntax"></a>構文
```js
rangeObject.getLastCell();
```

#### <a name="parameters"></a>パラメーター
なし

#### <a name="returns"></a>戻り値
[Range](range.md)

#### <a name="examples"></a>例

```js

Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:F8";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress).getLastCell();
    range.load('address');
    return ctx.sync().then(function() {
        console.log(range.address); // prints Sheet1!F8
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getlastcolumn"></a>getLastColumn()
範囲内の最後の列を取得します。たとえば、"B2:D5" の最後の列は "D2:D5" になります。

#### <a name="syntax"></a>構文
```js
rangeObject.getLastColumn();
```

#### <a name="parameters"></a>パラメーター
なし

#### <a name="returns"></a>戻り値
[Range](range.md)

#### <a name="examples"></a>例

```js

Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:F8";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress).getLastColumn();
    range.load('address');
    return ctx.sync().then(function() {
        console.log(range.address); // prints Sheet1!F1:F8
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getlastrow"></a>getLastRow()
範囲内の最後の行を取得します。たとえば、"B2:D5" の最後の行は "B5:D5" になります。

#### <a name="syntax"></a>構文
```js
rangeObject.getLastRow();
```

#### <a name="parameters"></a>パラメーター
なし

#### <a name="returns"></a>戻り値
[Range](range.md)

#### <a name="examples"></a>例

```js

Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:F8";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress).getLastRow();
    range.load('address');
    return ctx.sync().then(function() {
        console.log(range.address); // prints Sheet1!A8:F8
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```



### <a name="getoffsetrangerowoffset-number-columnoffset-number"></a>getOffsetRange(rowOffset: number, columnOffset: number)
指定した範囲からのオフセットで範囲を表すオブジェクトを取得します。返される範囲のディメンションは、この範囲と一致します。結果の範囲がワークシートのグリッドの境界線の外にはみ出る場合は、エラーがスローされます。

#### <a name="syntax"></a>構文
```js
rangeObject.getOffsetRange(rowOffset, columnOffset);
```

#### <a name="parameters"></a>パラメーター
| パラメーター       | 型    |説明|
|:---------------|:--------|:----------|:---|
|rowOffset|number|範囲をオフセットする行数 (正、負、または 0)。正の値は下方向へのオフセットです。また、負の値は上方向へのオフセットです。|
|columnOffset|number|範囲をオフセットする列数 (正、負、または 0)。正の値は右方向へのオフセットです。また、負の値は左方向へのオフセットです。|

#### <a name="returns"></a>戻り値
[Range](range.md)

#### <a name="examples"></a>例

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "D4:F6";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress).getOffsetRange(-1,4);
    range.load('address');
    return ctx.sync().then(function() {
        console.log(range.address); // prints Sheet1!H3:K5
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getresizedrangedeltarows-number-deltacolumns-number"></a>getResizedRange(deltaRows: number, deltaColumns: number)
現在の Range オブジェクトに似た (ただし、右下隅がいくつかの行と列で拡張 (または縮小) されている) Range オブジェクトを取得します。

#### <a name="syntax"></a>構文
```js
rangeObject.getResizedRange(deltaRows, deltaColumns);
```

#### <a name="parameters"></a>パラメーター
| パラメーター       | 型    |説明|
|:---------------|:--------|:----------|:---|
|deltaRows|number|現在の範囲を基準にして、右下隅を拡張する行の数です。範囲を拡張するには正の数値、または範囲を縮小するには負の数値を使用します。|
|deltaColumns|number|現在の範囲を基準にして、右下隅を拡張する列の数です。範囲を拡張するには正の数値、または範囲を縮小するには負の数値を使用します。|

#### <a name="returns"></a>戻り値
[Range](range.md)

### <a name="getrowrow-number"></a>getRow(row: number)
範囲に含まれている行を 1 つ取得します。

#### <a name="syntax"></a>構文
```js
rangeObject.getRow(row);
```

#### <a name="parameters"></a>パラメーター
| パラメーター       | 型    |説明|
|:---------------|:--------|:----------|:---|
|row|number|取得する範囲の行番号。0 を起点とする番号になります。|

#### <a name="returns"></a>戻り値
[Range](range.md)

#### <a name="examples"></a>例

```js

Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:F8";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress).getRow(1);
    range.load('address');
    return ctx.sync().then(function() {
        console.log(range.address); // prints Sheet1!A2:F2
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getrowsabovecount-number"></a>getRowsAbove(count: number)
現在の Range オブジェクトの上にある特定の行数を取得します。

#### <a name="syntax"></a>構文
```js
rangeObject.getRowsAbove(count);
```

#### <a name="parameters"></a>パラメーター
| パラメーター       | 型    |説明|
|:---------------|:--------|:----------|:---|
|count|number|省略可能。結果の範囲に含める行の数です。通常、正の数値を使用して現在の範囲外に範囲を作成します。負の数値を使用して、現在の範囲内に範囲を作成することもできます。既定値は 1 です。|

#### <a name="returns"></a>戻り値
[Range](range.md)

### <a name="getrowsbelowcount-number"></a>getRowsBelow(count: number)
現在の Range オブジェクトの下にある特定の行数を取得します。

#### <a name="syntax"></a>構文
```js
rangeObject.getRowsBelow(count);
```

#### <a name="parameters"></a>パラメーター
| パラメーター       | 型    |説明|
|:---------------|:--------|:----------|:---|
|count|number|省略可能。結果の範囲に含める行の数です。通常、正の数値を使用して現在の範囲外に範囲を作成します。負の数値を使用して、現在の範囲内に範囲を作成することもできます。既定値は 1 です。|

#### <a name="returns"></a>戻り値
[Range](range.md)

### <a name="getusedrangevaluesonly-apisetversion"></a>getUsedRange(valuesOnly: [ApiSet(Version)
指定した範囲オブジェクトのうち使用されている範囲を返します。範囲内に使用済みのセルがない場合、この関数は ItemNotFound エラーをスローします。

#### <a name="syntax"></a>構文
```js
rangeObject.getUsedRange(valuesOnly);
```

#### <a name="parameters"></a>パラメーター
| パラメーター       | 型    |説明|
|:---------------|:--------|:----------|:---|
|valuesOnly|[ApiSet(Version|値の入っているセルのみを使用セルと見なします。|

#### <a name="returns"></a>戻り値
[Range](range.md)

#### <a name="examples"></a>例

```js

Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "D:F";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var rangeUR = range.getUsedRange();
    rangeUR.load('address');
    return ctx.sync().then(function() {
        console.log(rangeUR.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getusedrangeornullobjectvaluesonly-bool"></a>getUsedRangeOrNullObject(valuesOnly: bool)
指定した範囲オブジェクトのうち使用されている範囲を返します。範囲内に使用済みのセルがない場合、この関数は null オブジェクトを返します。

#### <a name="syntax"></a>構文
```js
rangeObject.getUsedRangeOrNullObject(valuesOnly);
```

#### <a name="parameters"></a>パラメーター
| パラメーター       | 型    |説明|
|:---------------|:--------|:----------|:---|
|valuesOnly|bool|省略可能。値の入っているセルのみを使用セルと見なします。|

#### <a name="returns"></a>戻り値
[Range](range.md)

### <a name="getvisibleview"></a>getVisibleView()
現在の範囲の表示されている行を表します。

#### <a name="syntax"></a>構文
```js
rangeObject.getVisibleView();
```

#### <a name="parameters"></a>パラメーター
なし

#### <a name="returns"></a>戻り値
[RangeView](rangeview.md)

### <a name="insertshift-string"></a>insert(shift: string)
この範囲を占めるセルまたはセルの範囲をワークシートに挿入し、領域を空けるために他のセルをシフトします。この時点で空き領域に位置する、新しい Range オブジェクトが返されます。

#### <a name="syntax"></a>構文
```js
rangeObject.insert(shift);
```

#### <a name="parameters"></a>パラメーター
| パラメーター       | 型    |説明|
|:---------------|:--------|:----------|:---|
|<legacyBold>Shift</legacyBold>|文字列|セルをシフトする方向を指定します。使用可能な値は次のとおりです。Down、Right|

#### <a name="returns"></a>戻り値
[Range](range.md)

#### <a name="examples"></a>例

```js
    
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "F5:F10";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.insert();
    return ctx.sync(); 
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="mergeacross-bool"></a>merge(across: bool)
範囲内のセルをワークシートの 1 つの領域にマージします。

#### <a name="syntax"></a>構文
```js
rangeObject.merge(across);
```

#### <a name="parameters"></a>パラメーター
| パラメーター       | 型    |説明|
|:---------------|:--------|:----------|:---|
|across|bool|省略可能。指定した範囲のセルを行ごとに結合して、行ごとに別のセルを作成する場合は True に設定します。既定値は False です。|

#### <a name="returns"></a>戻り値
void

#### <a name="examples"></a>例
```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:C3";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.merge(true);
    return ctx.sync(); 
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
    var sheetName = "Sheet1";
    var rangeAddress = "A1:C3";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.unmerge();
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="select"></a>select()
Excel UI で指定した範囲を選択します。

#### <a name="syntax"></a>構文
```js
rangeObject.select();
```

#### <a name="parameters"></a>パラメーター
なし

#### <a name="returns"></a>戻り値
void

#### <a name="examples"></a>例

```js

Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "F5:F10"; 
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.select();
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="unmerge"></a>unmerge()
範囲内のセルを結合解除して別々のセルにします。

#### <a name="syntax"></a>構文
```js
rangeObject.unmerge();
```

#### <a name="parameters"></a>パラメーター
なし

#### <a name="returns"></a>戻り値
void

#### <a name="examples"></a>例
```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:C3";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.unmerge();
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### <a name="property-access-examples"></a>プロパティのアクセスの例

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
    var rangeName = 'MyRange';
    var range = ctx.workbook.names.getItem(rangeName).range;
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

次の例では、2 x 3 のグリッドを含んでいるグリッドに対して、数値の書式、値および数式を設定します。

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "F5:G7";
    var numberFormat = [[null, "d-mmm"], [null, "d-mmm"], [null, null]]
    var values = [["Today", 42147], ["Tomorrow", "5/24"], ["Difference in days", null]];
    var formulas = [[null,null], [null,null], [null,"=G6-G5"]];
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.numberFormat = numberFormat;
    range.values = values;
    range.formulas= formulas;
    range.load('text');
    return ctx.sync().then(function() {
        console.log(range.text);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
範囲を含んでいるワークシートを取得します。 

```js
/* This might be broken still - it was broken before because it 
    it was missing 'var', but might still be wrong because of
    getting information without loading properly. */
Excel.run(function (ctx) { 
    var names = ctx.workbook.names;
    var namedItem = names.getItem('MyRange');
    var range = namedItem.range;
    var rangeWorksheet = range.worksheet;
    rangeWorksheet.load('name');
    return ctx.sync().then(function() {
            console.log(rangeWorksheet.name);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

