# Range オブジェクト (JavaScript API for Excel)

_適用対象: Excel 2016、Excel Online、Excel for iOS、Office 2016_

範囲は、1 つ以上の隣接するセル (セル、行、列、セルのブロックなど) のセットを表します。

## プロパティ

| プロパティ     | 型   |説明
|:---------------|:--------|:----------|
|address|string|A1 スタイルの範囲参照を表します。アドレス値には、シート参照が格納されます (例: Sheet1!A1:B4)。読み取り専用です。|
|addressLocal|string|ユーザーの言語で指定された範囲の範囲参照を表します。読み取り専用です。|
|cellCount|int|範囲に含まれるセルの数。読み取り専用です。|
|columnCount|int|範囲に含まれる列の合計数を表します。読み取り専用です。|
|columnHidden|bool|現在の範囲のすべての列が非表示になっているかどうかを表します。|
|columnIndex|int|範囲に含まれる最初のセルの列番号を表します。0 を起点とする番号になります。読み取り専用です。|
|formulas|object[]|A1 スタイル表記の数式を表します。|
|formulasLocal|object[][]|ユーザーの言語と数値書式ロケールで、A1 スタイル表記の数式を表します。たとえば、英語の数式 "=SUM(A1, 1.5)" は、ドイツ語では "=SUMME(A1; 1,5)" になります。|
|formulasR1C1|object[][]|R1C1 スタイル表記の数式を表します。|
|Hidden|bool|現在の範囲のすべてのセルが非表示になっているかどうかを表します。読み取り専用です。|
|numberFormat|object[][]|指定したセルの数値書式コードを表します。|
|rowCount|int|範囲に含まれる行の合計数を返します。読み取り専用です。|
|rowHidden|bool|現在の範囲のすべての行が非表示になっているかどうかを表します。|
|rowIndex|int|範囲に含まれる最初のセルの行番号を返します。0 を起点とする番号になります。読み取り専用です。|
|text|object[][]|指定した範囲のテキスト値。テキスト値は、セルの幅には依存しません。Excel UI で発生する # 記号による置換は、この API から返されるテキスト値には影響しません。読み取り専用です。|
|valueTypes|string|各セルのデータの種類を表します。読み取り専用です。使用可能な値は次のとおりです。Unknown、Empty、String、Integer、Double、Boolean、Error。|
|values|object[][]|指定した範囲の Raw 値を表します。返されるデータの型は、文字列、数値、またはブール値のいずれかになります。エラーが含まれているセルは、エラー文字列を返します。|

_プロパティのアクセスの[例](#例)をご覧ください。_

## リレーションシップ
| リレーションシップ | 型   |説明|
|:---------------|:--------|:----------|
|オプション パラメーターを適用する|[RangeFormat](rangeformat.md)|Format オブジェクト (範囲のフォント、塗りつぶし、罫線、配置などのプロパティをカプセル化するオブジェクト) を返します。読み取り専用です。|
|sort|[RangeSort](rangesort.md)|範囲のソート順の構成を表します。読み取り専用です。|
|worksheet|[ワークシート](worksheet.md)|現在の範囲を含んでいるワークシート。読み取り専用です。|

## メソッド

| メソッド           | 戻り値の型    |説明|
|:---------------|:--------|:----------|
|[clear(applyTo: string)](#clearapplyto-string)|void|範囲の値、書式、塗りつぶし、罫線などをクリアします。|
|[delete(shift: string)](#deleteshift-string)|void|範囲に関連付けられているセルを削除します。|
|[getBoundingRect(anotherRange: Range or string)](#getboundingrectanotherrange-range-or-string)|[範囲](range.md)|指定した範囲を包含する、最小の Range オブジェクトを取得します。たとえば、"B2:C5" と "D10:E15" の getBoundingRect は、"B2:E15" になります。|
|[getCell(row: number, column: number)](#getcellrow-number-column-number)|[範囲](range.md)|行と列の番号に基づいて、1 つのセルを含んだ範囲オブジェクトを取得します。このセルは、ワークシートのグリッド内であれば、親の範囲の境界の外のセルであってもかまいません。返されるセルは、範囲の左上のセルを基準に配置されます。|
|[getColumn(column: number)](#getcolumncolumn-number)|[範囲](range.md)|範囲に含まれる列を 1 つ取得します。|
|[getEntireColumn()](#getentirecolumn)|[範囲](range.md)|範囲に含まれるすべての列を表すオブジェクトを取得します。|
|[getEntireRow()](#getentirerow)|[範囲](range.md)|範囲に含まれるすべての行を表すオブジェクトを取得します。|
|[getIntersection(anotherRange: Range or string)](#getintersectionanotherrange-range-or-string)|[範囲](range.md)|指定した範囲の長方形の交差を表す Range オブジェクトを取得します。|
|[getLastCell()](#getlastcell)|[範囲](range.md)|範囲内の最後のセルを取得します。たとえば、"B2:D5" の最後のセルは "D5" になります。|
|[getLastColumn()](#getlastcolumn)|[範囲](range.md)|範囲内の最後の列を取得します。たとえば、"B2:D5" の最後の列は "D2:D5" になります。|
|[getLastRow()](#getlastrow)|[範囲](range.md)|範囲内の最後の行を取得します。たとえば、"B2:D5" の最後の行は "B5:D5" になります。|
|[getOffsetRange(rowOffset: number, columnOffset: number)](#getoffsetrangerowoffset-number-columnoffset-number)|[範囲](range.md)|指定した範囲からのオフセットで範囲を表すオブジェクトを取得します。返される範囲のディメンションは、この範囲と一致します。結果の範囲が、ワークシートのグリッドの境界線の外にはみ出る場合は、例外がスローされます。|
|[getRow(row: number)](#getrowrow-number)|[範囲](range.md)|範囲に含まれている行を 1 つ取得します。|
|[getUsedRange(valuesOnly: bool)](#getusedrangevaluesonly-bool)|[範囲](range.md)|Range オブジェクトのうち使用されている部分範囲を返します。|
|[insert(shift: string)](#insertshift-string)|[範囲](range.md)|この範囲を占めるセルまたはセルの範囲をワークシートに挿入し、領域を空けるために他のセルをシフトします。この時点で空き領域に位置する、新しい Range オブジェクトが返されます。|
|[load(param: object)](#loadparam-object)|void|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|
|[merge(across: bool)](#mergeacross-bool)|void|範囲内のセルをワークシートの 1 つの領域にマージします。|
|[select()](#select)|void|Excel UI で指定した範囲を選択します。|
|[unmerge()](#unmerge)|void|範囲内のセルを結合解除して別々のセルにします。|

## メソッドの詳細


### clear(applyTo: string)
範囲の値、書式、塗りつぶし、罫線などをクリアします。

#### 構文
```js
rangeObject.clear(applyTo);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|applyTo|string|省略可能。クリア操作の種類を決定します。使用可能な値は次のとおりです。`All` (既定のオプション)、`Formats`、`Contents`|

#### 戻り値
void

#### 例

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


### delete(shift: string)
範囲に関連付けられているセルを削除します。

#### 構文
```js
rangeObject.delete(shift);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|<legacyBold>Shift</legacyBold>|string|セルをシフトする方向を指定します。使用可能な値は次のとおりです。Up、Left|

#### 戻り値
void

#### 例

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


### getBoundingRect(anotherRange: Range or string)
指定した範囲を包含する、最小の Range オブジェクトを取得します。たとえば、"B2:C5" と "D10:E15" の GetBoundingRect は、"B2:E15" になります。

#### 構文
```js
rangeObject.getBoundingRect(anotherRange);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|anotherRange|Range または string|Range オブジェクト、アドレスまたは範囲名。|

#### 戻り値
[範囲](range.md)

#### 例

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


### getCell(row: number, column: number)
行と列の番号に基づいて、1 つのセルを含んだ範囲オブジェクトを取得します。このセルは、ワークシートのグリッド内であれば、親の範囲の境界の外のセルであってもかまいません。返されるセルは、範囲の左上のセルを基準に配置されます。

#### 構文
```js
rangeObject.getCell(row, column);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|row|number|取得するセルの行番号。0 を起点とする番号になります。|
|列|number|取得セルの列番号。0 を起点とする番号になります。|

#### 戻り値
[範囲](range.md)

#### 例

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:F8";
    var worksheet = ctx.workbook.worksheets.getItem(sheetName);
    var range = worksheet.getRange(rangeAddress);
    var cell = range.getCell(0,0);
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


### getColumn(column: number)
範囲に含まれる列を 1 つ取得します。

#### 構文
```js
rangeObject.getColumn(column);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|column|number|取得する範囲の列番号。0 を起点とする番号になります。|

#### 戻り値
[範囲](range.md)

#### 例

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


### getEntireColumn()
範囲に含まれるすべての列を表すオブジェクトを取得します。

#### 構文
```js
rangeObject.getEntireColumn();
```

#### パラメーター
なし

#### 戻り値
[範囲](range.md)

#### 例

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

### getEntireRow()
範囲に含まれるすべての行を表すオブジェクトを取得します。

#### 構文
```js
rangeObject.getEntireRow();
```

#### パラメーター
なし

#### 戻り値
[範囲](range.md)

#### 例
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
Range のグリッド プロパティ (values、numberFormat、formulas) には、当該の範囲に境界がないため、`null` が格納されます。

### getIntersection(anotherRange: Range or string)
指定した範囲の長方形の交差を表す Range オブジェクトを取得します。

#### 構文
```js
rangeObject.getIntersection(anotherRange);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|anotherRange|Range または string|範囲の交差を判断するために使用される、Range オブジェクトまたは Range アドレス。|

#### 戻り値
[範囲](range.md)

#### 例

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


### getLastCell()
範囲内の最後のセルを取得します。たとえば、"B2:D5" の最後のセルは "D5" になります。

#### 構文
```js
rangeObject.getLastCell();
```

#### パラメーター
なし

#### 戻り値
[範囲](range.md)

#### 例

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


### getLastColumn()
範囲内の最後の列を取得します。たとえば、"B2:D5" の最後の列は "D2:D5" になります。

#### 構文
```js
rangeObject.getLastColumn();
```

#### パラメーター
なし

#### 戻り値
[範囲](range.md)

#### 例

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


### getLastRow()
範囲内の最後の行を取得します。たとえば、"B2:D5" の最後の行は "B5:D5" になります。

#### 構文
```js
rangeObject.getLastRow();
```

#### パラメーター
なし

#### 戻り値
[範囲](range.md)

#### 例

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



### getOffsetRange(rowOffset: number, columnOffset: number)
指定した範囲からのオフセットで範囲を表すオブジェクトを取得します。返される範囲のディメンションは、この範囲と一致します。結果の範囲が、ワークシートのグリッドの境界線の外にはみ出る場合は、例外がスローされます。

#### 構文
```js
rangeObject.getOffsetRange(rowOffset, columnOffset);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|rowOffset|number|範囲をオフセットする行数 (正、負、または 0)。正の値は下方向へのオフセットです。また、負の値は上方向へのオフセットです。|
|columnOffset|number|範囲をオフセットする列数 (正、負、または 0)。正の値は右方向へのオフセットです。また、負の値は左方向へのオフセットです。|

#### 戻り値
[範囲](range.md)

#### 例

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


### getRow(row: number)
範囲に含まれている行を 1 つ取得します。

#### 構文
```js
rangeObject.getRow(row);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|row|number|取得する範囲の行番号。0 を起点とする番号になります。|

#### 戻り値
[範囲](range.md)

#### 例

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


### getUsedRange(valuesOnly: bool)
指定した Range オブジェクトのうち使用されている範囲を返します。

#### 構文
```js
rangeObject.getUsedRange(valuesOnly);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|valuesOnly|bool|省略可能。true の場合、現在値が入っているセルのみが使用中のセルとされます。既定値の false の場合、これまで使用されたことのあるすべてのセルが使用中とされます。|

#### 戻り値
[範囲](range.md)

#### 例

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


### insert(shift: string)
この範囲を占めるセルまたはセルの範囲をワークシートに挿入し、領域を空けるために他のセルをシフトします。この時点で空き領域に位置する、新しい Range オブジェクトが返されます。

#### 構文
```js
rangeObject.insert(shift);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|<legacyBold>Shift</legacyBold>|string|セルをシフトする方向を指定します。使用可能な値は次のとおりです。Down、Right|

#### 戻り値
[範囲](range.md)

#### 例

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

### merge(across: bool)
範囲内のセルをワークシートの 1 つの領域にマージします。

#### 構文
```js
rangeObject.merge(across);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|across|bool|省略可能。指定した範囲のセルを行ごとに結合して、行ごとに別のセルを作成する場合は True に設定します。既定値は False です。|

#### 戻り値
void

#### 例
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


### select()
Excel UI で指定した範囲を選択します。

#### 構文
```js
rangeObject.select();
```

#### パラメーター
なし

#### 戻り値
void

#### 例

```js

Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "F5:F10"; 
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.select();
    return ctx.sync(); 
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### unmerge()
範囲内の結合済みセルを結合解除して別々のセルにします。

#### 構文
```js
rangeObject.unmerge();
```

#### パラメーター
なし

#### 戻り値
void

#### 例
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

### プロパティのアクセスの例

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

次の例では、2 x 3 のグリッドを含んでいるグリッドに対して、numberFormat、values、formulas を設定します。

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
次の例は、数式に R1C1 表記を使用する点を除いて、上記のものと同じです。

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "F5:G7";
    var numberFormat = [[null, "d-mmm"], [null, "d-mmm"], [null, null]]
    var values = [["Today", 42147], ["Tomorrow", "5/24"], ["Difference in days", null]];
    var formulasR1C1 = [[null,null], [null,null], [null,"=R[-1]C-R[-2]C"]];
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.numberFormat = numberFormat;
    range.values = values;
    range.formulasR1C1= formulasR1C1;
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
Excel.run(function (ctx) { 
    var names = ctx.workbook.names;
    var namedItem = names.getItem('MyRange');
    range = namedItem.range;
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

