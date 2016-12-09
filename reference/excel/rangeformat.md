# <a name="rangeformat-object-javascript-api-for-excel"></a>RangeFormat オブジェクト (JavaScript API for Excel)

範囲のフォント、塗りつぶし、境界線、配置などのプロパティをカプセル化する、書式設定オブジェクトです。

## <a name="properties"></a>プロパティ

| プロパティ     | 型   |説明| 要件セット|
|:---------------|:--------|:----------|:----|
|columnWidth|double|範囲内のすべての列の幅を取得または設定します。列の幅が同じでない場合は、null が返されます。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|horizontalAlignment|string|指定したオブジェクトの水平方向の配置を表します。使用可能な値は次のとおりです。General、Left、Center、Right、Fill、Justify、CenterAcrossSelection、Distributed。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|rowHeight|double|範囲内のすべての行の高さを取得または設定します。行の高さが同じでない場合は、null が返されます。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|verticalAlignment|string|指定したオブジェクトの垂直方向の配置を表します。使用可能な値は次のとおりです。Top、Center、Bottom、Justify、Distributed。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|wrapText|bool|オブジェクト内のテキストを Excel でラップするかどうかを表します。null 値は、範囲全体に一様なラップ設定がないことを表します。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_プロパティのアクセスの[例を参照してください。](#property-access-examples)_

## <a name="relationships"></a>関係
| リレーションシップ | 型   |説明| 要件セット|
|:---------------|:--------|:----------|:----|
|borders|[RangeBorderCollection](rangebordercollection.md)|選択した範囲全体に適用する境界線オブジェクトのコレクションです。読み取り専用です。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|fill|[RangeFill](rangefill.md)|範囲全体に定義された塗りつぶしオブジェクトを返します。読み取り専用です。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|font|[RangeFont](rangefont.md)|選択した範囲全体に定義されているフォント オブジェクトを返します。読み取り専用です。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|protection|[FormatProtection](formatprotection.md)|範囲に対する書式保護オブジェクトを返します。読み取り専用です。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>メソッド

| メソッド           | 戻り値の型    |説明| 要件セット|
|:---------------|:--------|:----------|:----|
|[autofitColumns()](#autofitcolumns)|void|現在の列のデータに基づいて、現在の範囲の列の幅を最適な幅に変更します。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[autofitRows()](#autofitrows)|void|現在の行のデータに基づいて、現在の範囲の行の高さを最適な高さに変更します。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[load(param: object)](#loadparam-object)|(非推奨)|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>メソッドの詳細


### <a name="autofitcolumns"></a>autofitColumns()
現在の列のデータに基づいて、現在の範囲の列の幅を最適な幅に変更します。

#### <a name="syntax"></a>構文
```js
rangeFormatObject.autofitColumns();
```

#### <a name="parameters"></a>パラメーター
なし

#### <a name="returns"></a>戻り値
void

### <a name="autofitrows"></a>autofitRows()
現在の行のデータに基づいて、現在の範囲の行の高さを最適な高さに変更します。

#### <a name="syntax"></a>構文
```js
rangeFormatObject.autofitRows();
```

#### <a name="parameters"></a>パラメーター
なし

#### <a name="returns"></a>戻り値
void

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

次の例では、範囲の書式設定プロパティをすべて選択しています。 

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "F:G";
    var worksheet = ctx.workbook.worksheets.getItem(sheetName);
    var range = worksheet.getRange(rangeAddress);
    range.load(["format/*", "format/fill", "format/borders", "format/font"]);
    return ctx.sync().then(function() {
        console.log(range.format.wrapText);
        console.log(range.format.fill.color);
        console.log(range.format.font.name);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

次の例では、フォント名、塗りつぶし色およびテキストのラップを設定しています。 

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "F:G";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.format.wrapText = true;
    range.format.font.name = 'Times New Roman';
    range.format.fill.color = '0000FF';
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

次の例では、範囲を取り囲むグリッドの境界線を追加しています。

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "F:G";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.format.borders.getItem('InsideHorizontal').style = 'Continuous';
    range.format.borders.getItem('InsideVertical').style = 'Continuous';
    range.format.borders.getItem('EdgeBottom').style = 'Continuous';
    range.format.borders.getItem('EdgeLeft').style = 'Continuous';
    range.format.borders.getItem('EdgeRight').style = 'Continuous';
    range.format.borders.getItem('EdgeTop').style = 'Continuous';
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```