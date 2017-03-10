# <a name="rangefont-object-javascript-api-for-excel"></a>RangeFont オブジェクト (JavaScript API for Excel)

このオブジェクトは、オブジェクトのフォントの属性 (フォント名、フォント サイズ、色など) を表します。

## <a name="properties"></a>プロパティ

| プロパティ       | 型    |説明| 要件セット|
|:---------------|:--------|:----------|:----|
|bold|bool|フォントの太字の状態を表します。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|color|string|テキストの色の HTML カラー コード表記。たとえば、#FF0000 は赤を表します。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|italic|bool|フォントの斜体の状態を表します。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|name|string|フォント名 (例: "Calibri")|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|size|double|フォント サイズ。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|underline|string|フォントに適用する下線の種類。使用可能な値は次のとおりです。None、Single、Double、SingleAccountant、DoubleAccountant。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_プロパティのアクセスの[例を参照してください。](#property-access-examples)_

## <a name="relationships"></a>関係
なし


## <a name="methods"></a>メソッド
なし


## <a name="method-details"></a>メソッドの詳細

### <a name="property-access-examples"></a>プロパティのアクセスの例

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "F:G";
    var worksheet = ctx.workbook.worksheets.getItem(sheetName);
    var range = worksheet.getRange(rangeAddress);
    var rangeFont = range.format.font;
    rangeFont.load('name');
    return ctx.sync().then(function() {
        console.log(rangeFont.name);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
次の例では、フォント名を設定します。 

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "F:G";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.format.font.name = 'Times New Roman';
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```