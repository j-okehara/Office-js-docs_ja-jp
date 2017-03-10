# <a name="rangeborder-object-javascript-api-for-excel"></a>RangeBorder オブジェクト (JavaScript API for Excel)

オブジェクトの輪郭を表します。

## <a name="properties"></a>プロパティ

| プロパティ       | 型    |説明| 要件セット|
|:---------------|:--------|:----------|:----|
|color|string|枠線の色を表す HTML カラー コード。形式は #RRGGBB (例: "FFA500")、または名前付きの HTML 色 (例: "オレンジ") です。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|id|string|罫線の識別子を表します。読み取り専用です。使用可能な値は次のとおりです。EdgeTop、EdgeBottom、EdgeLeft、EdgeRight、InsideVertical、InsideHorizontal、DiagonalDown、DiagonalUp。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|sideIndex|string|罫線の特定の辺を表す定数値。読み取り専用です。使用可能な値は次のとおりです。EdgeTop、EdgeBottom、EdgeLeft、EdgeRight、InsideVertical、InsideHorizontal、DiagonalDown、DiagonalUp。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|style|string|罫線の線スタイルを指定する、線スタイル定数のいずれか 1 つ。使用可能な値は次のとおりです。None、Continuous、Dash、DashDot、DashDotDot、Dot、Double、SlantDashDot。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|weight|string|範囲を取り囲む罫線の太さを指定します。使用可能な値は次のとおりです。Hairline、Thin、Medium、Thick。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

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
    var rangeAddress = "A1:F8";
    var worksheet = ctx.workbook.worksheets.getItem(sheetName);
    var range = worksheet.getRange(rangeAddress);
    var borders = range.format.borders;
    borders.load('items');
    return ctx.sync().then(function() {
        console.log(borders.count);
        for (var i = 0; i < borders.items.length; i++)
        {
            console.log(borders.items[i].sideIndex);
        }
    });
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
    var rangeAddress = "A1:F8";
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

