# <a name="chartfont-object-javascript-api-for-excel"></a>ChartFont オブジェクト (JavaScript API for Excel)

このオブジェクトは、グラフ オブジェクトのフォント属性 (フォント名、フォント サイズ、色など) を表します。

## <a name="properties"></a>プロパティ

| プロパティ       | 型    |説明| 要件セット|
|:---------------|:--------|:----------|:----|
|bold|bool|フォントの太字の状態を表します。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|color|string|テキストの色の HTML カラー コード表記。たとえば、#FF0000 は赤を表します。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|italic|bool|フォントの斜体の状態を表します。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|name|string|フォント名 (例: "Calibri")|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|size|double|フォント サイズ (例: 11)|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|underline|string|フォントに適用する下線の種類。使用可能な値は次のとおりです。なし、一重線。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_プロパティのアクセスの[例を参照してください。](#property-access-examples)_

## <a name="relationships"></a>関係
なし


## <a name="methods"></a>メソッド
なし


## <a name="method-details"></a>メソッドの詳細

### <a name="property-access-examples"></a>プロパティのアクセスの例

グラフのタイトルを例として使用します。

```js
Excel.run(function (ctx) { 
    var title = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").title;
    title.format.font.name = "Calibri";
    title.format.font.size = 12;
    title.format.font.color = "#FF0000";
    title.format.font.italic =  false;
    title.format.font.bold = true;
    title.format.font.underline = "None";
    return ctx.sync();
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

グラフのタイトルを Calibri、サイズ 10、太字、赤色に設定します。 

```js
Excel.run(function (ctx) { 
    var title = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").title;
    title.format.font.name = "Calibri";
    title.format.font.size = 12;
    title.format.font.color = "#FF0000";
    title.format.font.italic =  false;
    title.format.font.bold = true;
    title.format.font.underline = "None";
    return ctx.sync();
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
