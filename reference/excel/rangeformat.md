# RangeFormat オブジェクト (JavaScript API for Excel)

範囲のフォント、塗りつぶし、境界線、配置などのプロパティをカプセル化する、書式設定オブジェクトです。

## プロパティ

| プロパティ     | 型   |説明
|:---------------|:--------|:----------|
|columnWidth|double|範囲内のすべての列の幅を取得または設定します。列の幅が均一でない場合は、null が返されます。|
|horizontalAlignment|string|指定したオブジェクトの水平方向の配置を表します。使用可能な値は次のとおりです。General、Left、Center、Right、Fill、Justify、CenterAcrossSelection、Distributed。|
|rowHeight|double|範囲内のすべての行の高さを取得または設定します。行の高さが均一でない場合は、null が返されます。|
|verticalAlignment|string|指定したオブジェクトの垂直方向の配置を表します。使用可能な値は次のとおりです。Top、Center、Bottom、Justify、Distributed。|
|wrapText|bool|Excel テキスト コントロールがオブジェクト内のテキストをラップするよう設定されていることを表します。null 値は、範囲全体で一様なラップ テキスト設定が使用されないことを表します。|

_プロパティのアクセスの[例](#例)をご覧ください。_

## リレーションシップ
| リレーションシップ | 型   |説明|
|:---------------|:--------|:----------|
|Borders|[RangeBorderCollection](rangebordercollection.md)|選択した範囲全体に適用する Border オブジェクトのコレクションです。読み取り専用です。|
|fill|[RangeFill](rangefill.md)|範囲全体に定義された塗りつぶしオブジェクトを返します。読み取り専用です。|
|Font|[RangeFont](rangefont.md)|選択した範囲全体に定義されているフォント オブジェクトを返します。読み取り専用です。|
|protection|[FormatProtection](formatprotection.md)|範囲に対する書式保護オブジェクトを返します。読み取り専用です。|

## メソッド

| メソッド           | 戻り値の型    |説明|
|:---------------|:--------|:----------|
|[autofitColumns()](#autofitcolumns)|void|現在の列のデータに基づいて、現在の範囲の列の幅を最適な幅に変更します。|
|[autofitRows()](#autofitrows)|void|現在の行のデータに基づいて、現在の範囲の行の高さを最適な高さに変更します。|
|[load(param: object)](#loadparam-object)|void|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|

## メソッドの詳細


### autofitColumns()
現在の列のデータに基づいて、現在の範囲の列の幅を最適な幅に変更します。

#### 構文
```js
rangeFormatObject.autofitColumns();
```

#### パラメーター
なし

#### 戻り値
void

### autofitRows()
現在の行のデータに基づいて、現在の範囲の行の高さを最適な高さに変更します。

#### 構文
```js
rangeFormatObject.autofitRows();
```

#### パラメーター
なし

#### 戻り値
void

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

次の例では、範囲のすべての書式設定プロパティを出力します。 

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

次の例では、範囲のフォント名と塗りつぶし色を設定し、テキストをラップしています。 

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
    range.format.borders('InsideHorizontal').lineStyle = 'Continuous';
    range.format.borders('InsideVertical').lineStyle = 'Continuous';
    range.format.borders('EdgeBottom').lineStyle = 'Continuous';
    range.format.borders('EdgeLeft').lineStyle = 'Continuous';
    range.format.borders('EdgeRight').lineStyle = 'Continuous';
    range.format.borders('EdgeTop').lineStyle = 'Continuous';
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
