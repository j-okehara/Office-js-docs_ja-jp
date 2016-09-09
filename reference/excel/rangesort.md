# RangeSort オブジェクト (JavaScript API for Excel)

_適用対象: Excel 2016、Excel Online、Excel for iOS、Office 2016_

Range オブジェクトの並べ替え操作を管理します。

## プロパティ

なし

## 関係
なし


## メソッド

| メソッド           | 戻り値の型    |説明|
|:---------------|:--------|:----------|
|[apply(fields:SortField[], matchCase: bool, hasHeaders: bool, orientation: string, method: string)](#applyfields-sortfield-matchcase-bool-hasheaders-bool-orientation-string-method-string)|void|並べ替え操作を実行します。|

## メソッドの詳細


### apply(fields:SortField[], matchCase: bool, hasHeaders: bool, orientation: string, method: string)
並べ替え操作を実行します。

#### 構文
```js
rangeSortObject.apply(fields, matchCase, hasHeaders, orientation, method);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|fields|SortField[]|並べ替えに使用する条件の一覧。|
|matchCase|bool|省略可能。大文字小文字の区別が文字列の順序に影響を与えるかどうか。|
|hasHeaders|bool|省略可能。範囲にヘッダーがあるかどうか。|
|orientation|string|省略可能。操作が行と列のどちらの並べ替えかを示します。使用可能な値は次のとおりです。Rows、Columns|
|method|string|省略可能。中国語文字に使用される順序付けの方法です。使用可能な値は次のとおりです。PinYin、StrokeCount|

#### 戻り値
void

#### 例
```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "D4:G6";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.sort.apply([ 
            {
                key: 2,
                ascending: true
            },
        ], true);
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```