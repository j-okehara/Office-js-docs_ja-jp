# オブジェクトをフィルタリングする (JavaScript API for Excel)

_適用対象: Excel 2016、Excel Online、Excel for iOS、Office 2016_

テーブルの列のフィルター処理を管理します。

## プロパティ

なし

## 関係
| リレーションシップ | 型|説明|
|:---------------|:--------|:----------|
|条件|[FilterCriteria](filtercriteria.md)|指定した列に現在適用されているフィルターです。読み取り専用です。|

## メソッド

| メソッド   | 戻り値の型|説明|
|:---------------|:--------|:----------|
|[適用(条件:フィルター条件)](#applycriteria-filtercriteria)|void|指定した列に指定されたフィルター条件を適用します。次のヘルパー メソッドのどれでも、同じ機能を実現できます。|
|[applyBottomItemsFilter(count: number)](#applybottomitemsfiltercount-number)|void|指定した数の要素の列に [下位アイテム] フィルターを適用します。|
|[applyBottomPercentFilter(percent: number)](#applybottompercentfilterpercent-number)|void|指定したパーセンテージの要素の列に [下位パーセント] フィルターを適用します。|
|[applyCellColorFilter(color: string)](#applycellcolorfiltercolor-string)|void|指定した色の列に [セルの色] フィルターを適用します。|
|[applyCustomFilter(criteria1: string, criteria2: string, oper:FilterOperator)](#applycustomfiltercriteria1-string-criteria2-string-oper-filteroperator)|void|指定した条件の文字列の列に [アイコン] フィルターを適用します。|
|[applyDynamicFilter(criteria: string)](#applydynamicfiltercriteria-string)|void|列に [動的] フィルターを適用します。|
|[applyFontColorFilter(color: string)](#applyfontcolorfiltercolor-string)|void|指定した色の列に [フォントの色] フィルターを適用します。|
|[applyIconFilter(icon:Icon)](#applyiconfiltericon-icon)|void|指定したアイコンの列に [アイコン] フィルターを適用します。|
|[applyTopItemsFilter(count: number)](#applytopitemsfiltercount-number)|void|指定した数の要素の列に [上位アイテム] フィルターを適用します。|
|[applyTopPercentFilter(percent: number)](#applytoppercentfilterpercent-number)|void|指定したパーセンテージの要素の列に [上位パーセント] フィルターを適用します。|
|[applyValuesFilter(values: ()[])](#applyvaluesfiltervalues-)|void|指定した値の列に [値] フィルターを適用します。|
|[clear()](#clear)|void|指定した列のフィルターをクリアします。|
|[load(param: object)](#loadparam-object)|void|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|

## メソッドの詳細


### 適用(条件:フィルター条件)
指定した列に指定されたフィルター条件を適用します。次のヘルパー メソッドのどれでも、同じ機能を実現できます。 

#### 構文
```js
filterObject.apply(criteria);
```

#### パラメータ
| パラメーター   | 型|説明|
|:---------------|:--------|:----------|
|条件|FilterCriteria|適用する基準。|

#### 戻り値
void

#### 例
次の例はジェネリック apply() メソッドを使用してカスタム フィルターを適用する方法を示します。

```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    var filterCriteria = { 
		filterOn: Excel.FilterOn.custom,
		criterion1: ">50",
		operator: Excel.FilterOperator.and,
		criterion2: "<100"
    	} 
    column.filter.apply(filterCriteria);
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### applyBottomItemsFilter(count: number)
指定した数の要素の列に [下位アイテム] フィルターを適用します。

#### 構文
```js
filterObject.applyBottomItemsFilter(count);
```

#### パラメータ
| パラメーター   | 型|説明|
|:---------------|:--------|:----------|
|count|number|表示する下位からの要素の数。|

#### 戻り値
void

#### 例
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.applyBottomItemsFilter(3);
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### applyBottomPercentFilter(percent: number)
指定したパーセンテージの要素の列に [下位パーセント] フィルターを適用します。

#### 構文
```js
filterObject.applyBottomPercentFilter(percent);
```

#### パラメータ
| パラメーター   | 型|説明|
|:---------------|:--------|:----------|
|percent|number|表示する下位からの要素のパーセンテージ。|

#### 戻り値
void

#### 例
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.applyBottomPercentFilter(30);
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
### applyCellColorFilter(color: string)
指定した色の列に [セルの色] フィルターを適用します。


#### 構文
```js
filterObject.applyCellColorFilter(color);
```

#### パラメータ
| パラメーター   | 型|説明|
|:---------------|:--------|:----------|
|color|string|表示するセルの背景色です。|

#### 戻り値
void

#### 例
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.applyCellColorFilter('red');
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### applyCustomFilter(criteria1: string, criteria2: string, oper:FilterOperator)
指定した条件の文字列の列に [アイコン] フィルターを適用します。

#### 構文
```js
filterObject.applyCustomFilter(criteria1, criteria2, oper);
```

#### パラメータ
| パラメーター   | 型|説明|
|:---------------|:--------|:----------|
|criteria1|string|最初の条件の文字列です。|
|criteria2|string|省略可能。2 つ目の条件の文字列です。|
|oper|FilterOperator|省略可能。2 つの条件を結合する方法を記述する演算子です。|

#### 戻り値
void


#### 例
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.applyCustomFilter('>50','<100','and');
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### applyDynamicFilter(criteria: string)
列に [動的] フィルターを適用します。

#### 構文
```js
filterObject.applyDynamicFilter(criteria);
```

#### パラメータ
| パラメーター   | 型|説明|
|:---------------|:--------|:----------|
|criteria|string|適用する動的な条件です。使用可能な値は次のとおりです。Unknown、AboveAverage、AllDatesInPeriodApril、AllDatesInPeriodAugust、AllDatesInPeriodDecember、AllDatesInPeriodFebruray、AllDatesInPeriodJanuary、AllDatesInPeriodJuly、AllDatesInPeriodJune、AllDatesInPeriodMarch、AllDatesInPeriodMay、AllDatesInPeriodNovember、AllDatesInPeriodOctober、AllDatesInPeriodQuarter1、AllDatesInPeriodQuarter2、AllDatesInPeriodQuarter3、AllDatesInPeriodQuarter4、AllDatesInPeriodSeptember、BelowAverage、LastMonth、LastQuarter、LastWeek、LastYear、NextMonth、NextQuarter、NextWeek、NextYear、ThisMonth、ThisQuarter、ThisWeek、ThisYear、Today、Tomorrow、YearToDate、Yesterday|

#### 戻り値
void

#### 例
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.applyDynamicFilter(Excel.DynamicFilterCriteria.aboveAverage);
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### applyFontColorFilter(color: string)
指定した色の列に [フォントの色] フィルターを適用します。

#### 構文
```js
filterObject.applyFontColorFilter(color);
```

#### パラメータ
| パラメーター   | 型|説明|
|:---------------|:--------|:----------|
|color|string|表示するセルのフォントの色です。|

#### 戻り値
void

#### 例
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.applyFontColorFilter('red');
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### applyIconFilter(icon:Icon)
指定したアイコンの列に [アイコン] フィルターを適用します。

#### 構文
```js
filterObject.applyIconFilter(icon);
```

#### パラメータ
| パラメーター   | 型|説明|
|:---------------|:--------|:----------|
|アイコン|Icon|表示するセルのアイコンです。|

#### 戻り値
void

#### 例
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.applyIconFilter(Excel.icons.fiveArrows.yellowDownInclineArrow);
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### applyTopItemsFilter(count: number)
指定した数の要素の列に [上位アイテム] フィルターを適用します。

#### 構文
```js
filterObject.applyTopItemsFilter(count);
```

#### パラメータ
| パラメーター   | 型|説明|
|:---------------|:--------|:----------|
|count|number|表示する上位からの要素の数。|

#### 戻り値
void

#### 例
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.applyTopItemsFilter(3);
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### applyTopPercentFilter(percent: number)
指定したパーセンテージの要素の列に [上位パーセント] フィルターを適用します。

#### 構文
```js
filterObject.applyTopPercentFilter(percent);
```

#### パラメータ
| パラメーター   | 型|説明|
|:---------------|:--------|:----------|
|percent|number|表示する上位からの要素のパーセンテージ。|

#### 戻り値
void

#### 例
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.applyTopPercentFilter(30);
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
### applyValuesFilter(values: ()[])
指定した値の列に [値] フィルターを適用します。

#### 構文
```js
filterObject.applyValuesFilter(values);
```

#### パラメータ
| パラメーター   | 型|説明|
|:---------------|:--------|:----------|
|values|()[]|表示する値のリスト。|

#### 戻り値
void

#### 例
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.applyValuesFilter(['a','b']);
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### clear()
指定した列のフィルターをクリアします。

#### 構文
```js
filterObject.clear();
```

#### パラメーター
なし

#### 戻り値
void

#### 例
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.clear();
    return ctx.sync(); 
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

#### パラメータ
| パラメーター   | 型|説明|
|:---------------|:--------|:----------|
|param|object|省略可能。パラメーター名とリレーションシップ名を、区切られた文字列または 1 つの配列として受け入れます。あるいは、[loadOption](loadoption.md) オブジェクトを提供します。|

#### 戻り値
void

