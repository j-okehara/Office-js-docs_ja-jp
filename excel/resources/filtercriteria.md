# FilterCriteria オブジェクト (JavaScript API for Excel)

_適用対象: Excel 2016、Excel Online、Excel for iOS、Office 2016_

列に適用するフィルター条件を表します。

## プロパティ

| プロパティ   | 型|説明
|:---------------|:--------|:----------|
|color|string|セルをフィルター処理するために使用する HTML カラー文字列。「CellColor」フィルターおよび「fontColor」フィルターと併用します。|
|criterion1|string|データをフィルター処理するために使用する最初の条件。「カスタム」フィルター処理の場合には、演算子として使用されます。|
|criterion2|string|データをフィルター処理するために使用する 2 番目の条件。「カスタム」フィルター処理の場合には、演算子としてのみ使用されます。|
|dynamicCriteria|string|この列に適用する Excel.DynamicFilterCriteria の動的条件。「動的」フィルター処理で使用します。使用可能な値は次のいずれかです。Unknown、AboveAverage、AllDatesInPeriodApril、AllDatesInPeriodAugust、AllDatesInPeriodDecember、AllDatesInPeriodFebruray、AllDatesInPeriodJanuary、AllDatesInPeriodJuly、AllDatesInPeriodJune、AllDatesInPeriodMarch、AllDatesInPeriodMay、AllDatesInPeriodNovember、AllDatesInPeriodOctober、AllDatesInPeriodQuarter1、AllDatesInPeriodQuarter2、AllDatesInPeriodQuarter3、AllDatesInPeriodQuarter4、AllDatesInPeriodSeptember、BelowAverage、LastMonth、LastQuarter、LastWeek、LastYear、NextMonth、NextQuarter、NextWeek、NextYear、ThisMonth、ThisQuarter、ThisWeek、ThisYear、Today、Tomorrow、YearToDate、Yesterday。|
|filterOn|string|値を表示するかどうかを判別するために、フィルターで使用するプロパティ。使用可能な値は次のいずれかです。	BottomItems、BottomPercent、CellColor、Dynamic、FontColor、Values、TopItems、TopPercent、Icon、Custom |
|values|object[]|「値」フィルター処理の一部として使用する値のセット。|

## リレーションシップ
| リレーションシップ | 型|説明|
|:---------------|:--------|:----------|
|アイコン|[Icon](icon.md)|セルをフィルター処理するために使用するアイコン。「アイコン」フィルター処理で使用します。|
|operator|FilterOperator|「カスタム」フィルター処理を使用するときに、条件 1 と条件 2 を結合するために使用する演算子。|

## メソッド

| メソッド   | 戻り値の型|説明|
|:---------------|:--------|:----------|
|[load(param: object)](#loadparam-object)|void|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|

## メソッドの詳細


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

