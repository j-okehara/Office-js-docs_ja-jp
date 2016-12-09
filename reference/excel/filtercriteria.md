# <a name="filtercriteria-object-javascript-api-for-excel"></a>FilterCriteria オブジェクト (JavaScript API for Excel)

列に適用するフィルター条件を表します。

## <a name="properties"></a>プロパティ

| プロパティ     | 型   |説明| 要件セット|
|:---------------|:--------|:----------|:----|
|color|string|セルをフィルター処理するために使用する HTML カラー文字列。"CellColor" フィルターと "fontColor" フィルターで使用します。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|criterion1|string|データをフィルター処理するために使用する最初の条件。"カスタム" フィルター処理の場合には、演算子として使用されます。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|criterion2|string|データをフィルター処理するために使用する 2 番目の条件。"カスタム" フィルター処理の場合には、演算子としてのみ使用されます。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|dynamicCriteria|string|この列に適用する Excel.DynamicFilterCriteria の動的条件。「動的」フィルター処理で使用します。使用可能な値は次のいずれかです。Unknown、AboveAverage、AllDatesInPeriodApril、AllDatesInPeriodAugust、AllDatesInPeriodDecember、AllDatesInPeriodFebruray、AllDatesInPeriodJanuary、AllDatesInPeriodJuly、AllDatesInPeriodJune、AllDatesInPeriodMarch、AllDatesInPeriodMay、AllDatesInPeriodNovember、AllDatesInPeriodOctober、AllDatesInPeriodQuarter1、AllDatesInPeriodQuarter2、AllDatesInPeriodQuarter3、AllDatesInPeriodQuarter4、AllDatesInPeriodSeptember、BelowAverage、LastMonth、LastQuarter、LastWeek、LastYear、NextMonth、NextQuarter、NextWeek、NextYear、ThisMonth、ThisQuarter、ThisWeek、ThisYear、Today、Tomorrow、YearToDate、Yesterday。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|filterOn|string|値を表示したままにするかどうかを判別するために、フィルターで使用するプロパティ。使用可能な値は次のとおりです。BottomItems、BottomPercent、CellColor、Dynamic、FontColor、Values、TopItems、TopPercent、Icon、Custom。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|operator|string|"カスタム" フィルター処理を使用するときに、条件 1 と条件 2 を結合するために使用する演算子。使用可能な値は次のとおりです。And、Or。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|values|object[]|"値" フィルター処理の一部として使用する値のセット。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

_プロパティのアクセスの[例を参照してください。](#property-access-examples)_

## <a name="relationships"></a>関係
| リレーションシップ | 型   |説明| 要件セット|
|:---------------|:--------|:----------|:----|
|icon|[Icon](icon.md)|セルをフィルター処理するために使用するアイコン。"アイコン" フィルター処理で使用します。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>メソッド

| メソッド           | 戻り値の型    |説明| 要件セット|
|:---------------|:--------|:----------|:----|
|[load(param: object)](#loadparam-object)|(非推奨)|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>メソッドの詳細


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
