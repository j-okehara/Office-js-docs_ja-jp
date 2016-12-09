# <a name="filter-object-javascript-api-for-excel"></a>Filter オブジェクト (JavaScript API for Excel)

テーブルの列のフィルター処理を管理します。

## <a name="properties"></a>プロパティ

なし

## <a name="relationships"></a>関係
| リレーションシップ | 型   |説明| 要件セット|
|:---------------|:--------|:----------|:----|
|criteria|[FilterCriteria](filtercriteria.md)|指定した列に現在適用されているフィルターです。読み取り専用です。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>メソッド

| メソッド           | 戻り値の型    |説明| 要件セット|
|:---------------|:--------|:----------|:----|
|[apply(criteria:FilterCriteria)](#applycriteria-filtercriteria)|void|指定した列に、指定されたフィルター条件を適用します。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[applyBottomItemsFilter(count: number)](#applybottomitemsfiltercount-number)|void|指定した数の要素の列に "下位アイテム" フィルターを適用します。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[applyBottomPercentFilter(percent: number)](#applybottompercentfilterpercent-number)|void|指定した割合の要素の列に "下位パーセント" フィルターを適用します。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[applyCellColorFilter(color: string)](#applycellcolorfiltercolor-string)|void|指定した色の列に "セルの色" フィルターを適用します。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[applyCustomFilter(criteria1: string, criteria2: string, oper: string)](#applycustomfiltercriteria1-string-criteria2-string-oper-string)|void|指定した条件の文字列の列に "アイコン" フィルターを適用します。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[applyDynamicFilter(criteria: string)](#applydynamicfiltercriteria-string)|void|列に "動的" フィルターを適用します。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[applyFontColorFilter(color: string)](#applyfontcolorfiltercolor-string)|void|指定した色の列に "フォントの色" フィルターを適用します。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[applyIconFilter(icon:Icon)](#applyiconfiltericon-icon)|void|指定したアイコンの列に "アイコン" フィルターを適用します。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[applyTopItemsFilter(count: number)](#applytopitemsfiltercount-number)|void|指定した数の要素の列に "上位アイテム" フィルターを適用します。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[applyTopPercentFilter(percent: number)](#applytoppercentfilterpercent-number)|void|指定した割合の要素の列に "上位パーセント" フィルターを適用します。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[applyValuesFilter(values: ()[])](#applyvaluesfiltervalues-)|void|指定した値の列に "値" フィルターを適用します。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[clear()](#clear)|void|指定した列のフィルターをクリアします。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[load(param: object)](#loadparam-object)|(非推奨)|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>メソッドの詳細


### <a name="applycriteria-filtercriteria"></a>適用(条件:フィルター条件)
指定した列に、指定されたフィルター条件を適用します。

#### <a name="syntax"></a>構文
```js
filterObject.apply(criteria);
```

#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|:---|
|条件|FilterCriteria|適用する基準。|

#### <a name="returns"></a>戻り値
void

### <a name="applybottomitemsfiltercount-number"></a>applyBottomItemsFilter(count: number)
指定した数の要素の列に [下位アイテム] フィルターを適用します。

#### <a name="syntax"></a>構文
```js
filterObject.applyBottomItemsFilter(count);
```

#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|:---|
|count|number|表示する下位からの要素の数。|

#### <a name="returns"></a>戻り値
void

### <a name="applybottompercentfilterpercent-number"></a>applyBottomPercentFilter(percent: number)
指定したパーセンテージの要素の列に [下位パーセント] フィルターを適用します。

#### <a name="syntax"></a>構文
```js
filterObject.applyBottomPercentFilter(percent);
```

#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|:---|
|percent|number|表示する下位からの要素のパーセンテージ。|

#### <a name="returns"></a>戻り値
void

### <a name="applycellcolorfiltercolor-string"></a>applyCellColorFilter(color: string)
指定した色の列に [セルの色] フィルターを適用します。

#### <a name="syntax"></a>構文
```js
filterObject.applyCellColorFilter(color);
```

#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|:---|
|color|string|表示するセルの背景色です。|

#### <a name="returns"></a>戻り値
void

### <a name="applycustomfiltercriteria1-string-criteria2-string-oper-string"></a>applyCustomFilter(criteria1: string, criteria2: string, oper: string)
指定した条件の文字列の列に "アイコン" フィルターを適用します。

#### <a name="syntax"></a>構文
```js
filterObject.applyCustomFilter(criteria1, criteria2, oper);
```

#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|:---|
|criteria1|string|最初の条件の文字列です。|
|criteria2|string|省略可能。2 つ目の条件の文字列です。|
|oper|string|省略可能。2 つの条件を結合する方法を記述する演算子です。使用可能な値は次のとおりです。And、Or|

#### <a name="returns"></a>戻り値
void

### <a name="applydynamicfiltercriteria-string"></a>applyDynamicFilter(criteria: string)
列に [動的] フィルターを適用します。

#### <a name="syntax"></a>構文
```js
filterObject.applyDynamicFilter(criteria);
```

#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|:---|
|criteria|string|適用する動的な条件です。使用可能な値は次のとおりです。Unknown、AboveAverage、AllDatesInPeriodApril、AllDatesInPeriodAugust、AllDatesInPeriodDecember、AllDatesInPeriodFebruray、AllDatesInPeriodJanuary、AllDatesInPeriodJuly、AllDatesInPeriodJune、AllDatesInPeriodMarch、AllDatesInPeriodMay、AllDatesInPeriodNovember、AllDatesInPeriodOctober、AllDatesInPeriodQuarter1、AllDatesInPeriodQuarter2、AllDatesInPeriodQuarter3、AllDatesInPeriodQuarter4、AllDatesInPeriodSeptember、BelowAverage、LastMonth、LastQuarter、LastWeek、LastYear、NextMonth、NextQuarter、NextWeek、NextYear、ThisMonth、ThisQuarter、ThisWeek、ThisYear、Today、Tomorrow、YearToDate、Yesterday|

#### <a name="returns"></a>戻り値
void

### <a name="applyfontcolorfiltercolor-string"></a>applyFontColorFilter(color: string)
指定した色の列に [フォントの色] フィルターを適用します。

#### <a name="syntax"></a>構文
```js
filterObject.applyFontColorFilter(color);
```

#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|:---|
|color|string|表示するセルのフォントの色です。|

#### <a name="returns"></a>戻り値
void

### <a name="applyiconfiltericon-icon"></a>applyIconFilter(icon:Icon)
指定したアイコンの列に [アイコン] フィルターを適用します。

#### <a name="syntax"></a>構文
```js
filterObject.applyIconFilter(icon);
```

#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|:---|
|アイコン|Icon|表示するセルのアイコンです。|

#### <a name="returns"></a>戻り値
void

### <a name="applytopitemsfiltercount-number"></a>applyTopItemsFilter(count: number)
指定した数の要素の列に [上位アイテム] フィルターを適用します。

#### <a name="syntax"></a>構文
```js
filterObject.applyTopItemsFilter(count);
```

#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|:---|
|count|number|表示する上位からの要素の数。|

#### <a name="returns"></a>戻り値
void

### <a name="applytoppercentfilterpercent-number"></a>applyTopPercentFilter(percent: number)
指定したパーセンテージの要素の列に [上位パーセント] フィルターを適用します。

#### <a name="syntax"></a>構文
```js
filterObject.applyTopPercentFilter(percent);
```

#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|:---|
|percent|number|表示する上位からの要素のパーセンテージ。|

#### <a name="returns"></a>戻り値
void

### <a name="applyvaluesfiltervalues-"></a>applyValuesFilter(values: ()[])
指定した値の列に [値] フィルターを適用します。

#### <a name="syntax"></a>構文
```js
filterObject.applyValuesFilter(values);
```

#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|:---|
|values|()[]|表示する値のリスト。|

#### <a name="returns"></a>戻り値
void

### <a name="clear"></a>clear()
指定した列のフィルターをクリアします。

#### <a name="syntax"></a>構文
```js
filterObject.clear();
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
