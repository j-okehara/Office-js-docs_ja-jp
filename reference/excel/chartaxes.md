# <a name="chartaxes-object-javascript-api-for-excel"></a>ChartAxes オブジェクト (JavaScript API for Excel)

グラフの軸を表します。

## <a name="properties"></a>プロパティ

なし

## <a name="relationships"></a>関係
| リレーションシップ | 型   |説明| 要件セット|
|:---------------|:--------|:----------|:----|
|categoryAxis|[ChartAxis](chartaxis.md)|グラフの項目軸を表します。読み取り専用です。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|seriesAxis|[ChartAxis](chartaxis.md)|3 次元グラフの系列軸を表します。読み取り専用です。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|valueAxis|[ChartAxis](chartaxis.md)|軸の数値軸を表します。読み取り専用です。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

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
