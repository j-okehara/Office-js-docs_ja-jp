# <a name="chartpoint-object-javascript-api-for-excel"></a>ChartPoint オブジェクト (JavaScript API for Excel)

グラフの系列のポイントを表します。

## <a name="properties"></a>プロパティ

| プロパティ     | 型   |説明| 要件セット|
|:---------------|:--------|:----------|:----|
|value|object|グラフのポイントの値を返します。読み取り専用です。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="relationships"></a>リレーションシップ
| リレーションシップ | 型   |説明| 要件セット|
|:---------------|:--------|:----------|:----|
|format|[ChartPointFormat](chartpointformat.md)|グラフのポイントの書式設定プロパティをカプセル化します。読み取り専用です。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

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
