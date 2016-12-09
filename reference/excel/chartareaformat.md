# <a name="chartareaformat-object-javascript-api-for-excel"></a>ChartAreaFormat オブジェクト (JavaScript API for Excel)

グラフ領域全体の書式設定プロパティをカプセル化します。

## <a name="properties"></a>プロパティ

なし

## <a name="relationships"></a>関係
| リレーションシップ | 型   |説明| 要件セット|
|:---------------|:--------|:----------|:----|
|fill|[ChartFill](chartfill.md)|背景の書式設定情報を含む、オブジェクトの塗りつぶしの書式を表します。読み取り専用です。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|font|[ChartFont](chartfont.md)|現在のオブジェクトのフォント属性 (フォント名、フォント サイズ、色など) を表します。読み取り専用です。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

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
