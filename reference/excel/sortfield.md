# <a name="sortfield-object-javascript-api-for-excel"></a>SortField オブジェクト (JavaScript API for Excel)

並べ替え操作の条件を表します。

## <a name="properties"></a>プロパティ

| プロパティ     | 型   |説明| 要件セット|
|:---------------|:--------|:----------|:----|
|ascending|bool|昇順で並べ替えるかどうかを表します。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|color|string|並べ替えがフォントまたはセルの色で行われる場合に、条件の対象となる色を表します。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|dataOption|string|このフィールドのその他の並べ替えオプションを表します。使用可能な値は次のとおりです。Normal、TextAsNumber。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|key|int|条件の対象とする列 (または行。並べ替えの方向によって異なります) を表します。最初の列 (または行) からのオフセットとして表します。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|sortOn|string|この条件の並べ替えの種類を表します。使用可能な値は次のとおりです。Value、CellColor、FontColor、Icon。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

_プロパティのアクセスの[例を参照してください。](#property-access-examples)_

## <a name="relationships"></a>関係
| リレーションシップ | 型   |説明| 要件セット|
|:---------------|:--------|:----------|:----|
|icon|[Icon](icon.md)|並べ替えがセルのアイコンで行われる場合に、条件の対象となるアイコンを表します。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

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
