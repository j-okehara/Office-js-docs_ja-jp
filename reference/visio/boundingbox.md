# <a name="boundingbox-object-javascript-api-for-visio"></a>BoundingBox オブジェクト (JavaScript API for Visio)

適用対象:_Visio Online_

図形の BoundingBox を表します。

## <a name="properties"></a>プロパティ

| プロパティ       | 型    |説明|
|:---------------|:--------|:----------|
|height|int|図形に関連付けられたデータ グラフィックスを除く、図形の境界ボックスの上端と下端の間の距離です。|
|width|int|図形に関連付けられたデータ グラフィックスを除く、図形の境界ボックスの左端と右端の間の距離です。|
|x|int|境界ボックスの x 座標を指定する整数です。|
|y|int|境界ボックスの y 座標を指定する整数です。|

_プロパティのアクセスの[例を参照してください。](#property-access-examples)_

## <a name="relationships"></a>関係
なし


## <a name="methods"></a>メソッド

| メソッド           | 戻り値の型    |説明|
|:---------------|:--------|:----------|
|[load(param: object)](#loadparam-object)|(非推奨)|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|

## <a name="method-details"></a>メソッドの詳細


### <a name="loadparam-object"></a>load(param: object)
JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。

#### <a name="syntax"></a>構文
```js
object.load(param);
```

#### <a name="parameters"></a>パラメーター
| パラメーター       | 型    |説明|
|:---------------|:--------|:----------|:---|
|param|object|省略可能。パラメーターとリレーションシップ名を、区切られた文字列または 1 つの配列として受け入れます。あるいは、[loadOption](loadoption.md) オブジェクトを提供します。|

#### <a name="returns"></a>戻り値
void
