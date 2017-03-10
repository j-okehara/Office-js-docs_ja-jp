# <a name="position-object-javascript-api-for-visio"></a>Position オブジェクト (JavaScript API for Visio)

適用対象:_Visio Online_

ビュー内のオブジェクトの位置を表します。

## <a name="properties"></a>プロパティ

| プロパティ       | 型    |説明|
|:---------------|:--------|:----------|
|x|int|オブジェクトの x 座標を指定する整数です。これは、ビューポートの中心からページの左境界までの距離 (ピクセル単位) の符号付きの値です。|
|y|int|オブジェクトの y 座標を指定する整数です。これは、ビューポートの中心からページの上境界までの距離 (ピクセル単位) の符号付きの値です。|

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
