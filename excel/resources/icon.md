# Icon オブジェクト (JavaScript API for Excel)

_適用対象:Excel 2016、Excel Online、Excel for iOS、Office 2016_

セルのアイコンを表します。

## プロパティ

| プロパティ	  | 型	|説明
|:---------------|:--------|:----------||index|int|指定セットのアイコンのインデックスを表します。||set|string|対象アイコンが含まれるセットを表しています。使用可能な値は次のいずれかです。Invalid、ThreeArrows、ThreeArrowsGray、ThreeFlags、ThreeTrafficLights1、ThreeTrafficLights2、ThreeSigns、ThreeSymbols、ThreeSymbols2、FourArrows、FourArrowsGray、FourRedToBlack、FourRating、FourTrafficLights、FiveArrows、FiveArrowsGray、FiveRating、FiveQuarters、ThreeStars、ThreeTriangles、FiveBoxes。|_プロパティの使用[例](#property-access-examples)_をご覧ください。

## リレーションシップ
なし


## メソッド

| メソッド		  | 戻り値の型	|説明||:---------------|:--------|:----------||[load(param: object)](#loadparam-object)|void|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|

## メソッドの詳細


### load(param: object)
JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。

#### 構文
```js
object.load(param);
```

#### パラメーター
| パラメーター	  | 型	|説明||:---------------|:--------|:----------||param|object|省略可能。パラメーターとリレーションシップ名を、区切られた文字列または 1 つの配列として受け入れます。あるいは、[loadOption](loadoption.md) オブジェクトを提供します。|

#### 戻り値
void

