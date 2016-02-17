# ChartAxes オブジェクト (JavaScript API for Excel)

_適用対象: Excel 2016、Excel Online、Office 2016_

グラフの軸を表します。

## プロパティ

なし

## 関係
| リレーションシップ | 型|説明|
|:---------------|:--------|:----------|
|categoryAxis|[ChartAxis](chartaxis.md)|グラフの項目軸を表します。値の取得のみ可能です。|
|seriesAxis|[ChartAxis](chartaxis.md)|3 次元グラフの系列軸を表します。値の取得のみ可能です。|
|valueAxis|[ChartAxis](chartaxis.md)|軸の数値軸を表します。値の取得のみ可能です。|

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

