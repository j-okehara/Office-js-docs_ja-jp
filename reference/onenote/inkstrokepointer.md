# InkStrokePointer オブジェクト (JavaScript API for OneNote)

_適用対象:OneNote Online_  


インク ストローク オブジェクトとそのコンテンツの親への弱い参照

## プロパティ

| プロパティ     | 型   |説明|フィードバック|
|:---------------|:--------|:----------|:-------|
|contentId|string|このストロークに対応するページ コンテンツ オブジェクトの id を表します。|[実行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkStrokePointer-contentId)|
|inkStrokeId|string|インク ストロークの id を表します。|[実行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkStrokePointer-inkStrokeId)|

_プロパティのアクセスの[例](#例)を参照してください。_

## リレーションシップ
なし


## メソッド

| メソッド           | 戻り値の型    |説明| フィードバック|
|:---------------|:--------|:----------|:-------|
|[load(param: object)](#loadparam-object)|void|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkStrokePointer-load)|

## メソッドの詳細


### load(param: object)
JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。

#### 構文
```js
object.load(param);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|param|object|省略可能。パラメーターとリレーションシップ名を、区切られた文字列または 1 つの配列として受け入れます。あるいは、[loadOption](loadoption.md) オブジェクトを提供します。|

#### 戻り値
void
