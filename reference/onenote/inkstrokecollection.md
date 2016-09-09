# InkStrokeCollection オブジェクト (JavaScript API for OneNote)

_適用対象:OneNote Online_   


InkStroke オブジェクトのコレクションを表します。

## プロパティ

| プロパティ     | 型   |説明|フィードバック|
|:---------------|:--------|:----------|:-------|
|count|int|ページ内の InkStroke の数を返します。 読み取り専用です。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkStrokeCollection-count)|
|items|[InkStroke[]](inkstroke.md)|InkStroke オブジェクトのコレクション。 読み取り専用です。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkStrokeCollection-items)|

_プロパティのアクセスの[例](#例)を参照してください。_

## リレーションシップ
なし


## メソッド

| メソッド           | 戻り値の型    |説明| フィードバック|
|:---------------|:--------|:----------|:-------|
|[getItem(index: number または string)](#getitemindex-number-または-string)|[InkStroke](inkstroke.md)|ID かコレクション内のインデックスにより、InkStroke オブジェクトを取得します。 読み取り専用です。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkStrokeCollection-getItem)|
|[getItemAt(index: number)](#getitematindex-number)|[InkStroke](inkstroke.md)|コレクション内での位置を基に InkStroke を取得します。|[実行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkStrokeCollection-getItemAt)|
|[load(param: object)](#loadparam-object)|void|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkStrokeCollection-load)|

## メソッドの詳細


### getItem(index: number または string)
ID かコレクション内のインデックスにより、InkStroke オブジェクトを取得します。 読み取り専用です。

#### 構文
```js
inkStrokeCollectionObject.getItem(index);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|index|number または string|InkStroke オブジェクトの ID、またはコレクション内の InkStroke オブジェクトのインデックス位置です。|

#### 戻り値
[InkStroke](inkstroke.md)

### getItemAt(index: number)
コレクション内での位置を基に InkStroke を取得します。

#### 構文
```js
inkStrokeCollectionObject.getItemAt(index);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|index|number|取得するオブジェクトのインデックス値。0 を起点とする番号になります。|

#### 戻り値
[InkStroke](inkstroke.md)

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
