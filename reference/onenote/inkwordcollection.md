# InkWordCollection オブジェクト (JavaScript API for OneNote)

_適用対象:OneNote Online_  


InkWord オブジェクトのコレクションを表します。

## プロパティ

| プロパティ     | 型   |説明|フィードバック|
|:---------------|:--------|:----------|:-------|
|count|int|ページ内の InkWord の数を返します。 読み取り専用です。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkWordCollection-count)|
|items|[InkWord[]](inkword.md)|InkWord オブジェクトのコレクション。 読み取り専用です。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkWordCollection-items)|

_プロパティのアクセスの[例](#例)を参照してください。_

## リレーションシップ
なし


## メソッド

| メソッド           | 戻り値の型    |説明| フィードバック|
|:---------------|:--------|:----------|:-------|
|[getItem(index: number または string)](#getitemindex-number-または-string)|[InkWord](inkword.md)|ID かコレクション内のインデックスにより、InkWord オブジェクトを取得します。 読み取り専用です。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkWordCollection-getItem)|
|[getItemAt(index: number)](#getitematindex-number)|[InkWord](inkword.md)|コレクション内での位置を基に InkWord を取得します。|[実行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkWordCollection-getItemAt)|
|[load(param: object)](#loadparam-object)|void|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkWordCollection-load)|

## メソッドの詳細


### getItem(index: number または string)
ID かコレクション内のインデックスにより、InkWord オブジェクトを取得します。 読み取り専用です。

#### 構文
```js
inkWordCollectionObject.getItem(index);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|index|number または string|InkWord オブジェクトの ID、またはコレクション内の InkWord オブジェクトのインデックス位置です。|

#### 戻り値
[InkWord](inkword.md)

### getItemAt(index: number)
コレクション内での位置を基に InkWord を取得します。

#### 構文
```js
inkWordCollectionObject.getItemAt(index);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|index|number|取得するオブジェクトのインデックス値。0 を起点とする番号になります。|

#### 戻り値
[InkWord](inkword.md)

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
