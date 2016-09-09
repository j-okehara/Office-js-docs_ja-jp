# TableRowCollection オブジェクト (JavaScript API for OneNote)

_適用対象:OneNote Online_  


TableRow オブジェクトのコレクションが含まれます。

## プロパティ

| プロパティ     | 型   |説明|フィードバック|
|:---------------|:--------|:----------|:-------|
|count|int|このコレクション内のテーブル行数を返します。 読み取り専用です。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableRowCollection-count)|
|Items|[TableRow[]](tablerow.md)|tableRow オブジェクトのコレクション。 読み取り専用です。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableRowCollection-items)|

_プロパティのアクセスの[例](#例)を参照してください。_

## リレーションシップ
なし


## メソッド

| メソッド           | 戻り値の型    |説明| フィードバック|
|:---------------|:--------|:----------|:-------|
|[getItem(index: number or string)](#getitemindex-number-or-string)|[TableRow](tablerow.md)|ID やコレクション内のインデックスで、テーブル行オブジェクトを取得します。 読み取り専用です。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableRowCollection-getItem)|
|[getItemAt(index: number)](#getitematindex-number)|[TableRow](tablerow.md)|コレクション内の位置に基づいてテーブル行を取得します。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableRowCollection-getItemAt)|
|[load(param: object)](#loadparam-object)|void|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableRowCollection-load)|

## メソッドの詳細


### getItem(index: number or string)
ID やコレクション内のインデックスで、テーブル行オブジェクトを取得します。 読み取り専用です。

#### 構文
```js
tableRowCollectionObject.getItem(index);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|index|number または string|テーブル行オブジェクトのインデックス位置を識別する番号です。|

#### 戻り値
[TableRow](tablerow.md)

### getItemAt(index: number)
コレクション内の位置に基づいてテーブル行を取得します。

#### 構文
```js
tableRowCollectionObject.getItemAt(index);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|index|number|取得するオブジェクトのインデックス値。0 を起点とする番号になります。|

#### 戻り値
[TableRow](tablerow.md)

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
