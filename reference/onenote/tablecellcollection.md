# TableCellCollection オブジェクト (JavaScript API for OneNote)

_適用対象:OneNote Online_  


TableCell オブジェクトのコレクションが含まれています。

## プロパティ

| プロパティ     | 型   |説明|フィードバック|
|:---------------|:--------|:----------|:-------|
|count|int|このコレクション内の tablecells の数を返します。 読み取り専用です。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCellCollection-count)|
|Items|[TableCell[]](tablecell.md)|TableCell オブジェクトのコレクション。 読み取り専用です。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCellCollection-items)|

_プロパティのアクセスの[例](#例)を参照してください。_

## リレーションシップ
なし


## メソッド

| メソッド           | 戻り値の型    |説明| フィードバック|
|:---------------|:--------|:----------|:-------|
|[getItem(index: number または string)](#getitemindex-number-または-string)|[TableCell](tablecell.md)|ID やコレクション内のインデックスで、テーブル セル オブジェクトを取得します。 読み取り専用です。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCellCollection-getItem)|
|[getItemAt(index: number)](#getitematindex-number)|[TableCell](tablecell.md)|コレクション内の位置に基づいてテーブル セルを取得します。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCellCollection-getItemAt)|
|[load(param: object)](#loadparam-object)|void|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCellCollection-load)|

## メソッドの詳細


### getItem(index: number or string)
ID やコレクション内のインデックスで、テーブル セル オブジェクトを取得します。 読み取り専用です。

#### 構文
```js
tableCellCollectionObject.getItem(index);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|index|number または string|テーブル セル オブジェクトのインデックス位置を識別する番号です。|

#### 戻り値
[TableCell](tablecell.md)

### getItemAt(index: number)
コレクション内の位置に基づいてテーブル セルを取得します。

#### 構文
```js
tableCellCollectionObject.getItemAt(index);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|index|number|取得するオブジェクトのインデックス値。0 を起点とする番号になります。|

#### 戻り値
[TableCell](tablecell.md)

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
