# InkAnalysisLineCollection オブジェクト (JavaScript API for OneNote)

_適用対象:OneNote Online_  


InkAnalysisLine オブジェクトのコレクションを表します。

## プロパティ

| プロパティ     | 型   |説明|フィードバック|
|:---------------|:--------|:----------|:-------|
|count|int|ページ内の InkAnalysisLine の数を返します。 読み取り専用です。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisLineCollection-count)|
|items|[InkAnalysisLine[]](inkanalysisline.md)|InkAnalysisLine オブジェクトのコレクション。 読み取り専用です。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisLineCollection-items)|

_プロパティのアクセスの[例](#例)を参照してください。_

## リレーションシップ
なし


## メソッド

| メソッド           | 戻り値の型    |説明| フィードバック|
|:---------------|:--------|:----------|:-------|
|[getItem(index: number または string)](#getitemindex-number-または-string)|[InkAnalysisLine](inkanalysisline.md)|ID かコレクション内のインデックスにより、InkAnalysisLine オブジェクトを取得します。 読み取り専用です。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisLineCollection-getItem)|
|[getItemAt(index: number)](#getitematindex-number)|[InkAnalysisLine](inkanalysisline.md)|コレクション内での位置を基に InkAnalysisLine を取得します。|[実行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisLineCollection-getItemAt)|
|[load(param: object)](#loadparam-object)|void|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisLineCollection-load)|

## メソッドの詳細


### getItem(index: number または string)
ID かコレクション内のインデックスにより、InkAnalysisLine オブジェクトを取得します。 読み取り専用です。

#### 構文
```js
inkAnalysisLineCollectionObject.getItem(index);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|index|number または string|InkAnalysisLine オブジェクトの ID、またはコレクション内の InkAnalysisLine オブジェクトのインデックス位置です。|

#### 戻り値
[InkAnalysisLine](inkanalysisline.md)

### getItemAt(index: number)
コレクション内での位置を基に InkAnalysisLine を取得します。

#### 構文
```js
inkAnalysisLineCollectionObject.getItemAt(index);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|index|number|取得するオブジェクトのインデックス値。0 を起点とする番号になります。|

#### 戻り値
[InkAnalysisLine](inkanalysisline.md)

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
