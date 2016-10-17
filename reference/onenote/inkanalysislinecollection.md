# <a name="inkanalysislinecollection-object-(javascript-api-for-onenote)"></a>InkAnalysisLineCollection オブジェクト (JavaScript API for OneNote)

_適用対象:OneNote Online_  


InkAnalysisLine オブジェクトのコレクションを表します。

## <a name="properties"></a>プロパティ

| プロパティ     | 型   |説明|フィードバック|
|:---------------|:--------|:----------|:-------|
|count|int|ページ内の InkAnalysisLine の数を返します。読み取り専用です。|[移動](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisLineCollection-count)|
|items|[InkAnalysisLine[]](inkanalysisline.md)|InkAnalysisLine オブジェクトのコレクション。読み取り専用です。|[移動](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisLineCollection-items)|

_プロパティのアクセスの[例を参照してください。](#property-access-examples)_

## <a name="relationships"></a>関係
なし


## <a name="methods"></a>メソッド

| メソッド           | 戻り値の型    |説明| フィードバック|
|:---------------|:--------|:----------|:-------|
|[getItem(index: number or string)](#getitemindex-number-or-string)|[InkAnalysisLine](inkanalysisline.md)|ID かコレクション内のインデックスにより、InkAnalysisLine オブジェクトを取得します。読み取り専用です。|[移動](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisLineCollection-getItem)|
|[getItemAt(index: number)](#getitematindex-number)|[InkAnalysisLine](inkanalysisline.md)|コレクション内での位置を基に InkAnalysisLine を取得します。|[移動](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisLineCollection-getItemAt)|
|[load(param: object)](#loadparam-object)|(非推奨)|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|[移動](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisLineCollection-load)|

## <a name="method-details"></a>メソッドの詳細


### <a name="getitem(index:-number-or-string)"></a>getItem(index: number または string)
ID かコレクション内のインデックスにより、InkAnalysisLine オブジェクトを取得します。読み取り専用です。

#### <a name="syntax"></a>構文
```js
inkAnalysisLineCollectionObject.getItem(index);
```

#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|index|number または string|InkAnalysisLine オブジェクトの ID、またはコレクション内の InkAnalysisLine オブジェクトのインデックス位置です。|

#### <a name="returns"></a>戻り値
[InkAnalysisLine](inkanalysisline.md)

### <a name="getitemat(index:-number)"></a>getItemAt(index: number)
コレクション内での位置を基に InkAnalysisLine を取得します。

#### <a name="syntax"></a>構文
```js
inkAnalysisLineCollectionObject.getItemAt(index);
```

#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|index|number|取得するオブジェクトのインデックス値。0 を起点とする番号になります。|

#### <a name="returns"></a>戻り値
[InkAnalysisLine](inkanalysisline.md)

### <a name="load(param:-object)"></a>load(param: object)
JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。

#### <a name="syntax"></a>構文
```js
object.load(param);
```

#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|param|object|省略可能。パラメーターとリレーションシップ名を、区切られた文字列または 1 つの配列として受け入れます。あるいは、[loadOption](loadoption.md) オブジェクトを提供します。|

#### <a name="returns"></a>戻り値
void
