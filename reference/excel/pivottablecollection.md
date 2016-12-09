# <a name="pivottablecollection-object-javascript-api-for-excel"></a>PivotTableCollection オブジェクト (JavaScript API for Excel)

ブックまたはワークシートの一部として含まれている、すべてのピボットテーブルのコレクションを表します。

## <a name="properties"></a>プロパティ

| プロパティ     | 型   |説明| 要件セット|
|:---------------|:--------|:----------|:----|
|items|[PivotTable[]](pivottable.md)|PivotTable オブジェクトのコレクション。読み取り専用です。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|

_プロパティのアクセスの[例を参照してください。](#property-access-examples)_

## <a name="relationships"></a>関係
なし


## <a name="methods"></a>メソッド

| メソッド           | 戻り値の型    |説明| 要件セット|
|:---------------|:--------|:----------|:----|
|[getItem(name: string)](#getitemname-string)|[PivotTable](pivottable.md)|名前を使用してピボットテーブルを取得します。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemOrNull(name: string)](#getitemornullname-string)|[PivotTable](pivottable.md)|名前を使用してピボットテーブルを取得します。ピボットテーブルが存在しない場合、戻りオブジェクトの isNull プロパティは true になります。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|[load(param: object)](#loadparam-object)|(非推奨)|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[refreshAll()](#refreshall)|void|コレクション内のすべてのピボットテーブルを更新します。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>メソッドの詳細


### <a name="getitemname-string"></a>getItem(name: string)
名前を使用してピボットテーブルを取得します。

#### <a name="syntax"></a>構文
```js
pivotTableCollectionObject.getItem(name);
```

#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|:---|
|name|string|取得するピボットテーブルの名前。|

#### <a name="returns"></a>戻り値
[PivotTable](pivottable.md)

### <a name="getitemornullname-string"></a>getItemOrNull(name: string)
名前を使用してピボットテーブルを取得します。ピボットテーブルが存在しない場合、戻りオブジェクトの isNull プロパティは true になります。

#### <a name="syntax"></a>構文
```js
pivotTableCollectionObject.getItemOrNull(name);
```

#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|:---|
|name|string|取得するピボットテーブルの名前。|

#### <a name="returns"></a>戻り値
[PivotTable](pivottable.md)

### <a name="loadparam-object"></a>load(param: object)
JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。

#### <a name="syntax"></a>構文
```js
object.load(param);
```

#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|:---|
|param|object|省略可能。パラメーターとリレーションシップ名を、区切られた文字列または 1 つの配列として受け入れます。あるいは、[loadOption](loadoption.md) オブジェクトを提供します。|

#### <a name="returns"></a>戻り値
void

### <a name="refreshall"></a>refreshAll()
コレクション内のすべてのピボットテーブルを更新します。

#### <a name="syntax"></a>構文
```js
pivotTableCollectionObject.refreshAll();
```

#### <a name="parameters"></a>パラメーター
なし

#### <a name="returns"></a>戻り値
void
