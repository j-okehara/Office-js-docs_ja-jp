# <a name="pivottablecollection-object-javascript-api-for-excel"></a>PivotTableCollection オブジェクト (JavaScript API for Excel)

ブックまたはワークシートの一部として含まれている、すべてのピボットテーブルのコレクションを表します。

## <a name="properties"></a>プロパティ

| プロパティ       | 型    |説明| 要件セット|
|:---------------|:--------|:----------|:----|
|items|[PivotTable[]](pivottable.md)|ピボットテーブル オブジェクトのコレクション。読み取り専用。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|

_プロパティのアクセスの[例を参照してください。](#property-access-examples)_

## <a name="relationships"></a>関係
なし


## <a name="methods"></a>メソッド

| メソッド           | 戻り値の型    |説明| 要件セット|
|:---------------|:--------|:----------|:----|
|[getCount()](#getcount)|int|コレクション内のピボット テーブルの数を取得します。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getItem(name: string)](#getitemname-string)|[PivotTable](pivottable.md)|名前を使用してピボットテーブルを取得します。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemOrNullObject(name: string)](#getitemornullobjectname-string)|[PivotTable](pivottable.md)|名前を使用してピボットテーブルを取得します。PivotTable が存在しない場合は null オブジェクトを返します。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[refreshAll()](#refreshall)|void|コレクション内のすべてのピボットテーブルを更新します。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>メソッドの詳細


### <a name="getcount"></a>getCount()
コレクション内のピボット テーブルの数を取得します。

#### <a name="syntax"></a>構文
```js
pivotTableCollectionObject.getCount();
```

#### <a name="parameters"></a>パラメーター
なし

#### <a name="returns"></a>戻り値
int

### <a name="getitemname-string"></a>getItem(name: string)
名前を使用してピボットテーブルを取得します。

#### <a name="syntax"></a>構文
```js
pivotTableCollectionObject.getItem(name);
```

#### <a name="parameters"></a>パラメーター
| パラメーター       | 型    |説明|
|:---------------|:--------|:----------|:---|
|name|string|取得するピボットテーブルの名前。|

#### <a name="returns"></a>戻り値
[PivotTable](pivottable.md)

### <a name="getitemornullobjectname-string"></a>getItemOrNullObject(name: string)
名前を使用してピボットテーブルを取得します。PivotTable が存在しない場合は null オブジェクトを返します。

#### <a name="syntax"></a>構文
```js
pivotTableCollectionObject.getItemOrNullObject(name);
```

#### <a name="parameters"></a>パラメーター
| パラメーター       | 型    |説明|
|:---------------|:--------|:----------|:---|
|name|string|取得するピボットテーブルの名前。|

#### <a name="returns"></a>戻り値
[PivotTable](pivottable.md)

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
