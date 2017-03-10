# <a name="rangeviewcollection-object-javascript-api-for-excel"></a>RangeViewCollection オブジェクト (JavaScript API for Excel)

RangeView オブジェクトのコレクションを表します。

## <a name="properties"></a>プロパティ

| プロパティ       | 型    |説明| 要件セット|
|:---------------|:--------|:----------|:----|
|items|[RangeView[]](rangeview.md)|rangeView オブジェクトのコレクション。読み取り専用。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|

_プロパティのアクセスの[例を参照してください。](#property-access-examples)_

## <a name="relationships"></a>関係
なし


## <a name="methods"></a>メソッド

| メソッド           | 戻り値の型    |説明| 要件セット|
|:---------------|:--------|:----------|:----|
|[getCount()](#getcount)|int|コレクション内にある RangeView オブジェクトの数を取得します。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemAt(index: number)](#getitematindex-number)|[RangeView](rangeview.md)|RangeView のインデックスから RangeView の行番号を取得します。0 を起点とする番号になります。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>メソッドの詳細


### <a name="getcount"></a>getCount()
コレクション内にある RangeView オブジェクトの数を取得します。

#### <a name="syntax"></a>構文
```js
rangeViewCollectionObject.getCount();
```

#### <a name="parameters"></a>パラメーター
なし

#### <a name="returns"></a>戻り値
int

### <a name="getitematindex-number"></a>getItemAt(index: number)
RangeView のインデックスから RangeView の行番号を取得します。0 を起点とする番号になります。

#### <a name="syntax"></a>構文
```js
rangeViewCollectionObject.getItemAt(index);
```

#### <a name="parameters"></a>パラメーター
| パラメーター       | 型    |説明|
|:---------------|:--------|:----------|:---|
|index|number|表示されている行のインデックス。|

#### <a name="returns"></a>戻り値
[RangeView](rangeview.md)
