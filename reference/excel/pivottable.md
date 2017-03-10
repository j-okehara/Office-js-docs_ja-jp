# <a name="pivottable-object-javascript-api-for-excel"></a>PivotTable オブジェクト (JavaScript API for Excel)

Excel のピボットテーブルを表します。

## <a name="properties"></a>プロパティ

| プロパティ       | 型    |説明| 要件セット|
|:---------------|:--------|:----------|:----|
|name|string|ピボットテーブルの名前。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|

_プロパティのアクセスの[例を参照してください。](#property-access-examples)_

## <a name="relationships"></a>関係
| リレーションシップ | 型    |説明| 要件セット|
|:---------------|:--------|:----------|:----|
|ワークシート|[Worksheet](worksheet.md)|現在のピボットテーブルを含んでいるワークシート。読み取り専用。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>メソッド

| メソッド           | 戻り値の型    |説明| 要件セット|
|:---------------|:--------|:----------|:----|
|[refresh()](#refresh)|void|ピボットテーブルを更新します。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>メソッドの詳細


### <a name="refresh"></a>refresh()
ピボットテーブルを更新します。

#### <a name="syntax"></a>構文
```js
pivotTableObject.refresh();
```

#### <a name="parameters"></a>パラメーター
なし

#### <a name="returns"></a>戻り値
void
