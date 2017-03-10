# <a name="setting-object-javascript-api-for-excel"></a>Setting オブジェクト (JavaScript API for Excel)

Setting は、ドキュメントに永続化されている設定のキーと値のペアを表します。

## <a name="properties"></a>プロパティ

| プロパティ       | 型    |説明| 要件セット|
|:---------------|:--------|:----------|:----|
|Key|string|Setting の ID を表すキーを返します。読み取り専用です。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|value|object|この設定に格納されている値を表します。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|

_プロパティのアクセスの[例を参照してください。](#property-access-examples)_

## <a name="relationships"></a>関係
なし


## <a name="methods"></a>メソッド

| メソッド           | 戻り値の型    |説明| 要件セット|
|:---------------|:--------|:----------|:----|
|[delete()](#delete)|void|設定を削除します。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>メソッドの詳細


### <a name="delete"></a>delete()
設定を削除します。

#### <a name="syntax"></a>構文
```js
settingObject.delete();
```

#### <a name="parameters"></a>パラメーター
なし

#### <a name="returns"></a>戻り値
void
