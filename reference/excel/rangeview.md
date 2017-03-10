# <a name="rangeview-object-javascript-api-for-excel"></a>RangeView オブジェクト (JavaScript API for Excel)

RangeView は、親の範囲の表示されているセルのセットを表します。

## <a name="properties"></a>プロパティ

| プロパティ       | 型    |説明| 要件セット|
|:---------------|:--------|:----------|:----|
|cellAddresses|object[][]|RangeView のセル アドレスを表します。読み取り専用。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|columnCount|int|表示されている列の数を返します。読み取り専用。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|formulas|object[][]|A1 スタイル表記の数式を表します。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|formulasLocal|object[][]|ユーザーの言語と数値書式ロケールで、A1 スタイル表記の数式を表します。たとえば、英語の数式 "=SUM(A1, 1.5)" は、ドイツ語では "=SUMME(A1; 1,5)" になります。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|formulasR1C1|object[][]|R1C1 スタイル表記の数式を表します。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|index|int|RangeView のインデックスを表す値を返します。読み取り専用。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|numberFormat|オブジェクト型 (Object)|指定したセルの Excel の数値書式コードを表します。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|rowCount|int|表示されている行の数を返します。読み取り専用。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|text|object[][]|指定した範囲のテキスト値。テキスト値は、セルの幅には依存しません。Excel UI で発生する # 記号による置換は、この API から返されるテキスト値には影響しません。読み取り専用です。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|valueTypes|string|各セルのデータの種類を表します。読み取り専用です。使用可能な値は次のとおりです。Unknown、Empty、String、Integer、Double、Boolean、Error。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|values|object[][]|指定した範囲ビューの Raw 値を表します。返されるデータの型は、文字列、数値、ブール値のいずれかになります。エラーが含まれているセルは、エラー文字列を返します。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|

_プロパティのアクセスの[例を参照してください。](#property-access-examples)_

## <a name="relationships"></a>関係
| リレーションシップ | 型    |説明| 要件セット|
|:---------------|:--------|:----------|:----|
|rows|[RangeViewCollection](rangeviewcollection.md)|範囲に関連付けられている範囲ビューのコレクションを表します。読み取り専用。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>メソッド

| メソッド           | 戻り値の型    |説明| 要件セット|
|:---------------|:--------|:----------|:----|
|[getRange()](#getrange)|[Range](range.md)|現在の RangeView に関連付けられている親の範囲を取得します。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>メソッドの詳細


### <a name="getrange"></a>getRange()
現在の RangeView に関連付けられている親の範囲を取得します。

#### <a name="syntax"></a>構文
```js
rangeViewObject.getRange();
```

#### <a name="parameters"></a>パラメーター
なし

#### <a name="returns"></a>戻り値
[Range](range.md)
