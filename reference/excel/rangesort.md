# <a name="rangesort-object-javascript-api-for-excel"></a>RangeSort オブジェクト (JavaScript API for Excel)

Range オブジェクトの並べ替え操作を管理します。

## <a name="properties"></a>プロパティ

なし

## <a name="relationships"></a>関係
なし


## <a name="methods"></a>メソッド

| メソッド           | 戻り値の型    |説明| 要件セット|
|:---------------|:--------|:----------|:----|
|[apply(fields:SortField[], matchCase: bool, hasHeaders: bool, orientation: string, method: string)](#applyfields-sortfield-matchcase-bool-hasheaders-bool-orientation-string-method-string)|void|並べ替え操作を実行します。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>メソッドの詳細


### <a name="applyfields-sortfield-matchcase-bool-hasheaders-bool-orientation-string-method-string"></a>apply(fields:SortField[], matchCase: bool, hasHeaders: bool, orientation: string, method: string)
並べ替え操作を実行します。

#### <a name="syntax"></a>構文
```js
rangeSortObject.apply(fields, matchCase, hasHeaders, orientation, method);
```

#### <a name="parameters"></a>パラメーター
| パラメーター       | 型    |説明|
|:---------------|:--------|:----------|:---|
|fields|SortField[]|並べ替えに使用する条件の一覧。|
|matchCase|bool|省略可能。大文字小文字の区別が文字列の順序に影響を与えるかどうか。|
|hasHeaders|bool|省略可能。範囲にヘッダーがあるかどうか。|
|orientation|string|省略可能。操作が行と列のどちらの並べ替えかを示します。使用可能な値は次のとおりです。Rows、Columns|
|method|string|省略可能。中国語文字に使用される順序付けの方法です。使用可能な値は次のとおりです。PinYin、StrokeCount|

#### <a name="returns"></a>戻り値
void
