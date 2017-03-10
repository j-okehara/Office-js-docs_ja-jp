# <a name="filterdatetime-object-javascript-api-for-excel"></a>FilterDatetime オブジェクト (JavaScript API for Excel)

値をフィルター処理するときに日付をフィルター処理する方法を表します。

## <a name="properties"></a>プロパティ

| プロパティ       | 型    |説明| 要件セット|
|:---------------|:--------|:----------|:----|
|日付|string|データのフィルター処理に使用する ISO8601 形式の日付です。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|specificity|string|データを保持するのに、日付をどの程度詳細に使用するか。たとえば、date が 2005-04-02 で "month" に設定した場合、フィルター操作では 2005 年 4 月の日付データを含むすべての行が保持されます。使用可能な値は次のとおりです。Year、Month、Day、Hour、Minute、Second。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

_プロパティのアクセスの[例を参照してください。](#property-access-examples)_

## <a name="relationships"></a>関係
なし


## <a name="methods"></a>メソッド
なし

