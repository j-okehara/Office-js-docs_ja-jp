# <a name="tablesort-object-javascript-api-for-excel"></a>TableSort オブジェクト (JavaScript API for Excel)

Table オブジェクトの並べ替え操作を管理します。

## <a name="properties"></a>プロパティ

| プロパティ     | 型   |説明| 要件セット|
|:---------------|:--------|:----------|:----|
|matchCase|bool|大文字小文字の区別が、テーブルの最後の並べ替え操作に影響を与えたかどうかを表します。読み取り専用です。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|method|string|テーブルの並べ替えで最後に使用した中国語文字の順序付け方法を表します。読み取り専用です。使用可能な値は次のとおりです。PinYin、StrokeCount。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="relationships"></a>リレーションシップ
| リレーションシップ | 型   |説明| 要件セット|
|:---------------|:--------|:----------|:----|
|fields|[SortField](sortfield.md)|テーブルの最後の並べ替えに使用する現在の条件を表します。読み取り専用です。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>メソッド

| メソッド           | 戻り値の型    |説明| 要件セット|
|:---------------|:--------|:----------|:----|
|[apply(fields:SortField[], matchCase: bool, method: string)](#applyfields-sortfield-matchcase-bool-method-string)|void|並べ替え操作を実行します。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[clear()](#clear)|void|テーブルに現在設定されている並べ替えをクリアします。これにより表の順序が変更されることはありませんが、ヘッダーのボタンの状態がクリアされます。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[load(param: object)](#loadparam-object)|(非推奨)|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[reapply()](#reapply)|void|テーブルに、現在の並べ替えパラメーターを再適用します。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>メソッドの詳細


### <a name="applyfields-sortfield-matchcase-bool-method-string"></a>apply(fields:SortField[], matchCase: bool, method: string)
並べ替え操作を実行します。

#### <a name="syntax"></a>構文
```js
tableSortObject.apply(fields, matchCase, method);
```

#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|:---|
|fields|SortField[]|並べ替えに使用する条件の一覧。|
|matchCase|bool|省略可能。大文字小文字の区別が文字列の順序に影響を与えるかどうか。|
|method|string|省略可能。中国語文字に使用される順序付けの方法です。使用可能な値は次のとおりです。PinYin、StrokeCount|

#### <a name="returns"></a>戻り値
void

#### <a name="examples"></a>例
```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    table.sort.apply([ 
            {
                key: 2,
                ascending: true
            },
        ], true);
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### <a name="clear"></a>clear()
テーブルに現在設定されている並べ替えをクリアします。これにより表の順序が変更されることはありませんが、ヘッダーのボタンの状態がクリアされます。

#### <a name="syntax"></a>構文
```js
tableSortObject.clear();
```

#### <a name="parameters"></a>パラメーター
なし

#### <a name="returns"></a>戻り値
void

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

### <a name="reapply"></a>reapply()
テーブルに、現在の並べ替えパラメーターを再適用します。

#### <a name="syntax"></a>構文
```js
tableSortObject.reapply();
```

#### <a name="parameters"></a>パラメーター
なし

#### <a name="returns"></a>戻り値
void
