# <a name="workbook-object-javascript-api-for-excel"></a>Workbook オブジェクト (JavaScript API for Excel)

Workbook は、ワークシート、テーブル、範囲などの関連するブック オブジェクトを含む最上位オブジェクトです。

## <a name="properties"></a>プロパティ

なし

## <a name="relationships"></a>関係
| リレーションシップ | 型    |説明| 要件セット|
|:---------------|:--------|:----------|:----|
|アプリケーション|[Application](application.md)|このブックを含む Excel アプリケーションのインスタンスを表します。読み取り専用。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|bindings|[BindingCollection](bindingcollection.md)|ブックの一部であるバインドのコレクションを表します。読み取り専用。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|functions|[Functions](functions.md)|このブックを含む Excel アプリケーションのインスタンスを表します。読み取り専用。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|names|[NamedItemCollection](nameditemcollection.md)|ブック スコープの名前付き項目 (名前付き範囲と名前付き定数) のコレクションを表します。読み取り専用。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|pivotTables|[PivotTableCollection](pivottablecollection.md)|ブックに関連付けられているピボットテーブルのコレクションを表します。読み取り専用。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|settings|[SettingCollection](settingcollection.md)|ブックに関連付けられている Setting のコレクションを表します。読み取り専用です。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|tables|[TableCollection](tablecollection.md)|ブックに関連付けられているテーブルのコレクションを表します。読み取り専用。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|worksheets|[WorksheetCollection](worksheetcollection.md)|ブックに関連付けられているワークシートのコレクションを表します。読み取り専用。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>メソッド

| メソッド           | 戻り値の型    |説明| 要件セット|
|:---------------|:--------|:----------|:----|
|[getSelectedRange()](#getselectedrange)|[Range](range.md)|ブックから現在選択されている範囲を取得します。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>メソッドの詳細


### <a name="getselectedrange"></a>getSelectedRange()
ブックから現在選択されている範囲を取得します。

#### <a name="syntax"></a>構文
```js
workbookObject.getSelectedRange();
```

#### <a name="parameters"></a>パラメーター
なし

#### <a name="returns"></a>戻り値
[Range](range.md)

#### <a name="examples"></a>例

```js
Excel.run(function (ctx) { 
    var selectedRange = ctx.workbook.getSelectedRange();
    selectedRange.load('address');
    return ctx.sync().then(function() {
            console.log(selectedRange.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```