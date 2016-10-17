# <a name="workbook-object-(javascript-api-for-excel)"></a>Workbook オブジェクト (JavaScript API for Excel)

ブックは、ワークシート、表、範囲などの関連するブック オブジェクトを含む最上位オブジェクトです。

## <a name="properties"></a>プロパティ

なし

## <a name="relationships"></a>関係
| リレーションシップ | 型   |説明|
|:---------------|:--------|:----------|
|application|[Application](application.md)|このブックを含む Excel アプリケーションのインスタンスを表します。読み取り専用です。|
|bindings|[BindingCollection](bindingcollection.md)|ブックの一部であるバインドのコレクションを表します。読み取り専用です。|
|functions|[Functions](functions.md)|このブックを含む Excel アプリケーションのインスタンスを表します。読み取り専用です。|
|names|[NamedItemCollection](nameditemcollection.md)|ブック スコープの名前付き項目 (名前付き範囲と名前付き定数) のコレクションを表します。読み取り専用。|
|tables|[TableCollection](tablecollection.md)|ブックに関連付けられているテーブルのコレクションを表します。読み取り専用。|
|worksheets|[WorksheetCollection](worksheetcollection.md)|ブックに関連付けられているワークシートのコレクションを表します。読み取り専用。|

## <a name="methods"></a>メソッド

| メソッド           | 戻り値の型    |説明|
|:---------------|:--------|:----------|
|[getSelectedRange()](#getselectedrange)|[Range](range.md)|ブックから現在選択されている範囲を取得します。|
|[load(param: object)](#loadparam-object)|void|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|

## <a name="method-details"></a>メソッドの詳細


### <a name="getselectedrange()"></a>getSelectedRange()
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
### <a name="load(param:-object)"></a>load(param: object)
JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。

#### <a name="syntax"></a>構文
```js
object.load(param);
```

#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|param|object|省略可能。パラメーター名とリレーションシップ名を、区切られた文字列または 1 つの配列として受け入れます。あるいは、[loadOption](loadoption.md) オブジェクトを提供します。|

#### <a name="returns"></a>戻り値
void
