# <a name="application-object-javascript-api-for-excel"></a>Application オブジェクト (JavaScript API for Excel)

ブックを管理する Excel アプリケーションを表します。

## <a name="properties"></a>プロパティ

| プロパティ     | 型   |説明|要件セット|
|:---------------|:--------|:----------|:----------|
|calculationMode|string|ブックで使用される計算モードを返します。値の取得のみ可能です。使用可能な値は次のとおりです。`Automatic` Excel が再計算を制御します。`AutomaticExceptTables` Excel が再計算を制御しますが、テーブル内の変更は無視します。`Manual` 計算は、ユーザーが要求したときに行われます。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_プロパティのアクセスの[例を参照してください。](#property-access-examples)_

## <a name="relationships"></a>関係
なし


## <a name="methods"></a>メソッド

| メソッド           | 戻り値の型    |説明|要件セット|
|:---------------|:--------|:----------|:----------|
|[calculate(calculationType: string)](#calculatecalculationtype-string)|void|Excel で現在開いているすべてのブックを再計算します。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[load(param: object)](#loadparam-object)|(非推奨)|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>メソッドの詳細


### <a name="calculatecalculationtype-string"></a>calculate(calculationType: string)
Excel で現在開いているすべてのブックを再計算します。

#### <a name="syntax"></a>構文
```js
applicationObject.calculate(calculationType);
```

#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|calculationType|文字列|使用する計算の種類を指定します。使用可能な値は次のとおりです。`Recalculate` これはソフトの再計算であり、主に下位互換性のためのものです。`Full` Excel によってダーティのマークが付けられたすべてのセル (揮発性データと変更されたデータの参照先、およびプログラムによりダーティのマークが付けられたセル) を再計算します。`FullRebuild` 開いているすべてのブックに含まれるすべてのセルを再計算します。|

#### <a name="returns"></a>戻り値
void

#### <a name="examples"></a>例
```js
Excel.run(function (ctx) { 
    ctx.workbook.application.calculate('Full');
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="loadparam-object"></a>load(param: object)
JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。

#### <a name="syntax"></a>構文
```js
object.load(param);
```

#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|param|object|省略可能。パラメーター名とリレーションシップ名を、区切られた文字列または 1 つの配列として受け入れます。あるいは、[loadOption](loadoption.md) オブジェクトを受け入れます。|

#### <a name="returns"></a>戻り値
void
### <a name="property-access-examples"></a>プロパティのアクセスの例
```js
Excel.run(function (ctx) { 
    var application = ctx.workbook.application;
    application.load('calculationMode');
    return ctx.sync().then(function() {
        console.log(application.calculationMode);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
