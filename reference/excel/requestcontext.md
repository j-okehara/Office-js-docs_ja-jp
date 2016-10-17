# <a name="requestcontext-object-(javascript-api-for-excel)"></a>RequestContext オブジェクト (JavaScript API for Excel)

RequestContext オブジェクトは、Excel アプリケーションへの要求を容易にします。Office アドインと Excel アプリケーションは 2 つの異なるプロセスで実行されているため、アドインから Excel とその関連オブジェクト (ワークシートや表など) にアクセスするには要求のコンテキストが必要です。 

## <a name="properties"></a>プロパティ
なし

## <a name="methods"></a>メソッド

| メソッド         | 戻り値の型    |説明|
|:---------------|:--------|:----------|
|[load(object: object, option: object)](#loadobject-object-option-object)  |void     |JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオプションを設定します。|

## <a name="api-specification"></a>API 仕様

### <a name="load(object:-object,-option:-object)"></a>load(object: object, option: object)
JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオプションを設定します。

#### <a name="syntax"></a>構文
```js
requestContextObject.load(object, loadOption);
```

#### <a name="parameters"></a>パラメーター
| パラメーター       | 型    |説明|
|:----------------|:--------|:----------|
|object|object|省略可能。読み込むオブジェクトの名前を指定します。|
|option|[loadOption](loadoption.md)|省略可能。select、expand、skip、top などの読み込みオプションを指定します。詳細については、loadOption オブジェクトを参照してください。|

#### <a name="returns"></a>戻り値
void

##### <a name="examples"></a>例

次の例では、1 つの範囲からプロパティ値を読み込んで、それらを別の範囲にコピーしています。

```js
Excel.run(function (ctx) { 
    var range = ctx.workbook.worksheets.getActiveWorksheet().getRange("A1:A2");
    ctx.load(range, "values");
    return ctx.sync().then(function() {
        var myvalues=range.values;
        ctx.workbook.worksheets.getActiveWorksheet().getRange("B1:B2").values = myvalues;
        console.log(range.values);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
})
```
