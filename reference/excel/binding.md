# <a name="binding-object-(javascript-api-for-excel)"></a>Binding オブジェクト (JavaScript API for Excel)

ブックで定義されている Office.js のバインディングを表します。

## <a name="properties"></a>プロパティ

| プロパティ     | 型   |説明
|:---------------|:--------|:----------|
|id|string|バインド識別子を表します。読み取り専用。|
|type|string|バインドの型を返します。読み取り専用。使用可能な値は次のとおりです。Range, Table, Text。|

_プロパティのアクセスの[例](#property-access-examples)を参照してください。_

## <a name="relationships"></a>関係
なし


## <a name="methods"></a>メソッド

| メソッド           | 戻り値の型    |説明|
|:---------------|:--------|:----------|
|[getRange()](#getrange)|[Range](range.md)|バインディングによって表される範囲を返します。バインドが正しい型ではない場合、エラーがスローされます。|
|[getTable()](#gettable)|[Table](table.md)|バインドによって表されるテーブルを返します。バインドが正しい型ではない場合、エラーがスローされます。|
|[getText()](#gettext)|string|バインドによって表されるテキストを返します。バインドが正しい型ではない場合、エラーがスローされます。|
|[load(param: object)](#loadparam-object)|void|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|

## <a name="method-details"></a>メソッドの詳細


### <a name="getrange()"></a>getRange()
バインディングによって表される範囲を返します。バインドが正しい型ではない場合、エラーがスローされます。

#### <a name="syntax"></a>構文
```js
bindingObject.getRange();
```

#### <a name="parameters"></a>パラメーター
なし

#### <a name="returns"></a>戻り値
[Range](range.md)

#### <a name="examples"></a>例
以下の例では、バインド オブジェクトを使用して、関連付けられている範囲を取得しています。

```js
Excel.run(function (ctx) { 
    var binding = ctx.workbook.bindings.getItemAt(0);
    var range = binding.getRange();
    range.load('cellCount');
    return ctx.sync().then(function() {
        console.log(range.cellCount);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="gettable()"></a>getTable()
バインドによって表されるテーブルを返します。バインドが正しい型ではない場合、エラーがスローされます。

#### <a name="syntax"></a>構文
```js
bindingObject.getTable();
```

#### <a name="parameters"></a>パラメーター
なし

#### <a name="returns"></a>戻り値
[Table](table.md)

#### <a name="examples"></a>例
```js
Excel.run(function (ctx) { 
    var binding = ctx.workbook.bindings.getItemAt(0);
    var table = binding.getTable();
    table.load('name');
    return ctx.sync().then(function() {
            console.log(table.name);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="gettext()"></a>getText()
バインドによって表されるテキストを返します。バインドが正しい型ではない場合、エラーがスローされます。

#### <a name="syntax"></a>構文
```js
bindingObject.getText();
```

#### <a name="parameters"></a>パラメーター
なし

#### <a name="returns"></a>戻り値
string

#### <a name="examples"></a>例

```js
Excel.run(function (ctx) { 
    var binding = ctx.workbook.bindings.getItemAt(0);
    var text = binding.getText();
    ctx.load('text');
    return ctx.sync().then(function() {
        console.log(text);
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
|param|object|省略可能。パラメーター名とリレーションシップ名を、区切られた文字列または 1 つの配列として受け入れます。あるいは、[loadOption](loadoption.md) オブジェクトを受け入れます。|

#### <a name="returns"></a>戻り値
void
### <a name="property-access-examples"></a>プロパティのアクセスの例

```js
Excel.run(function (ctx) { 
    var binding = ctx.workbook.bindings.getItemAt(0);
    binding.load('type');
    return ctx.sync().then(function() {
        console.log(binding.type);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
