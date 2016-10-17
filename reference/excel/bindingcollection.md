# <a name="bindingcollection-object-(javascript-api-for-excel)"></a>BindingCollection オブジェクト (JavaScript API for Excel)

ブックの一部であるすべてのバインド オブジェクトのコレクションを表します。

## <a name="properties"></a>プロパティ

| プロパティ     | 型   |説明
|:---------------|:--------|:----------|
|count|int|コレクション内にあるバインドの数を取得します。値の取得のみ可能です。|
|items|[Binding[]](binding.md)|バインド オブジェクトのコレクション。読み取り専用です。|

_プロパティのアクセスの[例](#property-access-examples)を参照してください。_

## <a name="relationships"></a>関係
なし


## <a name="methods"></a>メソッド

| メソッド           | 戻り値の型    |説明|
|:---------------|:--------|:----------|
|[getItem(id: string)](#getitemid-string)|[Binding](binding.md)|ID によってバインド オブジェクトを取得します。|
|[getItemAt(index: number)](#getitematindex-number)|[Binding](binding.md)|項目の配列内の位置に基づいて、バインド オブジェクトを取得します。|
|[load(param: object)](#loadparam-object)|void|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|

## <a name="method-details"></a>メソッドの詳細


### <a name="getitem(id:-string)"></a>getItem(id: string)
ID によってバインド オブジェクトを取得します。

#### <a name="syntax"></a>構文
```js
bindingCollectionObject.getItem(id);
```

#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|id|string|取得するバインド オブジェクトの ID。|

#### <a name="returns"></a>戻り値
[Binding](binding.md)

#### <a name="examples"></a>例

表のデータ変更をモニターするための表バインドを作成します。データが変更されると、表の背景色がオレンジ色に変更されます。

```js
(function () {
    // Create myTable
    Excel.run(function (ctx) {
        var table = ctx.workbook.tables.add("Sheet1!A1:C4", true);
        table.name = "myTable";
        return ctx.sync().then(function () {
            console.log("MyTable is Created!");

            //Create a new table binding for myTable
            Office.context.document.bindings.addFromNamedItemAsync("myTable", Office.CoercionType.Table, { id: "myBinding" }, function (asyncResult) {
                if (asyncResult.status == "failed") {
                    console.log("Action failed with error: " + asyncResult.error.message);
                }
                else {
                    // If successful, add the event handler to the table binding.
                    Office.select("bindings#myBinding").addHandlerAsync(Office.EventType.BindingDataChanged, onBindingDataChanged);
                }
            });
        })
        .catch(function (error) {
            console.log(JSON.stringify(error));
        });
    });
    
    // When data in the table is changed, this event is triggered.
    function onBindingDataChanged(eventArgs) {
        Excel.run(function (ctx) {
            // Highlight the table in orange to indicate data changed.
            var fill = ctx.workbook.tables.getItem("myTable").getDataBodyRange().format.fill;
            fill.load("color");
            return ctx.sync().then(function () {
                if (fill.color != "Orange") {
                    ctx.workbook.bindings.getItem(eventArgs.binding.id).getTable().getDataBodyRange().format.fill.color = "Orange";
 
                    console.log("The value in this table got changed!");
                }
                else
                    
            })
                .then(ctx.sync)
            .catch(function (error) {
                console.log(JSON.stringify(error));
            });
        });
    } 
})();
 


```



#### <a name="examples"></a>例
```js
Excel.run(function (ctx) { 
    var lastPosition = ctx.workbook.bindings.count - 1;
    var binding = ctx.workbook.bindings.getItemAt(lastPosition);
    binding.load('type')
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


### <a name="getitemat(index:-number)"></a>getItemAt(index: number)
項目の配列内の位置に基づいて、バインド オブジェクトを取得します。

#### <a name="syntax"></a>構文
```js
bindingCollectionObject.getItemAt(index);
```

#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|index|number|取得するオブジェクトのインデックス値。0 を起点とする番号になります。|

#### <a name="returns"></a>戻り値
[Binding](binding.md)

#### <a name="examples"></a>例
```js
Excel.run(function (ctx) { 
    var lastPosition = ctx.workbook.bindings.count - 1;
    var binding = ctx.workbook.bindings.getItemAt(lastPosition);
    binding.load('type')
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
    var bindings = ctx.workbook.bindings;
    bindings.load('items');
    return ctx.sync().then(function() {
        for (var i = 0; i < bindings.items.length; i++)
        {
            console.log(bindings.items[i].id);
            console.log(bindings.items[i].index);
        }
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
バインドの数を取得します。

```js
Excel.run(function (ctx) { 
    var bindings = ctx.workbook.bindings;
    bindings.load('count');
    return ctx.sync().then(function() {
        console.log("Bindings: Count= " + bindings.count);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
