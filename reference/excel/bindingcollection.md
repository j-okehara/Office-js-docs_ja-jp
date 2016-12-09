# <a name="bindingcollection-object-javascript-api-for-excel"></a>BindingCollection オブジェクト (JavaScript API for Excel)

ブックの一部であるすべてのバインド オブジェクトのコレクションを表します。

## <a name="properties"></a>プロパティ

| プロパティ     | 型   |説明| 要件セット|
|:---------------|:--------|:----------|:----|
|count|int|コレクション内にあるバインドの数を取得します。読み取り専用です。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|items|[Binding[]](binding.md)|バインド オブジェクトのコレクション。読み取り専用です。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_プロパティのアクセスの[例を参照してください。](#property-access-examples)_

## <a name="relationships"></a>関係
なし


## <a name="methods"></a>メソッド

| メソッド           | 戻り値の型    |説明| 要件セット|
|:---------------|:--------|:----------|:----|
|[add(range:Range or string, bindingType: string, id: string)](#addrange-range-or-string-bindingtype-string-id-string)|[Binding](binding.md)|特定の範囲に新しいバインドを追加します。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|[addFromNamedItem(name: string, bindingType: string, id: string)](#addfromnameditemname-string-bindingtype-string-id-string)|[Binding](binding.md)|ブック内の名前付きアイテムに基づいて新しいバインドを追加します。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|[addFromSelection(bindingType: string, id: string)](#addfromselectionbindingtype-string-id-string)|[Binding](binding.md)|現在の選択範囲に基づいて新しいバインドを追加します。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|[getItem(id: string)](#getitemid-string)|[Binding](binding.md)|ID を使用してバインド オブジェクトを取得します。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemAt(index: number)](#getitematindex-number)|[Binding](binding.md)|項目の配列内の位置に基づいて、バインド オブジェクトを取得します。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemOrNull(id: string)](#getitemornullid-string)|[Binding](binding.md)|ID を使用してバインド オブジェクトを取得します。バインド オブジェクトが存在しない場合、戻りオブジェクトの isNull プロパティは true になります。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|[load(param: object)](#loadparam-object)|(非推奨)|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>メソッドの詳細


### <a name="addrange-range-or-string-bindingtype-string-id-string"></a>add(range:Range or string, bindingType: string, id: string)
特定の範囲に新しいバインドを追加します。

#### <a name="syntax"></a>構文
```js
bindingCollectionObject.add(range, bindingType, id);
```

#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|:---|
|range|Range または string|バインドをバインドする範囲です。Excel Range オブジェクト、または文字列である場合があります。文字列の場合は、シート名を含む完全なアドレスが含まれている必要があります|
|bindingType|string|バインドの種類です。使用可能な値は次のとおりです。Range、Table、Text|
|id|string|バインドの名前です。|

#### <a name="returns"></a>戻り値
[Binding](binding.md)

### <a name="addfromnameditemname-string-bindingtype-string-id-string"></a>addFromNamedItem(name: string, bindingType: string, id: string)
ブック内の名前付きアイテムに基づいて新しいバインドを追加します。

#### <a name="syntax"></a>構文
```js
bindingCollectionObject.addFromNamedItem(name, bindingType, id);
```

#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|:---|
|name|string|バインドの作成元の名前です。|
|bindingType|string|バインドの種類です。使用可能な値は次のとおりです。Range、Table、Text|
|id|string|バインドの名前です。|

#### <a name="returns"></a>戻り値
[Binding](binding.md)

### <a name="addfromselectionbindingtype-string-id-string"></a>addFromSelection(bindingType: string, id: string)
現在の選択範囲に基づいて新しいバインドを追加します。

#### <a name="syntax"></a>構文
```js
bindingCollectionObject.addFromSelection(bindingType, id);
```

#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|:---|
|bindingType|string|バインドの種類です。使用可能な値は次のとおりです。Range、Table、Text|
|id|string|バインドの名前です。|

#### <a name="returns"></a>戻り値
[Binding](binding.md)

### <a name="getitemid-string"></a>getItem(id: string)
ID によってバインド オブジェクトを取得します。

#### <a name="syntax"></a>構文
```js
bindingCollectionObject.getItem(id);
```

#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|:---|
|id|文字列|取得するバインド オブジェクトの ID。|

#### <a name="returns"></a>戻り値
[Binding](binding.md)

#### <a name="examples"></a>例

テーブルのデータ変更をモニターするためのテーブル バインドを作成します。データが変更されると、テーブルの背景色がオレンジ色に変更されます。

```js
function addEventHandler() {
    //Create Table1
Excel.run(function (ctx) { 
    ctx.workbook.tables.add("Sheet1!A1:C4", true);
    return ctx.sync().then(function() {
             console.log("My Diet Data Inserted!");
    })
    .catch(function (error) {
             console.log(JSON.stringify(error));
    });
});
    //Create a new table binding for Table1
Office.context.document.bindings.addFromNamedItemAsync("Table1", Office.CoercionType.Table, { id: "myBinding" }, function (asyncResult) {
    if (asyncResult.status == "failed") {
        console.log("Action failed with error: " + asyncResult.error.message);
    }
    else {
        // If succeeded, then add event handler to the table binding.
        Office.select("bindings#myBinding").addHandlerAsync(Office.EventType.BindingDataChanged, onBindingDataChanged);
    }
});
}
    
// when data in the table is changed, this event will be triggered.
function onBindingDataChanged(eventArgs) {
Excel.run(function (ctx) { 
    // highlight the table in orange to indicate data has been changed.
    ctx.workbook.bindings.getItem(eventArgs.binding.id).getTable().getDataBodyRange().format.fill.color = "Orange";
    return ctx.sync().then(function() {
            console.log("The value in this table got changed!");
    })
    .catch(function (error) {
            console.log(JSON.stringify(error));
    });
});
}

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


### <a name="getitematindex-number"></a>getItemAt(index: number)
項目の配列内の位置に基づいて、バインド オブジェクトを取得します。

#### <a name="syntax"></a>構文
```js
bindingCollectionObject.getItemAt(index);
```

#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|:---|
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


### <a name="getitemornullid-string"></a>getItemOrNull(id: string)
ID を使用してバインド オブジェクトを取得します。バインド オブジェクトが存在しない場合、戻りオブジェクトの isNull プロパティは true になります。

#### <a name="syntax"></a>構文
```js
bindingCollectionObject.getItemOrNull(id);
```

#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|:---|
|id|文字列|取得するバインド オブジェクトの ID。|

#### <a name="returns"></a>戻り値
[Binding](binding.md)

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
### <a name="property-access-examples"></a>プロパティのアクセスの例

```js
Excel.run(function (ctx) { 
    var bindings = ctx.workbook.bindings;
    bindings.load('items');
    return ctx.sync().then(function() {
        for (var i = 0; i < bindings.items.length; i++)
        {
            console.log(bindings.items[i].id);
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
