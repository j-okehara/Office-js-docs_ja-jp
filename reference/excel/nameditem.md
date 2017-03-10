# <a name="nameditem-object-javascript-api-for-excel"></a>NamedItem オブジェクト (JavaScript API for Excel)

セルまたは値の範囲の定義済みの名前を表します。名前には、(以下の型に見られるような) プリミティブ名前付きオブジェクト、範囲オブジェクト、範囲への参照を設定できます。このオブジェクトを使用して、名前に関連付けられた範囲オブジェクトを取得することができます。

## <a name="properties"></a>プロパティ

| プロパティ       | 型    |説明| 要件セット|
|:---------------|:--------|:----------|:----|
|comment|string|この名前に関連付けられているコメントを表します。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|name|string|オブジェクトの名前。値の取得のみ可能です。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|scope|string|名前がブックを対象にしているのか、特定のワークシートを対象にしているのかを示します。読み取り専用です。使用可能な値は次のとおりです。Equal、Greater、GreaterEqual、Less、LessEqual、NotEqual。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|type|string|名前の数式によって返される値の型を示します。読み取り専用です。使用可能な値は次のとおりです。String、Integer、Double、Boolean、Range。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|value|object|名前の数式で計算された値を表します。名前付き範囲の場合は範囲のアドレスを返します。読み取り専用です。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|visible|bool|オブジェクトを表示するかどうかを指定します。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_プロパティのアクセスの[例を参照してください。](#property-access-examples)_

## <a name="relationships"></a>関係
| リレーションシップ | 型    |説明| 要件セット|
|:---------------|:--------|:----------|:----|
|ワークシート|[Worksheet](worksheet.md)|名前付きのアイテムの対象になるワークシートを返します。アイテムがブックを対象にしている場合は、エラーをスローします。読み取り専用です。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|worksheetOrNullObject|[Worksheet](worksheet.md)|名前付きのアイテムの対象になるワークシートを返します。アイテムがブックを対象にしている場合は、null オブジェクトを返します。読み取り専用です。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>メソッド

| メソッド           | 戻り値の型    |説明| 要件セット|
|:---------------|:--------|:----------|:----|
|[delete()](#delete)|void|指定された名前を削除します。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getRange()](#getrange)|[Range](range.md)|名前に関連付けられている範囲オブジェクトを返します。名前付きアイテムの型が範囲でない場合、エラーをスローします。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getRangeOrNullObject()](#getrangeornullobject)|[Range](range.md)|名前に関連付けられている範囲オブジェクトを返します。名前付きアイテムの型が範囲でない場合は、null オブジェクトを返します。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>メソッドの詳細


### <a name="delete"></a>delete()
指定された名前を削除します。

#### <a name="syntax"></a>構文
```js
namedItemObject.delete();
```

#### <a name="parameters"></a>パラメーター
なし

#### <a name="returns"></a>戻り値
void

### <a name="getrange"></a>getRange()
名前に関連付けられている範囲オブジェクトを返します。名前付きアイテムの型が範囲でない場合、エラーをスローします。

#### <a name="syntax"></a>構文
```js
namedItemObject.getRange();
```

#### <a name="parameters"></a>パラメーター
なし

#### <a name="returns"></a>戻り値
[Range](range.md)

#### <a name="examples"></a>例

名前に関連付けられている Range オブジェクトを返します。名前の種類が `Range` でない場合は `null` を返します。注:この API は現在、ブック スコープの項目のみをサポートしています。**

```js
Excel.run(function (ctx) { 
    var names = ctx.workbook.names;
    var range = names.getItem('MyRange').getRange();
    range.load('address');
    return ctx.sync().then(function() {
            console.log(range.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getrangeornullobject"></a>getRangeOrNullObject()
名前に関連付けられている範囲オブジェクトを返します。名前付きアイテムの型が範囲でない場合は、null オブジェクトを返します。

#### <a name="syntax"></a>構文
```js
namedItemObject.getRangeOrNullObject();
```

#### <a name="parameters"></a>パラメーター
なし

#### <a name="returns"></a>戻り値
[Range](range.md)
### <a name="property-access-examples"></a>プロパティのアクセスの例

```js
Excel.run(function (ctx) { 
    var names = ctx.workbook.names;
    var namedItem = names.getItem('MyRange');
    namedItem.load('type');
    return ctx.sync().then(function() {
            console.log(namedItem.type);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
