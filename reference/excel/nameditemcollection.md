# <a name="nameditemcollection-object-javascript-api-for-excel"></a>NamedItemCollection オブジェクト (JavaScript API for Excel)

到達方法に応じた、ブックまたはワークシートの一部である、すべての名前付きアイテム オブジェクトのコレクションです。

## <a name="properties"></a>プロパティ

| プロパティ       | 型    |説明| 要件セット|
|:---------------|:--------|:----------|:----|
|items|[NamedItem[]](nameditem.md)|namedItem オブジェクトのコレクション。読み取り専用です。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_プロパティのアクセスの[例を参照してください。](#property-access-examples)_

## <a name="relationships"></a>関係
なし


## <a name="methods"></a>メソッド

| メソッド           | 戻り値の型    |説明| 要件セット|
|:---------------|:--------|:----------|:----|
|[add(name: string, reference:Range または string, comment: string)](#addname-string-reference-range-or-string-comment-string)|[NamedItem](nameditem.md)|新しい名前を指定したスコープのコレクションに追加します。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[addFormulaLocal(name: string, formula: string, comment: string)](#addformulalocalname-string-formula-string-comment-string)|[NamedItem](nameditem.md)|ユーザーのロケールを数式に使用して、新しい名前を指定したスコープのコレクションに追加します。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getCount()](#getcount)|int|コレクション内の名前付きアイテムの数を取得します。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getItem(name: string)](#getitemname-string)|[NamedItem](nameditem.md)|nameditem オブジェクトを、名前を使用して取得します|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemOrNullObject(name: string)](#getitemornullobjectname-string)|[NamedItem](nameditem.md)|名前を使用して、nameditem オブジェクトを取得します。nameditem オブジェクトが存在しない場合は null オブジェクトを返します。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>メソッドの詳細


### <a name="addname-string-reference-range-or-string-comment-string"></a>add(name: string, reference:Range または string, comment: string)
新しい名前を指定したスコープのコレクションに追加します。

#### <a name="syntax"></a>構文
```js
namedItemCollectionObject.add(name, reference, comment);
```

#### <a name="parameters"></a>パラメーター
| パラメーター       | 型    |説明|
|:---------------|:--------|:----------|:---|
|name|string|名前付きの項目の名前。|
|reference|Range または string|名前が参照する数式または範囲。|
|comment|string|省略可能。名前付きアイテムに関連付けられているコメント。|

#### <a name="returns"></a>戻り値
[NamedItem](nameditem.md)

### <a name="addformulalocalname-string-formula-string-comment-string"></a>addFormulaLocal(name: string, formula: string, comment: string)
ユーザーのロケールを数式に使用して、新しい名前を指定したスコープのコレクションに追加します。

#### <a name="syntax"></a>構文
```js
namedItemCollectionObject.addFormulaLocal(name, formula, comment);
```

#### <a name="parameters"></a>パラメーター
| パラメーター       | 型    |説明|
|:---------------|:--------|:----------|:---|
|name|string|名前付きアイテムの "名前"。|
|formula|string|名前が参照するユーザーのロケールの数式。|
|comment|string|省略可能。名前付きアイテムに関連付けられているコメント。|

#### <a name="returns"></a>戻り値
[NamedItem](nameditem.md)

### <a name="getcount"></a>getCount()
コレクション内の名前付きアイテムの数を取得します。

#### <a name="syntax"></a>構文
```js
namedItemCollectionObject.getCount();
```

#### <a name="parameters"></a>パラメーター
なし

#### <a name="returns"></a>戻り値
int

### <a name="getitemname-string"></a>getItem(name: string)
名前を使用して、nameditem オブジェクトを取得します。

#### <a name="syntax"></a>構文
```js
namedItemCollectionObject.getItem(name);
```

#### <a name="parameters"></a>パラメーター
| パラメーター       | 型    |説明|
|:---------------|:--------|:----------|:---|
|name|string|nameditem 名。|

#### <a name="returns"></a>戻り値
[NamedItem](nameditem.md)

#### <a name="examples"></a>例

```js
Excel.run(function (ctx) { 
    var sheetName = 'Sheet1';
    var nameditem = ctx.workbook.names.getItem(sheetName);
    nameditem.load('type');
    return ctx.sync().then(function() {
            console.log(nameditem.type);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
### <a name="getitemornullobjectname-string"></a>getItemOrNullObject(name: string)
名前を使用して、nameditem オブジェクトを取得します。nameditem オブジェクトが存在しない場合は null オブジェクトを返します。

#### <a name="syntax"></a>構文
```js
namedItemCollectionObject.getItemOrNullObject(name);
```

#### <a name="parameters"></a>パラメーター
| パラメーター       | 型    |説明|
|:---------------|:--------|:----------|:---|
|name|string|nameditem 名。|

#### <a name="returns"></a>戻り値
[NamedItem](nameditem.md)
### <a name="property-access-examples"></a>プロパティのアクセスの例

```js
Excel.run(function (ctx) { 
    var nameditems = ctx.workbook.names;
    nameditems.load('items');
    return ctx.sync().then(function() {
        for (var i = 0; i < nameditems.items.length; i++)
        {
            console.log(nameditems.items[i].name);
            console.log(nameditems.items[i].index);
        }
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


