# <a name="shapedataitemcollection-object-javascript-api-for-visio"></a>ShapeDataItemCollection オブジェクト (JavaScript API for Visio)

適用対象:_Visio Online_
>**注:**Visio JavaScript API は、現在プレビューまたは運用環境では使用できません。

特定の図形の ShapeDataItemCollection を表します。

## <a name="properties"></a>プロパティ

| プロパティ     | 型   |説明| フィードバック|
|:---------------|:--------|:----------|:---|
|items|[ShapeDataItem[]](shapedataitem.md)|shapeDataItem オブジェクトのコレクション。読み取り専用です。|[移動](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shapeDataItemCollection-items)|

_プロパティのアクセスの[例を参照してください。](#property-access-examples)_

## <a name="relationships"></a>関係
なし


## <a name="methods"></a>メソッド

| メソッド           | 戻り値の型    |説明| フィードバック|
|:---------------|:--------|:----------|:---|
|[getCount()](#getcount)|int|図形データ項目の数を取得します。|[移動](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shapeDataItemCollection-getCount)|
|[getItem(key: string)](#getitemkey-string)|[ShapeDataItem](shapedataitem.md)|名前を使用して ShapeDataItem を取得します。|[移動](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shapeDataItemCollection-getItem)|
|[load(param: object)](#loadparam-object)|(非推奨)|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|[移動](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shapeDataItemCollection-load)|

## <a name="method-details"></a>メソッドの詳細


### <a name="getcount"></a>getCount()
図形データ項目の数を取得します。

#### <a name="syntax"></a>構文
```js
shapeDataItemCollectionObject.getCount();
```

#### <a name="parameters"></a>パラメーター
なし

#### <a name="returns"></a>戻り値
int

### <a name="getitemkey-string"></a>getItem(key: string)
名前を使用して ShapeDataItem を取得します。

#### <a name="syntax"></a>構文
```js
shapeDataItemCollectionObject.getItem(key);
```

#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|:---|
|Key|string|キーは、取得する ShapeDataItem の名前です。|

#### <a name="returns"></a>戻り値
[ShapeDataItem](shapedataitem.md)

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
Visio.run(function (ctx) { 
    var activePage = ctx.document.getActivePage();
    var shape = activePage.shapes.getItem(0);
        var shapeDataItems = shape.shapeDataItems;
        shapeDataItems.load();
        return ctx.sync().then(function() {
            for (var i = 0; i < shapeDataItems.items.length; i++)
            {
                console.log(shapeDataItems.items[i].label);
                console.log(shapeDataItems.items[i].value);
            }
        });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
