# <a name="shape-object-javascript-api-for-visio"></a>Shape オブジェクト (JavaScript API for Visio)

適用対象:_Visio Online_

図形クラスを表します。

## <a name="properties"></a>プロパティ

| プロパティ       | 型    |説明|
|:---------------|:--------|:----------|
|id|int|図形の識別子。読み取り専用です。|
|name|string|図形の名前。読み取り専用です。|
|select|bool|図形が選択されている場合は true を返します。ユーザーが true に設定すれば、明示的に図形を選択できます。|[移動](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shape-select)|
|text|string|図形のテキスト。読み取り専用です。|

_プロパティのアクセスの[例を参照してください。](#property-access-examples)_

## <a name="relationships"></a>関係
| リレーションシップ | 型    |説明|
|:---------------|:--------|:----------|
|コメント|[CommentCollection](commentcollection.md)|Comments コレクションを返します。読み取り専用です。|
|ハイパーリンク|[HyperlinkCollection](hyperlinkcollection.md)|図形オブジェクトのハイパーリンク コレクションを返します。読み取り専用です。|
|shapeDataItems|[ShapeDataItemCollection](shapedataitemcollection.md)|図形のデータ セクションを返します。読み取り専用です。|
|subShapes|[ShapeCollection](shapecollection.md)|SubShape コレクションを取得します。読み取り専用です。|
|ビュー|[ShapeView](shapeview.md)|図形のビューを返します。読み取り専用です。|

## <a name="methods"></a>メソッド

| メソッド           | 戻り値の型    |説明|
|:---------------|:--------|:----------|
|[getBounds()](#getbounds)|[BoundingBox](boundingbox.md)|図形の境界ボックスを指定する BoundingBox オブジェクトを返します。|
|[load(param: object)](#loadparam-object)|(非推奨)|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|

## <a name="method-details"></a>メソッドの詳細


### <a name="getbounds"></a>getBounds()
図形の境界ボックスを指定する BoundingBox オブジェクトを返します。

#### <a name="syntax"></a>構文
```js
shapeObject.getBounds();
```

#### <a name="parameters"></a>パラメーター
なし

#### <a name="returns"></a>戻り値
[BoundingBox](boundingbox.md)

### <a name="loadparam-object"></a>load(param: object)
JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。

#### <a name="syntax"></a>構文
```js
object.load(param);
```

#### <a name="parameters"></a>パラメーター
| パラメーター       | 型    |説明|
|:---------------|:--------|:----------|:---|
|param|object|省略可能。パラメーターとリレーションシップ名を、区切られた文字列または 1 つの配列として受け入れます。あるいは、[loadOption](loadoption.md) オブジェクトを提供します。|

#### <a name="returns"></a>戻り値
void
### <a name="property-access-examples"></a>プロパティのアクセスの例
```js
Visio.run(function (ctx) { 
    var activePage = ctx.document.getActivePage();
    var shapeName = "Sample Name";
    var shape = activePage.shapes.getItem(shapeName);
    shape.load();
    return ctx.sync().then(function () {
        console.log(shape.name );
        console.log(shape.id );
        console.log(shape.Text );
        console.log(shape.Select );
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### <a name="property-access-examples"></a>プロパティのアクセスの例
```js
Visio.run(function (ctx) { 
    var activePage = ctx.document.getActivePage();
    var shape = activePage.shapes.getItem(0);
    shape.view.highlight = { color: "#E7E7E7", width: 100 };
    return ctx.sync();
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
