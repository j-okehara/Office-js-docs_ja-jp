# <a name="pageview-object-javascript-api-for-visio"></a>PageView オブジェクト (JavaScript API for Visio)

適用対象:_Visio Online_

PageView クラスを表します。

## <a name="properties"></a>プロパティ

| プロパティ | 型 |説明|
|:---------------|:--------|:----------|
|ズーム|整数|GetSet ページのズーム レベル。|

## <a name="relationships"></a>関係
なし

## <a name="methods"></a>メソッド

| メソッド           | 戻り値の型    |説明|
|:---------------|:--------|:----------|
|[centerViewportOnShape(ShapeId: number)](#centerviewportonshapeshapeid-number)|void|ビューの中央に指定した図形を配置する Visio の描画をパンします。|
|[fitToWindow()](#fittowindow)|void|現在のウィンドウにページを合わせます。|
|[getPosition()](#getposition)|[Position](position.md)|ビューでページの位置を指定する位置オブジェクトを返します。|
|[getSelection()](#getselection)|[Selection](selection.md)|ページの選択範囲を表します。|
|[isShapeInViewport(Shape:Shape)](#isshapeinviewportshape-shape)|bool|図形がページのビュー内にあるかどうかを確認します。|
|[load(param: object)](#loadparam-object)|(非推奨)|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|
|[setPosition(Position:Position)](#setpositionposition-position)|void|ビューでページの位置を設定します。|

## <a name="method-details"></a>メソッドの詳細


### <a name="centerviewportonshapeshapeid-number"></a>centerViewportOnShape(ShapeId: number)
ビューの中央に指定した図形を配置する Visio の描画をパンします。

#### <a name="syntax"></a>構文
```js
pageViewObject.centerViewportOnShape(ShapeId);
```

#### <a name="parameters"></a>パラメーター
| パラメーター       | 型    |説明|
|:---------------|:--------|:----------|:---|
|ShapeId|number|中央に表示するため ShapeId。|

#### <a name="returns"></a>戻り値
void

#### <a name="examples"></a>例
```js
Visio.run(function (ctx) { 
    var activePage = ctx.document.getActivePage();
    var shape = activePage.shapes.getItem(0);
    activePage.view.centerViewportOnShape(shape.Id);
    return ctx.sync();
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="fittowindow"></a>fitToWindow()
現在のウィンドウにページを合わせます。

#### <a name="syntax"></a>構文
```js
pageViewObject.fitToWindow();
```

#### <a name="parameters"></a>パラメーター
なし

#### <a name="returns"></a>戻り値
void

### <a name="getposition"></a>getPosition()
ビューでページの位置を指定する位置オブジェクトを返します。

#### <a name="syntax"></a>構文
```js
pageViewObject.getPosition();
```

#### <a name="parameters"></a>パラメーター
なし

#### <a name="returns"></a>戻り値
[Position](position.md)

### <a name="getselection"></a>getSelection()
ページの選択範囲を表します。

#### <a name="syntax"></a>構文
```js
pageViewObject.getSelection();
```

#### <a name="parameters"></a>パラメーター
なし

#### <a name="returns"></a>戻り値
[Selection](selection.md)

### <a name="isshapeinviewportshape-shape"></a>isShapeInViewport(Shape:Shape)
図形がページのビュー内にあるかどうかを確認します。

#### <a name="syntax"></a>構文
```js
pageViewObject.isShapeInViewport(Shape);
```

#### <a name="parameters"></a>パラメーター
| パラメーター       | 型    |説明|
|:---------------|:--------|:----------|:---|
|Shape|Shape|チェックする図形。|

#### <a name="returns"></a>戻り値
bool

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

### <a name="setpositionposition-position"></a>setPosition(Position:Position)
ビューでページの位置を設定します。

#### <a name="syntax"></a>構文
```js
pageViewObject.setPosition(Position);
```

#### <a name="parameters"></a>パラメーター
| パラメーター       | 型    |説明|
|:---------------|:--------|:----------|:---|
|位置|位置|ビューで、ページの新しい位置を指定する位置オブジェクト。|

#### <a name="returns"></a>戻り値
void
### <a name="property-access-examples"></a>プロパティのアクセスの例
```js
Visio.run(function (ctx) { 
    var activePage = ctx.document.getActivePage();
    activePage.view.zoom = 300;
    return ctx.sync();
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

