# <a name="pageview-object-javascript-api-for-visio"></a>PageView オブジェクト (JavaScript API for Visio)

適用対象:_Visio Online_
>**注:**Visio JavaScript API は、現在プレビューの段階であり、変更される可能性があります。Visio JavaScript API は、運用環境での使用は現在サポートされていません。

PageView クラスを表します。

## <a name="properties"></a>プロパティ

| プロパティ | 型 |説明| フィードバック|
|:---------------|:--------|:----------|:---|
|ズーム|整数|GetSet ページのズーム レベル。|[移動](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-pageView-zoom)|

## <a name="relationships"></a>関係

なし

## <a name="methods"></a>メソッド

| メソッド           | 戻り値の型    |説明| フィードバック|
|:---------------|:--------|:----------|:---|
|[centerViewportOnShape(ShapeId: number)](#centerviewportonshapeshapeid-number)|void|ビューの中央に指定した図形を配置する Visio の描画をパンします。|[移動](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-pageView-centerViewportOnShape)|
|[fitToWindow()](#fittowindow)|void|現在のウィンドウにページを合わせます。|[移動](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-pageView-fitToWindow)|
|[isShapeInViewport(Shape:Shape)](#isshapeinviewportshape-shape)|bool|図形がページのビュー内にあるかどうかを確認します。|[移動](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-pageView-isShapeInViewport)|
|[load(param: object)](#loadparam-object)|(非推奨)|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|[移動](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-pageView-load)|

## <a name="method-details"></a>メソッドの詳細


### <a name="centerviewportonshapeshapeid-number"></a>centerViewportOnShape(ShapeId: number)
ビューの中央に指定した図形を配置する Visio の描画をパンします。

#### <a name="syntax"></a>構文
```js
pageViewObject.centerViewportOnShape(ShapeId);
```

#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
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

### <a name="isshapeinviewportshape-shape"></a>isShapeInViewport(Shape:Shape)
図形がページのビュー内にあるかどうかを確認します。

#### <a name="syntax"></a>構文
```js
pageViewObject.isShapeInViewport(Shape);
```

#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
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
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|:---|
|param|object|省略可能。パラメーターとリレーションシップ名を、区切られた文字列または 1 つの配列として受け入れます。あるいは、[loadOption](loadoption.md) オブジェクトを提供します。|

#### <a name="returns"></a>戻り値
void

#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
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

