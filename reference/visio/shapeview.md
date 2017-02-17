# <a name="shapeview-object-javascript-api-for-visio"></a>ShapeView オブジェクト (JavaScript API for Visio)

適用対象:_Visio Online_
>**注:**Visio JavaScript API は、現在プレビューの段階であり、変更される可能性があります。Visio JavaScript API は、運用環境での使用は現在サポートされていません。

ShapeView クラスを表します。

## <a name="properties"></a>プロパティ

なし

## <a name="relationships"></a>関係
なし

## <a name="methods"></a>メソッド

| メソッド           | 戻り値の型    |説明| フィードバック|
|:---------------|:--------|:----------|:---|
|[addOverlay(OverlayType:OverlayType, Content: string, HorizontalAlignment:HorizontalAlignment, VerticalAlignment:VerticalAlignment, Width: number, Height: number)](#addoverlayoverlaytype-overlaytype-content-string-horizontalalignment-horizontalalignment-verticalalignment-verticalalignment-width-number-height-number)|int|図形の上にオーバーレイを追加します。|[移動](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shapeView-addOverlay)|
|[load(param: object)](#loadparam-object)|(非推奨)|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|[移動](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shapeView-load)|
|[removeOverlay(OverlayId: number)](#removeoverlayoverlayid-number)|void|特定のオーバーレイまたは図形上のすべてのオーバーレイを削除します。|[移動](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shapeView-removeOverlay)|

## <a name="method-details"></a>メソッドの詳細


### <a name="addoverlayoverlaytype-overlaytype-content-string-horizontalalignment-horizontalalignment-verticalalignment-verticalalignment-width-number-height-number"></a>addOverlay(OverlayType:OverlayType, Content: string, HorizontalAlignment:HorizontalAlignment, VerticalAlignment:VerticalAlignment, Width: number, Height: number)
図形の上にオーバーレイを追加します。

#### <a name="syntax"></a>構文
```js
shapeViewObject.addOverlay(OverlayType, Content, HorizontalAlignment, VerticalAlignment, Width, Height);
```

#### <a name="parameters"></a>パラメーター
| パラメーター       | 型    |説明|
|:---------------|:--------|:----------|:---|
|OverlayType|OverlayType|オーバーレイの種類 -テキスト、イメージ。|
|コンテンツ|string|オーバーレイのコンテンツ。|
|HorizontalAlignment|HorizontalAlignment|オーバーレイの水平方向の配置 - 左、中央、右|
|VerticalAlignment|VerticalAlignment|オーバーレイの垂直方向の配置 - 上、上下中央、下|
|幅|number|オーバーレイの幅。|
|高さ|number|オーバーレイの高さ。|

#### <a name="returns"></a>戻り値
int

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

### <a name="removeoverlayoverlayid-number"></a>removeOverlay(OverlayId: number)
特定のオーバーレイまたは図形上のすべてのオーバーレイを削除します。

#### <a name="syntax"></a>構文
```js
shapeViewObject.removeOverlay(OverlayId);
```

#### <a name="parameters"></a>パラメーター
| パラメーター       | 型    |説明|
|:---------------|:--------|:----------|:---|
|OverlayId|number|オーバーレイの ID。図形から特定のオーバーレイの ID を削除します。|

#### <a name="returns"></a>戻り値
void

### <a name="property-access-examples"></a>プロパティのアクセスの例
```js
Visio.run(function (ctx) { 
    var activePage = ctx.document.getActivePage();
    var shape = activePage.shapes.getItem(0);
    var overlayId=shape.view.addOverlay(1, "Visio Online", 2, 2, 50, 50);
    return ctx.sync();
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
    shape.view.removeOverlay(1);
    return ctx.sync();
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
