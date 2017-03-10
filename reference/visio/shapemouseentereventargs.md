# <a name="shapemouseentereventargs-object-javascript-api-for-visio"></a>ShapeMouseEnterEventArgs オブジェクト (JavaScript API for Visio)

適用対象:_Visio Online_

MouseEnter イベントが発生した図形に関する情報を提供します。

## <a name="properties"></a>プロパティ

| プロパティ       | 型    |説明
|:---------------|:--------|:----------|
|shapeName|string|MouseEnter イベントが発生した図形オブジェクトの名前を取得します。|
|pageName|string|MouseEnter イベントが発生した図形オブジェクトのあるページの名前を取得します。|

_プロパティのアクセスの[例を参照してください。](#property-access-examples)_

## <a name="relationships"></a>関係
なし

## <a name="methods"></a>メソッド
なし

### <a name="property-access-examples"></a>プロパティのアクセスの例
```js
Visio.run(function (ctx) { 
  var document1= ctx.document;
               var page = document1.getActivePage();
    eventResult2 = document1.onMouseEnter.add(
            function (args){            
                         console.log(Date.now()+":OnMouseEnter Event"+JSON.stringify(args));
            });
    return ctx.sync().then(function () {
           console.log("Success");
        });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```