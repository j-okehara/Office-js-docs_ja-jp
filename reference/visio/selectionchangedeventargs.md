# <a name="selectionchangedeventargs-object-javascript-api-for-visio"></a>SelectionChangedEventArgs オブジェクト (JavaScript API for Visio)

適用対象:_Visio Online_

SelectionChanged イベントが発生した図形のコレクションに関する情報を提供します。

## <a name="properties"></a>プロパティ

| プロパティ       | 型    |説明
|:---------------|:--------|:----------|
|shapeNames|string[]|SelectionChanged イベントが発生した図形名の配列を取得します。|
|pageName|string|SelectionChanged イベントが発生した ShapeCollection オブジェクトのあるページの名前を取得します。|

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
             eventResult1 = document1.onSelectionChanged.add(
        function (args){
                   console.log("Selected Shape Name: "+args.shapeNames[0]);
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
