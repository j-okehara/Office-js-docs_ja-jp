# <a name="pageloadcompleteeventargs-object-javascript-api-for-visio"></a>PageLoadCompleteEventArgs オブジェクト (JavaScript API for Visio)

適用対象:_Visio Online_

PageLoadComplete イベントが発生したページに関する情報を提供します。

## <a name="properties"></a>プロパティ

| プロパティ       | 型    |説明
|:---------------|:--------|:----------|
|pageName|string|PageLoad イベントが発生したページの名前を取得します。|
|success|bool|PageLoadComplete イベントの成否を取得します。|

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
             eventResult1 = document1.onPageLoadComplete.add(
            function (args){
                   console.log("Page name: "+args.pageName);
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
