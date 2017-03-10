# <a name="datarefreshcompleteeventargs-object-javascript-api-for-visio"></a>DataRefreshCompleteEventArgs オブジェクト (JavaScript API for Visio)

適用対象:_Visio Online_

DataRefreshComplete イベントが発生したドキュメントに関する情報を提供します。

## <a name="properties"></a>プロパティ

| プロパティ       | 型    |説明
|:---------------|:--------|:----------|
|success|bool|DataRefreshComplete イベントの successfailure を取得します。|
|document|[Document](document.md)|DataRefreshComplete イベントが発生したドキュメント オブジェクトを取得します。|

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
         eventResult1 = document1.onDataRefreshComplete.add(
    function (args){
           console.log("Data Refresh Result: "+args.success);
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
