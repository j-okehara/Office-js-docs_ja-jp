# <a name="chartpointscollection-object-javascript-api-for-excel"></a>ChartPointsCollection オブジェクト (JavaScript API for Excel)

グラフ内の系列内のすべてのグラフのポイントのコレクションです。

## <a name="properties"></a>プロパティ

| プロパティ       | 型    |説明| 要件セット|
|:---------------|:--------|:----------|:----|
|count|int|系列内にあるグラフのポイントの数を取得します。読み取り専用です。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|items|[ChartPoint[]](chartpoint.md)|chartPoints オブジェクトのコレクション。値の取得のみ可能です。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_プロパティのアクセスの[例を参照してください。](#property-access-examples)_

## <a name="relationships"></a>関係
なし


## <a name="methods"></a>メソッド

| メソッド           | 戻り値の型    |説明| 要件セット|
|:---------------|:--------|:----------|:----|
|[getCount()](#getcount)|int|系列内にあるグラフのポイントの数を取得します。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemAt(index: number)](#getitematindex-number)|[ChartPoint](chartpoint.md)|データ系列内の位置に基づくポイントを取得します。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>メソッドの詳細


### <a name="getcount"></a>getCount()
系列内にあるグラフのポイントの数を取得します。

#### <a name="syntax"></a>構文
```js
chartPointsCollectionObject.getCount();
```

#### <a name="parameters"></a>パラメーター
なし

#### <a name="returns"></a>戻り値
int

### <a name="getitematindex-number"></a>getItemAt(index: number)
データ系列内の位置に基づくポイントを取得します。

#### <a name="syntax"></a>構文
```js
chartPointsCollectionObject.getItemAt(index);
```

#### <a name="parameters"></a>パラメーター
| パラメーター       | 型    |説明|
|:---------------|:--------|:----------|:---|
|index|number|取得するオブジェクトのインデックス値。0 を起点とする番号になります。|

#### <a name="returns"></a>戻り値
[ChartPoint](chartpoint.md)

#### <a name="examples"></a>例
points コレクション内の最初の要素の境界線の色を設定します。

```js
Excel.run(function (ctx) { 
    var points = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").series.getItemAt(0).points;
    points.getItemAt(0).format.fill.setSolidColor("8FBC8F");
    return ctx.sync().then(function() {
        console.log("Point Border Color Changed");
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```### Property access examples

Get the names of points in the points collection

```js
Excel.run(function (ctx) { 
    var pointsCollection = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").series.getItemAt(0).points;
    pointsCollection.load('items');
    return ctx.sync().then(function() {
        console.log("Points Collection loaded");
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

ポイント数を取得します。

```js
Excel.run(function (ctx) { 
    var pointsCollection = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").series.getItemAt(0).points;
    pointsCollection.load('count');
    return ctx.sync().then(function() {
        console.log("points: Count= " + pointsCollection.count);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
