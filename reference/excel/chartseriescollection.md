# <a name="chartseriescollection-object-javascript-api-for-excel"></a>ChartSeriesCollection オブジェクト (JavaScript API for Excel)

グラフ系列のコレクションを表します。

## <a name="properties"></a>プロパティ

| プロパティ       | 型    |説明| 要件セット|
|:---------------|:--------|:----------|:----|
|count|int|コレクション内にあるデータ系列の数を取得します。値の取得のみ可能です。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|items|[ChartSeries[]](chartseries.md)|chartSeries オブジェクトのコレクション。値の取得のみ可能です。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_プロパティのアクセスの[例を参照してください。](#property-access-examples)_

## <a name="relationships"></a>関係
なし


## <a name="methods"></a>メソッド

| メソッド           | 戻り値の型    |説明| 要件セット|
|:---------------|:--------|:----------|:----|
|[getCount()](#getcount)|int|コレクション内にあるデータ系列の数を取得します。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemAt(index: number)](#getitematindex-number)|[ChartSeries](chartseries.md)|コレクション内の位置に基づいてデータ系列を取得します。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>メソッドの詳細


### <a name="getcount"></a>getCount()
コレクション内にあるデータ系列の数を取得します。

#### <a name="syntax"></a>構文
```js
chartSeriesCollectionObject.getCount();
```

#### <a name="parameters"></a>パラメーター
なし

#### <a name="returns"></a>戻り値
int

### <a name="getitematindex-number"></a>getItemAt(index: number)
コレクション内の位置に基づいてデータ系列を取得します。

#### <a name="syntax"></a>構文
```js
chartSeriesCollectionObject.getItemAt(index);
```

#### <a name="parameters"></a>パラメーター
| パラメーター       | 型    |説明|
|:---------------|:--------|:----------|:---|
|index|number|取得するオブジェクトのインデックス値。0 を起点とする番号になります。|

#### <a name="returns"></a>戻り値
[ChartSeries](chartseries.md)

#### <a name="examples"></a>例

データ系列コレクション内の最初のデータ系列の名前を取得します。

```js
Excel.run(function (ctx) { 
    var seriesCollection = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").series;
    seriesCollection.load('items');
    return ctx.sync().then(function() {
        console.log(seriesCollection.items[0].name);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### <a name="property-access-examples"></a>プロパティのアクセスの例
データ系列コレクション内のデータ系列の名前を取得する。

```js
Excel.run(function (ctx) { 
    var seriesCollection = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").series;
    seriesCollection.load('items');
    return ctx.sync().then(function() {
        for (var i = 0; i < seriesCollection.items.length; i++)
        {
            console.log(seriesCollection.items[i].name);
        }
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

コレクション内のグラフの系列の数を取得する。

```js
Excel.run(function (ctx) { 
    var seriesCollection = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").series;
    seriesCollection.load('count');
    return ctx.sync().then(function() {
        console.log("series: Count= " + seriesCollection.count);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

