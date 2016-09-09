# ChartTitle オブジェクト (JavaScript API for Excel)

グラフのグラフ タイトルのオブジェクトを表します。

## プロパティ

| プロパティ     | 型   |説明
|:---------------|:--------|:----------|
|overlay|bool|グラフのタイトルをグラフに重ねるかどうかを表すブール型の値。|
|text|string|グラフのタイトルのテキストを表します。|
|visible|bool|ChartTitle オブジェクトを表示または非表示にするかを表すブール型の値。|

_プロパティのアクセスの[例](#例)を参照してください。_

## リレーションシップ
| リレーションシップ | 型   |説明|
|:---------------|:--------|:----------|
|オプション パラメーターを適用する|[ChartTitleFormat](charttitleformat.md)|グラフ のタイトルの書式設定を表します。これには塗りつぶしとフォントの書式設定などがあります。値の取得のみ可能です。|

## メソッド

| メソッド           | 戻り値の型    |説明|
|:---------------|:--------|:----------|
|[load(param: object)](#loadparam-object)|void|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|

## メソッドの詳細


### load(param: object)
JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。

#### 構文
```js
object.load(param);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|param|object|省略可能。パラメーター名とリレーションシップ名を、区切られた文字列または 1 つの配列として受け入れます。あるいは、[loadOption](loadoption.md) オブジェクトを提供します。|

#### 戻り値
void
### プロパティのアクセスの例

Chart1 のグラフのタイトルの `text` を取得します。

```js
Excel.run(function (ctx) { 
var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1"); 

var title = chart.title;
title.load('text');
return ctx.sync().then(function() {
        console.log(title.text);
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

グラフのタイトルの `text` を "My Chart" に設定し、重ならないようにグラフの先頭に表示されるようにします。

```js
Excel.run(function (ctx) { 
var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1"); 

chart.title.text= "My Chart"; 
chart.title.visible=true;
chart.title.overlay=true;

return ctx.sync().then(function() {
        console.log("Char Title Changed");
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
