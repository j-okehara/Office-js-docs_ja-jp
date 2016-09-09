# Application オブジェクト (JavaScript API for Excel)

ブックを管理する Excel アプリケーションを表します。

## プロパティ

| プロパティ     | 型   |説明
|:---------------|:--------|:----------|
|calculationMode|string|ブックで使用される計算モードを返します。値の取得のみ可能です。使用可能な値は次のとおりです。`Automatic` Excel が再計算を制御します。`AutomaticExceptTables` Excel が再計算を制御しますが、テーブル内の変更は無視します。`Manual` 計算は、ユーザーが要求したときに行われます。|

_プロパティのアクセスの[例](#例)をご覧ください。_

## リレーションシップ
なし


## メソッド

| メソッド           | 戻り値の型    |説明|
|:---------------|:--------|:----------|
|[calculate(calculationType: string)](#calculatecalculationtype-string)|void|Excel で現在開いているすべてのブックを再計算します。|
|[load(param: object)](#loadparam-object)|void|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|

## メソッドの詳細


### calculate(calculationType: string)
Excel で現在開いているすべてのブックを再計算します。

#### 構文
```js
applicationObject.calculate(calculationType);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|calculationType|string|使用する計算の種類を指定します。使用可能な値は次のとおりです。`Recalculate` 既定のオプションであり、ブック内のすべての数式を計算して通常の計算を実行します。`Full` データ全体の計算を強制的に実行します。`FullRebuild` データ全体の計算を強制的に実行して依存関係を再構築します。|

#### 戻り値
void

#### 例
```js
Excel.run(function (ctx) { 
    ctx.workbook.application.calculate('Full');
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### load(param: object)
JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。

#### 構文
```js
object.load(param);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|param|object|省略可能。パラメーター名とリレーションシップ名を、区切られた文字列または 1 つの配列として受け入れます。あるいは、[loadOption](loadoption.md) オブジェクトを受け入れます。|

#### 戻り値
void
### プロパティのアクセスの例
```js
Excel.run(function (ctx) { 
    var application = ctx.workbook.application;
    application.load('calculationMode');
    return ctx.sync().then(function() {
        console.log(application.calculationMode);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

