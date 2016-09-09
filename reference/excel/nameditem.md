# NamedItem オブジェクト (JavaScript API for Excel)

セルまたは値の範囲に定義される名前を表します。名前には、(以下の型に見られるような) プリミティブ名前付きオブジェクト、範囲オブジェクト、範囲への参照を設定できます。このオブジェクトを使用して、名前に関連付けられた範囲オブジェクトを取得することができます。

## プロパティ

| プロパティ     | 型   |説明
|:---------------|:--------|:----------|
|name|string|オブジェクトの名前。値の取得のみ可能です。|
|type|string|名前に関連付けられている参照の型を示します。読み取り専用です。使用可能な値は次のとおりです。文字列、整数、倍精度浮動小数点、ブール、範囲。|
|value|object|定義されている名前が参照する数式を表します。たとえば、=Sheet14!$B$2:$H$12、=4.75 など。読み取り専用です。|
|visible|bool|オブジェクトを表示するかどうかを指定します。|

_プロパティのアクセスの[例](#例)を参照してください。_

## リレーションシップ
なし


## メソッド

| メソッド           | 戻り値の型    |説明|
|:---------------|:--------|:----------|
|[getRange()](#getrange)|[範囲](range.md)|名前に関連付けられている範囲オブジェクトを返します。名前付き項目の型が範囲でない場合、例外をスローします。|
|[load(param: object)](#loadparam-object)|void|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|

## メソッドの詳細


### getRange()
名前に関連付けられている範囲オブジェクトを返します。名前付き項目の型が範囲でない場合、例外をスローします。

#### 構文
```js
namedItemObject.getRange();
```

#### パラメーター
なし

#### 戻り値
[範囲](range.md)

#### 例

名前に関連付けられている Range オブジェクトを返します。名前の種類が `null` でない場合は `Range` を返します。注:この API は現在、ブック スコープの項目のみをサポートしています。

```js
Excel.run(function (ctx) { 
    var names = ctx.workbook.names;
    var range = names.getItem('MyRange').getRange();
    range.load('address');
    return ctx.sync().then(function() {
            console.log(range.address);
    });
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
|param|object|省略可能。パラメーター名とリレーションシップ名を、区切られた文字列または 1 つの配列として受け入れます。あるいは、[loadOption](loadoption.md) オブジェクトを提供します。|

#### 戻り値
void
### プロパティのアクセスの例

```js
Excel.run(function (ctx) { 
    var names = ctx.workbook.names;
    var namedItem = names.getItem('MyRange');
    namedItem.load('type');
    return ctx.sync().then(function() {
            console.log(namedItem.type);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
