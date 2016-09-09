# TableSort オブジェクト (JavaScript API for Excel)

_適用対象: Excel 2016、Excel Online、Excel for iOS、Office 2016_

Table オブジェクトの並べ替え操作を管理します。

## プロパティ

| プロパティ     | 型   |説明
|:---------------|:--------|:----------|
|matchCase|bool|大文字小文字の区別が、テーブルの最後の並べ替え操作に影響を与えたかどうかを表します。読み取り専用です。|
|method|string|テーブルの並べ替えで最後に使用した中国語文字の順序付け方法を表します。読み取り専用です。使用可能な値は次のとおりです。PinYin、StrokeCount。|

## リレーションシップ
| リレーションシップ | 型   |説明|
|:---------------|:--------|:----------|
|fields|[SortField](sortfield.md)|テーブルの最後の並べ替えに使用する現在の条件を表します。読み取り専用です。|

## メソッド

| メソッド           | 戻り値の型    |説明|
|:---------------|:--------|:----------|
|[apply(fields:SortField[], matchCase: bool, method: string)](#applyfields-sortfield-matchcase-bool-method-string)|void|並べ替え操作を実行します。|
|[clear()](#clear)|void|テーブルに現在設定されている並べ替えをクリアします。これにより表の順序が変更されることはありませんが、ヘッダーのボタンの状態がクリアされます。|
|[load(param: object)](#loadparam-object)|void|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|
|[reapply()](#reapply)|void|テーブルに、現在の並べ替えパラメーターを再適用します。|

## メソッドの詳細


### apply(fields:SortField[], matchCase: bool, method: string)
並べ替え操作を実行します。

#### 構文
```js
tableSortObject.apply(fields, matchCase, method);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|fields|SortField[]|並べ替えに使用する条件の一覧。|
|matchCase|bool|省略可能。大文字小文字の区別が文字列の順序に影響を与えるかどうか。|
|method|string|省略可能。中国語文字に使用される順序付けの方法です。使用可能な値は次のとおりです。PinYin、StrokeCount|

#### 戻り値
void

#### 例
```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    table.sort.apply([ 
            {
                key: 2,
                ascending: true
            },
        ], true);
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### clear()
テーブルに現在設定されている並べ替えをクリアします。これにより表の順序が変更されることはありませんが、ヘッダーのボタンの状態がクリアされます。

#### 構文
```js
tableSortObject.clear();
```

#### パラメーター
なし

#### 戻り値
void

#### 例
```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    table.sort.clear();
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});

### load(param: object)
JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。

#### Syntax
```js
object.load(param);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|param|object|省略可能。パラメーター名とリレーションシップ名を、区切られた文字列または 1 つの配列として受け入れます。あるいは、[loadOption](loadoption.md) オブジェクトを提供します。|

#### 戻り値
void

### reapply()
テーブルに、現在の並べ替えパラメーターを再適用します。

#### 構文
```js
tableSortObject.reapply();
```

#### パラメーター
なし

#### 戻り値
void

####例
```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    table.sort.reapply();   
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});