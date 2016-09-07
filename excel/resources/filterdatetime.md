# FilterDatetime オブジェクト (JavaScript API for Excel)

_適用対象:Excel 2016、Excel Online、Excel for iOS、Office 2016_

値をフィルター処理するときに日付をフィルター処理する方法を表します。

## プロパティ

| プロパティ	  | 型	|説明
|:---------------|:--------|:----------||date|string|データをフィルター処理するときに使用する、ISO8601 形式の日付。||specificity|string|データを保持するための日付の特定の使用方法。たとえば、date が 2005-04-02 で "month" に設定した場合、フィルター操作では 2005 年 4 月の日付データを含むすべての行が保持されます。使用可能な値は次のいずれかです。Year、Month、Day、Hour、Minute、Second。|_プロパティの使用[例](#property-access-examples)_をご覧ください。

## リレーションシップ
なし


## メソッド

| メソッド		  | 戻り値の型	|説明||:---------------|:--------|:----------||[load(param: object)](#loadparam-object)|void|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|

## メソッドの詳細


### load(param: object)
JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。

#### 構文
```js
object.load(param);
```

#### パラメーター
| パラメーター	  | 型	|説明||:---------------|:--------|:----------||param|object|省略可能。パラメーターとリレーションシップ名を、区切られた文字列または 1 つの配列として受け入れます。あるいは、[loadOption](loadoption.md) オブジェクトを提供します。|

#### 戻り値
void
