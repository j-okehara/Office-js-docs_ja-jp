# Workbook オブジェクト (JavaScript API for Excel)

_適用対象: Excel 2016、Excel Online、Office 2016_

ブックは、ワークシート、表、範囲などの関連するブック オブジェクトを含む最上位オブジェクトです。

## プロパティ

なし

## 関係
| リレーションシップ | 型|説明|
|:---------------|:--------|:----------|
|アプリケーション|[Application](application.md)|このブックを含む Excel アプリケーションのインスタンスを表します。読み取り専用。|
|bindings|[BindingCollection](bindingcollection.md)|ブックの一部であるバインドのコレクションを表します。読み取り専用。|
|名前|[NamedItemCollection](nameditemcollection.md)|ブック スコープの名前付き項目 (名前付き範囲と名前付き定数) のコレクションを表します。読み取り専用。|
|テーブル|[TableCollection](tablecollection.md)|ブックに関連付けられているテーブルのコレクションを表します。読み取り専用。|
|ワークシート|[WorksheetCollection](worksheetcollection.md)|ブックに関連付けられているワークシートのコレクションを表します。読み取り専用。|

## メソッド

| メソッド   | 戻り値の型|説明|
|:---------------|:--------|:----------|
|[getSelectedRange()](#getselectedrange)|[Range](range.md)|ブックから現在選択されている範囲を取得します。|
|[load(param: object)](#loadparam-object)|void|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|

## メソッドの詳細

### getSelectedRange()
ブックから現在選択されている範囲を取得します。

#### 構文
```js
workbookObject.getSelectedRange();
```

#### パラメーター
なし

#### 戻り値
[Range](range.md)

#### 例

```js
Excel.run(function (ctx) { 
	var selectedRange = ctx.workbook.getSelectedRange();
	selectedRange.load('address');
	return ctx.sync().then(function() {
			console.log(selectedRange.address);
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

#### パラメータ
| パラメーター   | 型|説明|
|:---------------|:--------|:----------|
|param|object|省略可能。パラメーター名とリレーションシップ名を、区切られた文字列または 1 つの配列として受け入れます。あるいは、[loadOption](loadoption.md) オブジェクトを提供します。|

#### 戻り値
void

