# RequestContext オブジェクト (JavaScript API for Excel)

_適用対象: Excel 2016、Excel Online、Office 2016_

RequestContext オブジェクトは、Excel アプリケーションへの要求を容易にします。Office アドインと Excel アプリケーションは 2 つの異なるプロセスで実行されているため、アドインから Excel とその関連オブジェクト (ワークシートや表など) にアクセスするには要求のコンテキストが必要です。 

## プロパティ
なし

## メソッド

| メソッド         | 戻り値の型    |説明|
|:---------------|:--------|:----------|
|[load(object: object, option: object)](#loadobject-object-option-object)  |void     |JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオプションを設定します。|

## API 仕様

### load(object: object, option: object)
JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオプションを設定します。

#### 構文
```js
requestContextObject.load(object, loadOption);
```

#### パラメーター
| パラメーター       | 型    |説明|
|:----------------|:--------|:----------|
|object|object|省略可能。読み込むオブジェクトの名前を指定します。|
|オプション|[loadOption](loadoption.md)|省略可能。select、expand、skip、top などの読み込みオプションを指定します。詳細については、loadOption オブジェクトを参照してください。|

#### 戻り値
(非推奨)

##### 例

次の例では、1 つの範囲からプロパティ値を読み込んで、それらを別の範囲にコピーしています。

```js
Excel.run(function (ctx) { 
	var range = ctx.workbook.worksheets.getActiveWorksheet().getRange("A1:A2");
	ctx.load(range, "values");
	return ctx.sync().then(function() {
		var myvalues=range.values;
		ctx.workbook.worksheets.getActiveWorksheet().getRange("B1:B2").values = myvalues;
		console.log(range.values);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
})
```

