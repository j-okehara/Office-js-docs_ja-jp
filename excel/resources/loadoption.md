# オブジェクト読み込みオプション (JavaScript API for Excel)

_適用対象: Excel 2016、Excel Online、Office 2016_

Excel オブジェクトとそれに対応するアドインの JavaScript のプロキシ オブジェクトの間で状態を同期する sync() メソッドの実行時に読み込まれるプロパティと関係のセットを指定する load メソッドに渡すことができるオブジェクトを表します。これは、オブジェクトに読み込まれるプロパティのセットを指定する select パラメータや expand パラメータなどのオプションを取りことができ、コレクションでの改ページを可能にします。

読み込まれるプロパティと関係を含む文字列、または読み込まれるプロパティと関係のリストを含む配列を提供するのにも有効です。例:

```js	
object.load  ('<var1>,<relation1/var2>');

// Pass the parameter as an array.
object.load (["var1", "relation1/var2"]);
```

## プロパティ
| プロパティ   | 型|説明|
|:---------------|:--------|:----------|
|select|object|executeAsync の呼び出しの際に読み込まれる、パラメーター/関係の名前のコンマ区切りのリストまたは配列を提供します。例: "property1, relationship1", [ "property1", "relationship1"]。省略可能。|
|expand|object|executeAsync の呼び出しの際に読み込まれる関係名のコンマ区切りのリストまたは配列を提供します。例: "relationship1, relationship2", [ "relationship1", "relationship2"]。省略可能。|
|top|int| 結果に組み込まれるクエリ コレクション内の項目の数を指定します。省略可能。|
|skip|int|スキップされて結果に含まれないコレクション内の項目の数を指定します。`top` が指定されている場合は、指定された数の項目がスキップされた後で結果の選択が開始されます。省略可能。|

#### 例

以下の例は、表の上から 100 行を選択します。

```js
Excel.run(function (ctx) { 
	var table = ctx.workbook.tables.getItem("Table1");
	var tableRows = table.rows.load({"select" : "index, values","top": 100, "skip": 0 })
	return ctx.sync().then(function() {
		for (var i = 0; i < tableRows.items.length; i++)
		{
			console.log(tableRows.items[i].index);
			console.log(tableRows.items[i].values);
		}
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
})
```
