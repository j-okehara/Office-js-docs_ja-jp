# <a name="object-load-options-javascript-api-for-visio"></a>オブジェクト読み込みオプション (JavaScript API for Visio)

Visio オブジェクトとそれに対応する JavaScript のプロキシ オブジェクトの間で状態を同期する **sync()** メソッドの実行時に読み込まれるプロパティおよび関係のセットを指定する load メソッドに渡すことができるオブジェクトを表します。これは、オブジェクトに読み込まれるプロパティのセットを指定する select パラメーターおよび expand パラメーターなどのオプションを取り、コレクションでの改ページ位置の自動修正を可能にします。

また、読み込まれるプロパティと関係を含む文字列、または読み込まれるプロパティと関係のリストを含む配列の提供にも使用できます。次の例をご覧ください。

```js
object.load  ('<var1>,<relation1/var2>');

// Pass the parameter as an array.
object.load (["var1", "relation1/var2"]);
```

## <a name="properties"></a>プロパティ

| プロパティ | 型  | 説明 |
|:---------|:------|:------------|
|select    |object |executeAsync の呼び出しの際に読み込まれるパラメーター名またはリレーションシップ名のコンマ区切りのリストまたは配列を提供します。例: "property1, relationship1", [ "property1", "relationship1"]。省略可能です。|
|expand    |object |executeAsync の呼び出しの際に読み込まれるリレーションシップ名のコンマ区切りのリストまたは配列を提供します。例: "relationship1, relationship2", [ "relationship1", "relationship2"]。省略可能です。|
|top       |int    |結果に組み込まれるクエリ コレクション内の項目の数を指定します。省略可能。|
|skip      |int    |スキップされて結果に含まれないコレクション内の項目の数を指定します。**top** が指定されている場合は、指定された数の項目がスキップされた後で結果の選択が開始されます。省略可能。|

