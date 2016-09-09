# オブジェクト読み込みオプション 

sync() メソッドの実行時に読み込まれるプロパティとリレーションシップのセットを指定する load メソッドに渡すことができるオブジェクトを表します。sync() メソッドは、OneNote オブジェクトとそれに対応するアドインの JavaScript のプロキシ オブジェクトの間で状態を同期します。これは、オブジェクトに読み込まれるプロパティのセットを指定する select パラメーターや expand パラメーターなどのオプションを取り、コレクションでの改ページを可能にします。

また、読み込まれるプロパティとリレーションシップを含む文字列、または読み込まれるプロパティとリレーションシップのリストを含む配列の提供にも使用できます。次の例をご覧ください。

```js   
object.load('<var1>,<relationship1/var2>');

// Pass the parameter as an array.
object.load(["var1", "relationship1/var2"]);
```

## プロパティ
| プロパティ     | 型   |説明|
|:---------------|:--------|:----------|
|select|object|同期呼び出しの際に読み込まれるパラメーター名またはリレーションシップ名のコンマ区切りのリストまたは配列を提供します。例: "property1, relationship1", [ "property1", "relationship1"]。省略可能です。|
|expand|object|同期呼び出しの際に読み込まれるリレーションシップ名のコンマ区切りのリストまたは配列を提供します。例: "relationship1, relationship2", [ "relationship1", "relationship2"]。省略可能です。|
|top|int|結果に組み込まれるクエリ コレクション内の項目の数を指定します。省略可能。|
|skip|int|スキップされて結果に含まれないコレクション内の項目の数を指定します。`top` が指定されている場合は、指定された数の項目がスキップされた後で結果の選択が開始されます。省略可能。|

#### 例

例では、現在のセクション内の最初の 5 ページのページ タイトルとインデント レベルを取得します。

```js
OneNote.run(function (context) { 
    
    // Get the pages in the current section.
    var pages = context.application.getActiveSection().pages;
            
    // Queue a command to load the pages.           
    pages.load({ "select":"title,pageLevel", "top":5, "skip":0 });
    return context.sync()
        .then(function() {
            
            // Iterate through the collection of pages.    
            $.each(pages.items, function(index, page) {
                
                // Show some properties.
                console.log("Page title: " + page.title);
                console.log("Indentation level: " + page.pageLevel);
                
            });
        }).catch(function(error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        })
    });
```
