# NotebookCollection オブジェクト (JavaScript API for OneNote)

_適用対象:OneNote Online_  


ノートブックのコレクションを表します。

## プロパティ

| プロパティ     | 型   |説明|フィードバック|
|:---------------|:--------|:----------|:-------|
|count|int|コレクション内のノートブックの数を取得します。読み取り専用です。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-notebookCollection-count)|
|Items|[Notebook[]](notebook.md)|ノートブック オブジェクトのコレクションです。読み取り専用です。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-notebookCollection-items)|

_プロパティのアクセスの[例](#例)を参照してください。_

## リレーションシップ
なし


## メソッド

| メソッド           | 戻り値の型    |説明| フィードバック|
|:---------------|:--------|:----------|:-------|
|[getByName(name: string)](#getbynamename-string)|[NotebookCollection](notebookcollection.md)|アプリケーション インスタンスで開いている、指定された名前のノートブックのコレクションを取得します。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-notebookCollection-getByName)|
|[getItem(index: number or string)](#getitemindex-number-or-string)|[ノートブック](notebook.md)|ID やコレクション内のインデックスで、ノートブックを取得します。読み取り専用です。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-notebookCollection-getItem)|
|[getItemAt(index: number)](#getitematindex-number)|[ノートブック](notebook.md)|コレクション内での位置を基にノートブックを取得します。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-notebookCollection-getItemAt)|
|[load(param: object)](#loadparam-object)|void|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-notebookCollection-load)|

## メソッドの詳細


### getByName(name: string)
アプリケーション インスタンスで開いている、指定された名前のノートブックのコレクションを取得します。

#### 構文
```js
notebookCollectionObject.getByName(name);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|name|string|ノートブックの名前。|

#### 戻り値
[NotebookCollection](notebookcollection.md)

#### 例
```js
OneNote.run(function (context) {

    // Get the notebooks that are open in the application instance and have the specified name.
    var notebooks = context.application.notebooks.getByName("Homework");

    // Queue a command to load the notebooks. 
    // For best performance, request specific properties.           
    notebooks.load("id,name");

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {

            // Iterate through the collection or access items individually by index, for example: notebooks.items[0]
            if (notebooks.items.length > 0) {
                console.log("Notebook name: " + notebooks.items[0].name);
                console.log("Notebook ID: " + notebooks.items[0].id);
            }
                
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

### getItem(index: number または string)
ID やコレクション内のインデックスで、ノートブックを取得します。読み取り専用です。

#### 構文
```js
notebookCollectionObject.getItem(index);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|index|number または string|ノートブックの ID、またはコレクション内のノートブックのインデックスの場所です。|

#### 戻り値
[ノートブック](notebook.md)

### getItemAt(index: number)
コレクション内での位置を基にノートブックを取得します。

#### 構文
```js
notebookCollectionObject.getItemAt(index);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|index|number|取得するオブジェクトのインデックス値。0 を起点とする番号になります。|

#### 戻り値
[ノートブック](notebook.md)

### load(param: object)
JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。

#### 構文
```js
object.load(param);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|param|object|省略可能。パラメーターとリレーションシップ名を、区切られた文字列または 1 つの配列として受け入れます。あるいは、[loadOption](loadoption.md) オブジェクトを提供します。|

#### 戻り値
void
### プロパティのアクセスの例

**Items**
```js
OneNote.run(function (context) {

    // Get the notebooks that are open in the application instance and have the specified name.
    var notebooks = context.application.notebooks.getByName("Homework");

    // Queue a command to load the notebooks. 
    // For best performance, request specific properties.           
    notebooks.load("id");

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {

            // Iterate through the collection or access items individually by index, for example: notebooks.items[0]
            $.each(notebooks.items, function(index, notebook) {
                notebook.addSection("Biology");
                notebook.addSection("Spanish");
                notebook.addSection("Computer Science");
            });
            
            return context.sync();
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

