# PageCollection オブジェクト (JavaScript API for OneNote)

_適用対象:OneNote Online_  


ページのコレクションを表します。

## プロパティ

| プロパティ     | 型   |説明|フィードバック|
|:---------------|:--------|:----------|:-------|
|count|int|コレクション内のページの数を返します。読み取り専用です。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageCollection-count)|
|items|[Page[]](page.md)|ページ オブジェクトのコレクション。読み取り専用です。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageCollection-items)|

_プロパティのアクセスの[例](#例)を参照してください。_

## リレーションシップ
なし


## メソッド

| メソッド           | 戻り値の型    |説明| フィードバック|
|:---------------|:--------|:----------|:-------|
|[getByTitle(title: string)](#getbytitletitle-string)|[PageCollection](pagecollection.md)|指定したタイトルのページのコレクションを取得します。|[実行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageCollection-getByTitle)|
|[getItem(index: number または string)](#getitemindex-number-または-string)|[Page](page.md)|ID やコレクション内のインデックスで、ページを取得します。読み取り専用です。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageCollection-getItem)|
|[getItemAt(index: number)](#getitematindex-number)|[Page](page.md)|コレクション内での位置を基にページを取得します。|[実行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageCollection-getItemAt)|
|[load(param: object)](#loadparam-object)|void|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageCollection-load)|

## メソッドの詳細


### getByTitle(title: string)
指定したタイトルのページのコレクションを取得します。

#### 構文
```js
pageCollectionObject.getByTitle(title);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|title|string|ページのタイトル。|

#### 戻り値
[PageCollection](pagecollection.md)

#### 例
```js
OneNote.run(function (context) {

    // Get all the pages in the current section.
    var allPages = context.application.getActiveSection().pages;

    // Queue a command to load the pages. 
    // For best performance, request specific properties.
    allPages.load("id"); 

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {

            // Get the sections with the specified name.
            var todoPages = allPages.getByTitle("Todo list");

            // Queue a command to load the section. 
            // For best performance, request specific properties.
            todoPages.load("id,title"); 

            return context.sync()
                .then(function () {

                    // Iterate through the collection or access items individually by index.
                    if (todoPages.items.length > 0) {
                        console.log("Page title: " + todoPages.items[0].title);
                        console.log("Page ID: " + todoPages.items[0].id);
                    }
                });
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
ID やコレクション内のインデックスで、ページを取得します。読み取り専用です。

#### 構文
```js
pageCollectionObject.getItem(index);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|index|number または string|ページの ID、またはコレクション内のページのインデックスの場所です。|

#### 戻り値
[Page](page.md)

### getItemAt(index: number)
コレクション内での位置を基にページを取得します。

#### 構文
```js
pageCollectionObject.getItemAt(index);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|index|number|取得するオブジェクトのインデックス値。0 を起点とする番号になります。|

#### 戻り値
[Page](page.md)

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
    
    // Get the pages in the current section.
    var pages = context.application.getActiveSection().pages;
    
    // Queue a command to load the id and title for each page.            
    pages.load('id,title');
    
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            
            // Display the properties.
            $.each(pages.items, function(index, page) {
                console.log(page.title);
                console.log(page.id);
            });
        }); 
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

