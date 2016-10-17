# <a name="pagecontentcollection-object-(javascript-api-for-onenote)"></a>PageContentCollection オブジェクト (JavaScript API for OneNote)

_適用対象:OneNote Online_  


PageContent オブジェクトのコレクションとして、ページのコンテンツを表します。

## <a name="properties"></a>プロパティ

| プロパティ     | 型   |説明|フィードバック|
|:---------------|:--------|:----------|:-------|
|count|int|コレクション内のページ コンテンツの数を返します。読み取り専用です。|[移動](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContentCollection-count)|
|items|[PageContent[]](pagecontent.md)|pageContent オブジェクトのコレクション。読み取り専用です。|[移動](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContentCollection-items)|

_プロパティのアクセスの[例を参照してください。](#property-access-examples)_

## <a name="relationships"></a>関係
なし


## <a name="methods"></a>メソッド

| メソッド           | 戻り値の型    |説明| フィードバック|
|:---------------|:--------|:----------|:-------|
|[getItem(index: number or string)](#getitemindex-number-or-string)|[PageContent](pagecontent.md)|ID やコレクション内のインデックスで、PageContent オブジェクトを取得します。読み取り専用です。|[移動](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContentCollection-getItem)|
|[getItemAt(index: number)](#getitematindex-number)|[PageContent](pagecontent.md)|コレクション内での位置を基にページ コンテンツを取得します。|[移動](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContentCollection-getItemAt)|
|[load(param: object)](#loadparam-object)|(非推奨)|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|[移動](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContentCollection-load)|

## <a name="method-details"></a>メソッドの詳細


### <a name="getitem(index:-number-or-string)"></a>getItem(index: number または string)
ID やコレクション内のインデックスで、PageContent オブジェクトを取得します。読み取り専用です。

#### <a name="syntax"></a>構文
```js
pageContentCollectionObject.getItem(index);
```

#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|index|number または string|PageContent オブジェクト の ID、またはコレクション内の PageContent オブジェクトのインデックスの場所です。|

#### <a name="returns"></a>戻り値
[PageContent](pagecontent.md)

### <a name="getitemat(index:-number)"></a>getItemAt(index: number)
コレクション内での位置を基にページ コンテンツを取得します。

#### <a name="syntax"></a>構文
```js
pageContentCollectionObject.getItemAt(index);
```

#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|index|number|取得するオブジェクトのインデックス値。0 を起点とする番号になります。|

#### <a name="returns"></a>戻り値
[PageContent](pagecontent.md)

#### <a name="examples"></a>例
```js
OneNote.run(function (context) {

    var page = context.application.getActivePage();
    var pageContents = page.contents;
    var firstPageContent = pageContents.getItemAt(0);
    firstPageContent.load('type');

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            console.log("The first page content item is of type: " + firstPageContent.type);
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

### <a name="load(param:-object)"></a>load(param: object)
JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。

#### <a name="syntax"></a>構文
```js
object.load(param);
```

#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|param|object|省略可能。パラメーターとリレーションシップ名を、区切られた文字列または 1 つの配列として受け入れます。あるいは、[loadOption](loadoption.md) オブジェクトを提供します。|

#### <a name="returns"></a>戻り値
void
### <a name="property-access-examples"></a>プロパティのアクセスの例

**items**
```js
OneNote.run(function (context) {

    // Get the collection of pageContent items from the page.
    var pageContents = context.application.getActivePage().contents;

    // Queue a command to load the type of each pageContent.
    pageContents.load("type");

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {

            $.each(pageContents.items, function(index, pageContent) {
                console.log("PageContent type: " + pageContent.type);
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

**アウトラインの走査**
```js
OneNote.run(function (context) {
   var page = context.application.getActivePage();
   var pageContents = page.contents;
   pageContents.load('type');
   var outlines = [];
   return context.sync()
       .then(function () {    
              $.each(pageContents.items, function (index, pageContent) {
                     console.log(pageContent.type);
                     if (pageContent.type === 'Outline') {
                           outlines.push(pageContent);
                     }
              });
              $.each(outlines, function (index, outline) {
                     outline.load("id,paragraphs,paragraphs/type");
              });
              return context.sync();
       })
       .then(function () {
              $.each(outlines, function (index, outline) {
                     console.log("An outline was found with id : " + outline.id);
              });
              return Promise.resolve(outlines);
       });
});
```

