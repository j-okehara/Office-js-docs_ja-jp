# FloatingInk オブジェクト (JavaScript API for OneNote)

_適用対象:OneNote Online_  


インク ストロークのグループを表します。

## プロパティ

| プロパティ     | 型   |説明|フィードバック|
|:---------------|:--------|:----------|:-------|
|id|string|FloatingInk オブジェクトの ID を取得します。 読み取り専用です。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-floatingInk-id)|

_プロパティのアクセスの[例](#例)を参照してください。_

## リレーションシップ
| リレーションシップ | 型   |説明| フィードバック|
|:---------------|:--------|:----------|:-------|
|inkStrokes|[InkStrokeCollection](inkstrokecollection.md)|FloatingInk オブジェクトのストロークを取得します。 読み取り専用です。|[実行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-floatingInk-inkStrokes)|
|pageContent|[ページ コンテンツ](pagecontent.md)|FloatingInk オブジェクトの PageContent の親を取得します。 読み取り専用です。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-floatingInk-pageContent)|

## メソッド

| メソッド           | 戻り値の型    |説明| フィードバック|
|:---------------|:--------|:----------|:-------|
|[load(param: object)](#loadparam-object)|void|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-floatingInk-load)|

## メソッドの詳細


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

**id**
```js
OneNote.run(function(context) {

    // Gets the active page.
    var page = context.application.getActivePage();
    var contents = page.contents;
    
    // Load page contents and their types.
    page.load('contents/type');
    return context.sync()
        .then(function(){
        
            // Load every ink content.
            $.each(contents.items, function(i, content) {
                if (content.type == "Ink")
                {
                    content.load('ink/id');
                }                           
            })
            return context.sync();
        })
        .then(function(){
        
            // Log ID of every ink content.
            $.each(contents.items, function(i, content) {
                if (content.type == "Ink")
                {
                    console.log(content.ink.id);
                }                           
            })              
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
}); 
```
