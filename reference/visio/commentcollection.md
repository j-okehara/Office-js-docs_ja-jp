# <a name="commentcollection-object-javascript-api-for-visio"></a>CommentCollection オブジェクト (JavaScript API for Visio)

適用対象:_Visio Online_

特定の図形の CommentCollection を表します。

## <a name="properties"></a>プロパティ

| プロパティ       | 型    |説明
|:---------------|:--------|:----------|
|items|[Comment[]](comment.md)|コメント オブジェクトのコレクションです。読み取り専用です。|

_プロパティのアクセスの[例を参照してください。](#property-access-examples)_

## <a name="relationships"></a>関係
なし


## <a name="methods"></a>メソッド

| メソッド           | 戻り値の型    |説明|
|:---------------|:--------|:----------|
|[getCount()](#getcount)|int|コメントの数を取得します。|
|[getItem(key: string)](#getitemkey-string)|[Comment](comment.md)|名前を使用してコメントを取得します。|
|[load(param: object)](#loadparam-object)|(非推奨)|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|

## <a name="method-details"></a>メソッドの詳細


### <a name="getcount"></a>getCount()
コメントの数を取得します。

#### <a name="syntax"></a>構文
```js
CommentCollectionObject.getCount();
```

#### <a name="parameters"></a>パラメーター
なし

#### <a name="returns"></a>戻り値
int

### <a name="getitemkey-string"></a>getItem(key: string)
名前を使用してコメントを取得します。

#### <a name="syntax"></a>構文
```js
CommentCollectionObject.getItem(key);
```

#### <a name="parameters"></a>パラメーター
| パラメーター       | 型    |説明|
|:---------------|:--------|:----------|
|Key|string|キーは、取得する Comment の名前です。|

#### <a name="returns"></a>戻り値
[Comment](comment.md)

### <a name="loadparam-object"></a>load(param: object)
JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。

#### <a name="syntax"></a>構文
```js
object.load(param);
```

#### <a name="parameters"></a>パラメーター
| パラメーター       | 型    |説明|
|:---------------|:--------|:----------|
|param|object|省略可能。パラメーターとリレーションシップ名を、区切られた文字列または 1 つの配列として受け入れます。あるいは、[loadOption](loadoption.md) オブジェクトを提供します。|

#### <a name="returns"></a>戻り値
void
### <a name="property-access-examples"></a>プロパティのアクセスの例
```js
 Visio.run(function (ctx) { 
    var activePage = ctx.document.getActivePage();
    var shapeName = "Position Belt.41";
    var shape = activePage.shapes.getItem(shapeName);
    var shapecomments= shape.comments;
        shapecomments.load();
        return ctx.sync().then(function () {
             for(var i=0; i<shapecomments.items.length;i++)
        {
                    var comment= shapecomments.items[i];
            console.log("comment Author: " + comment.author);
            console.log("Comment Text: " + comment.text);
            console.log("Date " + comment.date);
        }
     });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
