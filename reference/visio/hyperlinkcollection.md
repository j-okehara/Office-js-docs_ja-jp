# <a name="hyperlinkcollection-object-javascript-api-for-visio"></a>HyperlinkCollection オブジェクト (JavaScript API for Visio)

適用対象:_Visio Online_
>**注:**Visio JavaScript API は、現在プレビューまたは運用環境では使用できません。

ハイパーリンク コレクションを表します。

## <a name="properties"></a>プロパティ

| プロパティ     | 型   |説明| フィードバック|
|:---------------|:--------|:----------|:---|
|items|[Hyperlink[]](hyperlink.md)|ハイパーリンク オブジェクトのコレクションです。読み取り専用です。|[移動](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-hyperlinkCollection-items)|

_プロパティのアクセスの[例を参照してください。](#property-access-examples)_

## <a name="relationships"></a>関係
なし


## <a name="methods"></a>メソッド

| メソッド           | 戻り値の型    |説明| フィードバック|
|:---------------|:--------|:----------|:---|
|[getCount()](#getcount)|int|ハイパーリンクの数を取得します。|[移動](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-hyperlinkCollection-getCount)|
|[getItem(Key: number または string)](#getitemkey-number-or-string)|[ハイパーリンク](hyperlink.md)|そのキー (名前または ID) を使用してハイパーリンクを取得します。|[移動](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-hyperlinkCollection-getItem)|
|[load(param: object)](#loadparam-object)|(非推奨)|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|[移動](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-hyperlinkCollection-load)|

## <a name="method-details"></a>メソッドの詳細


### <a name="getcount"></a>getCount()
ハイパーリンクの数を取得します。

#### <a name="syntax"></a>構文
```js
hyperlinkCollectionObject.getCount();
```

#### <a name="parameters"></a>パラメーター
なし

#### <a name="returns"></a>戻り値
int

### <a name="getitemkey-number-or-string"></a>getItem(Key: number または string)
そのキー (名前または ID) を使用してハイパーリンクを取得します。

#### <a name="syntax"></a>構文
```js
hyperlinkCollectionObject.getItem(Key);
```

#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|:---|
|キー|number または string|キーは、取得するハイパーリンクの名前またはインデックスです。|

#### <a name="returns"></a>戻り値
[ハイパーリンク](hyperlink.md)

### <a name="loadparam-object"></a>load(param: object)
JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。

#### <a name="syntax"></a>構文
```js
object.load(param);
```

#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|:---|
|param|object|省略可能。パラメーターとリレーションシップ名を、区切られた文字列または 1 つの配列として受け入れます。あるいは、[loadOption](loadoption.md) オブジェクトを提供します。|

#### <a name="returns"></a>戻り値
void
### <a name="property-access-examples"></a>プロパティのアクセスの例
```js
Visio.run(function (ctx) { 
    var activePage = ctx.document.getActivePage();
    var shapeName = "Manager Belt";
    var shape = activePage.shapes.getItem(shapeName);
    var hyperlinks = shape.hyperlinks;
    shapeHyperlinks.load();
        ctx.sync().then(function () {
            for(var i=0; i<shapeHyperlinks.items.length;i++)
                {
                  var hyperlink = shapeHyperlinks.items[i];
                  console.log("Description:"+hyperlink.description +"Address:"+hyperlink.address +"SubAddress:  "+ hyperlink.subAddress);
                }

            });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
