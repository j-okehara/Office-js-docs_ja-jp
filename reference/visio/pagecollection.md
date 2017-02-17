# <a name="pagecollection-object-javascript-api-for-visio"></a>PageCollection オブジェクト (JavaScript API for Visio)

適用対象:_Visio Online_
>**注:**Visio JavaScript API は、現在プレビューの段階であり、変更される可能性があります。Visio JavaScript API は、運用環境での使用は現在サポートされていません。

ドキュメントの一部であるページ オブジェクトのコレクションを表します。

## <a name="properties"></a>プロパティ

| プロパティ       | 型    |説明| フィードバック|
|:---------------|:--------|:----------|:---|
|items|[Page[]](page.md)|ページ オブジェクトのコレクション。読み取り専用です。|[移動](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-pageCollection-items)|

## <a name="relationships"></a>関係
なし


## <a name="methods"></a>メソッド

| メソッド           | 戻り値の型    |説明| フィードバック|
|:---------------|:--------|:----------|:---|
|[getCount()](#getcount)|int|コレクション内のページの数を取得します。|[移動](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-pageCollection-getCount)|
|[getItem(key: number または string)](#getitemkey-number-or-string)|[Page](page.md)|そのキー (名前または ID) を使用してページを取得します。|[移動](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-pageCollection-getItem)|
|[load(param: object)](#loadparam-object)|(非推奨)|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|[移動](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-pageCollection-load)|

## <a name="method-details"></a>メソッドの詳細


### <a name="getcount"></a>getCount()
コレクション内のページの数を取得します。

#### <a name="syntax"></a>構文
```js
pageCollectionObject.getCount();
```

#### <a name="parameters"></a>パラメーター
なし

#### <a name="returns"></a>戻り値
int

### <a name="getitemkey-number-or-string"></a>getItem(key: number または string)
そのキー (名前または ID) を使用してページを取得します。

#### <a name="syntax"></a>構文
```js
pageCollectionObject.getItem(key);
```

#### <a name="parameters"></a>パラメーター
| パラメーター       | 型    |説明|
|:---------------|:--------|:----------|:---|
|Key|number または string|キーは、取得するページの名前または ID です。|

#### <a name="returns"></a>戻り値
[Page](page.md)

#### <a name="examples"></a>例
```js
Visio.run(function (ctx) { 
    var pageName = 'Page-1';
    var page = ctx.document.pages.getItem(pageName);
    page.activate();
    return ctx.sync();
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### <a name="loadparam-object"></a>load(param: object)
JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。

#### <a name="syntax"></a>構文
```js
object.load(param);
```

#### <a name="parameters"></a>パラメーター
| パラメーター       | 型    |説明|
|:---------------|:--------|:----------|:---|
|param|object|省略可能。パラメーターとリレーションシップ名を、区切られた文字列または 1 つの配列として受け入れます。あるいは、[loadOption](loadoption.md) オブジェクトを提供します。|

#### <a name="returns"></a>戻り値
void
