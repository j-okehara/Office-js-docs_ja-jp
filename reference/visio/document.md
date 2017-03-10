# <a name="document-object-javascript-api-for-visio"></a>Document オブジェクト (JavaScript API for Visio)

適用対象:_Visio Online_

ドキュメント クラスを表します。

## <a name="properties"></a>プロパティ

なし

## <a name="relationships"></a>関係
| リレーションシップ | 型    |説明|
|:---------------|:--------|:----------|
|application|[アプリケーション](application.md)|このドキュメントを含む Visio アプリケーションのインスタンスを表します。読み取り専用です。|
|pages|[PageCollection](pagecollection.md)|ドキュメントに関連付けられているページのコレクションを表します。読み取り専用です。|
|ビュー|[DocumentView](documentview.md)|DocumentView オブジェクトを返します。読み取り専用です。|

## <a name="methods"></a>メソッド

| メソッド           | 戻り値の型    |説明|
|:---------------|:--------|:----------|
|[getActivePage()](#getactivepage)|[Page](page.md)|ドキュメントのアクティブ ページを返します。|
|[load(param: object)](#loadparam-object)|(非推奨)|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|
|[setActivePage(PageName: string)](#setactivepagepagename-string)|void|ドキュメントのアクティブ ページを設定します。|
|[startDataRefresh()](#startdatarefresh)|void|すべてのページについて、図内のデータの更新をトリガーします。|

## <a name="method-details"></a>メソッドの詳細


### <a name="getactivepage"></a>getActivePage()
ドキュメントのアクティブ ページを返します。

#### <a name="syntax"></a>構文
```js
documentObject.getActivePage();
```

#### <a name="parameters"></a>パラメーター
なし

#### <a name="returns"></a>戻り値
[Page](page.md)

#### <a name="examples"></a>例
```js
Visio.run(function (ctx) { 
    var document = ctx.document;
    var activePage = document.getActivePage();
    activePage.load();
    return ctx.sync().then(function () {
    console.log("pageName: " +activePage.name);
      });   
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

### <a name="setactivepagepagename-string"></a>setActivePage(PageName: string)
ドキュメントのアクティブ ページを設定します。

#### <a name="syntax"></a>構文
```js
documentObject.setActivePage(PageName);
```

#### <a name="parameters"></a>パラメーター
| パラメーター       | 型    |説明|
|:---------------|:--------|:----------|:---|
|PageName|string|ページの名前|

#### <a name="returns"></a>戻り値
void

#### <a name="examples"></a>例
```js
Visio.run(function (ctx) { 
    var document = ctx.document;
    var pageName = "Page-1";
    document.setActivePage(pageName);
    return ctx.sync();
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="startdatarefresh"></a>startDataRefresh()
すべてのページについて、図内のデータの更新をトリガーします。

#### <a name="syntax"></a>構文
```js
documentObject.startDataRefresh();
```

#### <a name="parameters"></a>パラメーター
なし

#### <a name="returns"></a>戻り値
void

#### <a name="examples"></a>例
```js
Visio.run(function (ctx) { 
    var document = ctx.document;
    document.startDataRefresh();
    return ctx.sync();
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```### Property access examples
```js
Visio.run(function (ctx) { 
    var pages = ctx.document.pages;
    var pageCount = pages.getCount();
    return ctx.sync().then(function () {
        console.log("Pages Count: " +pageCount.value);
        });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### <a name="property-access-examples"></a>プロパティのアクセスの例
```js
Visio.run(function (ctx) { 
    var documentView = ctx.document.view;
    documentView.disableHyperlinks();
    return ctx.sync();
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### <a name="property-access-examples"></a>プロパティのアクセスの例
```js
Visio.run(function (ctx) { 
    var application = ctx.document.application;
    application.showToolbars = false;
    return ctx.sync();
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

