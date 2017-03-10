# <a name="shapecollection-object-javascript-api-for-visio"></a>ShapeCollection オブジェクト (JavaScript API for Visio)

適用対象:_Visio Online_

図形コレクションを表します。

## <a name="properties"></a>プロパティ

| プロパティ       | 型    |説明|
|:---------------|:--------|:----------|
|items|[Shape[]](shape.md)|図形オブジェクトのコレクション。読み取り専用です。|

## <a name="relationships"></a>関係
なし


## <a name="methods"></a>メソッド

| メソッド           | 戻り値の型    |説明|
|:---------------|:--------|:----------|
|[getCount()](#getcount)|int|コレクション内の図形の数を取得します。|
|[getItem(key: number または string)](#getitemkey-number-or-string)|[図形](shape.md)|そのキー (名前または ID) を使用して図形を取得します。|
|[load(param: object)](#loadparam-object)|(非推奨)|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|

## <a name="method-details"></a>メソッドの詳細


### <a name="getcount"></a>getCount()
コレクション内の図形の数を取得します。

#### <a name="syntax"></a>構文
```js
shapeCollectionObject.getCount();
```

#### <a name="parameters"></a>パラメーター
なし

#### <a name="returns"></a>戻り値
int

#### <a name="examples"></a>例
```js
Visio.run(function (ctx) { 
    var activePage = ctx.document.getActivePage();
    var numShapesActivePage = activePage.shapes.getCount();
    return ctx.sync().then(function () {
        console.log("Shapes Count: " + numShapesActivePage.value);
    });

}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### <a name="getitemkey-number-or-string"></a>getItem(key: number または string)
そのキー (名前または ID) を使用して図形を取得します。

#### <a name="syntax"></a>構文
```js
shapeCollectionObject.getItem(key);
```

#### <a name="parameters"></a>パラメーター
| パラメーター       | 型    |説明|
|:---------------|:--------|:----------|:---|
|Key|number または string|キーは、取得する図形の名前またはインデックスです。|

#### <a name="returns"></a>戻り値
[図形](shape.md)

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
