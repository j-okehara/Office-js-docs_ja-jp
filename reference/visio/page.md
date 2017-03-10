# <a name="page-object-javascript-api-for-visio"></a>Page オブジェクト (JavaScript API for Visio)

適用対象:_Visio Online_

Page クラスを表します。

## <a name="properties"></a>プロパティ

| プロパティ       | 型    |説明|
|:---------------|:--------|:----------|
|height|int|ページの高さを返します。読み取り専用です。|
|index|int|ページのインデックス。読み取り専用です。|
|isBackground|bool|ページが背景ページかどうか。読み取り専用です。|
|name|string|ページの名前。読み取り専用です。|
|width|int|ページの幅を返します。読み取り専用です。|

## <a name="relationships"></a>関係
| リレーションシップ | 型    |説明|
|:---------------|:--------|:----------|
|コメント|[CommentCollection](commentcollection.md)|Comments コレクションを返します。読み取り専用です。|
|図形|[ShapeCollection](shapecollection.md)|ページ内の図形。読み取り専用です。|
|ビュー|[PageView](pageview.md)|ページのビューを返します。読み取り専用です。|

## <a name="methods"></a>メソッド

| メソッド           | 戻り値の型    |説明|
|:---------------|:--------|:----------|
|[activate()](#activate)|void|ドキュメントのアクティブ ページとして設定します。|
|[load(param: object)](#loadparam-object)|(非推奨)|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|

## <a name="method-details"></a>メソッドの詳細


### <a name="activate"></a>activate()
ドキュメントのアクティブ ページとして設定します。

#### <a name="syntax"></a>構文
```js
pageObject.activate();
```

#### <a name="parameters"></a>パラメーター
なし

#### <a name="returns"></a>戻り値
void

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
