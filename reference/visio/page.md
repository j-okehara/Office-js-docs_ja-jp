# <a name="page-object-javascript-api-for-visio"></a>Page オブジェクト (JavaScript API for Visio)

適用対象:_Visio Online_
>**注:**Visio JavaScript API は、現在プレビューの段階であり、変更される可能性があります。Visio JavaScript API は、運用環境での使用は現在サポートされていません。

ページ クラスを表します。

## <a name="properties"></a>プロパティ

| プロパティ       | 型    |説明| フィードバック|
|:---------------|:--------|:----------|:---|
|index|int|ページのインデックス。読み取り専用です。|[移動](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-page-index)|
|isBackground|bool|ページが背景ページかどうか。読み取り専用です。|[移動](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-page-isBackground)|
|name|string|ページの名前。読み取り専用です。|[移動](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-page-name)|

## <a name="relationships"></a>関係
| リレーションシップ | 型    |説明| フィードバック|
|:---------------|:--------|:----------|:---|
|図形|[ShapeCollection](shapecollection.md)|ページ内の図形。読み取り専用です。|[移動](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-page-shapes)|
|ビュー|[PageView](pageview.md)|ページのビューを返します。読み取り専用です。|[移動](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-page-view)|

## <a name="methods"></a>メソッド

| メソッド           | 戻り値の型    |説明| フィードバック|
|:---------------|:--------|:----------|:---|
|[activate()](#activate)|void|ドキュメントのアクティブ ページとして設定します。|[移動](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-page-activate)|
|[load(param: object)](#loadparam-object)|(非推奨)|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|[移動](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-page-load)|

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
