# <a name="documentview-object-javascript-api-for-visio"></a>DocumentView オブジェクト (JavaScript API for Visio)

適用対象:_Visio Online_

DocumentView クラスを表します。

## <a name="properties"></a>プロパティ

| プロパティ       | 型    |説明|
|:---------------|:--------|:----------|
|disableHyperlinks|bool|ハイパーリンクを無効にします。|
|disablePan|bool|パンを無効にします。|
|disableZoom|bool|ズームを無効にします。|
|hideDiagramBoundry|bool|図の境界を非表示にします。|

## <a name="relationships"></a>関係
なし


## <a name="methods"></a>メソッド

| メソッド           | 戻り値の型    |説明|
|:---------------|:--------|:----------|
|[load(param: object)](#loadparam-object)|(非推奨)|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|

## <a name="method-details"></a>メソッドの詳細


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
