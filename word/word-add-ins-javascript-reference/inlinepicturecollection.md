# InlinePictureCollection オブジェクト (JavaScript API for Word)

[inlinePicture](inlinepicture.md) オブジェクトのコレクションが含まれます。

_適用対象:Word 2016、Word for iPad、Word for Mac_

## プロパティ
| プロパティ   | 型|説明
|:---------------|:--------|:----------|
|Items|[InlinePicture[]](inlinepicture.md)|inlinePicture オブジェクトのコレクション。読み取り専用です。|

_プロパティのアクセスの[例](#property-access-examples)を参照してください。_

## 関係
なし


## メソッド

| メソッド   | 戻り値の型|説明|
|:---------------|:--------|:----------|
|[load(param: object)](#loadparam-object)|void|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|

## メソッドの詳細

### load(param: object)
JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。

#### 構文
```js
object.load(param);
```

#### パラメータ
| パラメーター   | 型|説明|
|:---------------|:--------|:----------|
|param|object|省略可能。パラメーターとリレーションシップ名を、区切られた文字列または 1 つの配列として受け入れます。あるいは、[loadOption](loadoption.md) オブジェクトを提供します。|

#### 戻り値
void

## サポートの詳細

実行時のチェックで[要件セット](https://msdn.microsoft.com/EN-US/library/office/mt590206.aspx)を使用して、アプリケーションが Word のホスト バージョンによってサポートされていることを確かめます。Office ホスト アプリケーションとサーバーの要件の詳細については、「[Office アドインを実行するための要件](https://msdn.microsoft.com/EN-US/library/office/dn833104.aspx)」を参照してください。 

