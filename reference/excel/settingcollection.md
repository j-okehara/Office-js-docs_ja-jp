# <a name="settingcollection-object-javascript-api-for-excel"></a>SettingCollection オブジェクト (JavaScript API for Excel)

ブックの一部であるワークシート オブジェクトのコレクションを表します。

## <a name="properties"></a>プロパティ

| プロパティ     | 型   |説明| 要件セット|
|:---------------|:--------|:----------|:----|
|items|[Setting[]](setting.md)|設定オブジェクトのコレクション。読み取り専用です。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|

_プロパティのアクセスの[例を参照してください。](#property-access-examples)_

## <a name="relationships"></a>関係
なし


## <a name="methods"></a>メソッド

| メソッド           | 戻り値の型    |説明| 要件セット|
|:---------------|:--------|:----------|:----|
|[getItem(key: string)](#getitemkey-string)|[Setting](setting.md)|キーから Setting エントリを取得します。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemOrNull(key: string)](#getitemornullkey-string)|[Setting](setting.md)|キーから Setting エントリを取得します。Setting が存在しない場合、返されたオブジェクトの isNull プロパティは true になります。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|[load(param: object)](#loadparam-object)|(非推奨)|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[set(key: string, value: string)](#setkey-string-value-string)|[Setting](setting.md)|指定した設定をブックに設定または追加します。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>メソッドの詳細


### <a name="getitemkey-string"></a>getItem(key: string)
キーから Setting エントリを取得します。

#### <a name="syntax"></a>構文
```js
settingCollectionObject.getItem(key);
```

#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|:---|
|Key|string|設定のキーです。|

#### <a name="returns"></a>戻り値
[Setting](setting.md)

### <a name="getitemornullkey-string"></a>getItemOrNull(key: string)
キーから Setting エントリを取得します。Setting が存在しない場合、返されたオブジェクトの isNull プロパティは true になります。

#### <a name="syntax"></a>構文
```js
settingCollectionObject.getItemOrNull(key);
```

#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|:---|
|Key|string|設定のキーです。|

#### <a name="returns"></a>戻り値
[Setting](setting.md)

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

### <a name="setkey-string-value-string"></a>set(key: string, value: string)
指定した設定をブックに設定または追加します。

#### <a name="syntax"></a>構文
```js
settingCollectionObject.set(key, value);
```

#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|:---|
|Key|string|新しい設定のキーです。|
|value|string|新しい設定の値です。|

#### <a name="returns"></a>戻り値
[Setting](setting.md)
