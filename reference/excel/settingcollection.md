# <a name="settingcollection-object-javascript-api-for-excel"></a>SettingCollection オブジェクト (JavaScript API for Excel)

ブックの一部であるワークシート オブジェクトのコレクションを表します。

## <a name="properties"></a>プロパティ

| プロパティ       | 型    |説明| 要件セット|
|:---------------|:--------|:----------|:----|
|items|[Setting[]](setting.md)|設定オブジェクトのコレクション。読み取り専用です。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|

_プロパティのアクセスの[例を参照してください。](#property-access-examples)_

## <a name="relationships"></a>関係
なし


## <a name="methods"></a>メソッド

| メソッド           | 戻り値の型    |説明| 要件セット|
|:---------------|:--------|:----------|:----|
|[add(key: string, value: (any)[])](#addkey-string-value-any)|[Setting](setting.md)|指定した設定をブックに設定または追加します。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getCount()](#getcount)|int|コレクション内にある Setting の数を取得します。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getItem(key: string)](#getitemkey-string)|[Setting](setting.md)|キーから Setting エントリを取得します。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemOrNullObject(key: string)](#getitemornullobjectkey-string)|[Setting](setting.md)|キーから Setting エントリを取得します。Setting が存在しない場合は null オブジェクトを返します。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>メソッドの詳細


### <a name="addkey-string-value-any"></a>add(key: string, value: (any)[])
指定した設定をブックに設定または追加します。

#### <a name="syntax"></a>構文
```js
settingCollectionObject.add(key, value);
```

#### <a name="parameters"></a>パラメーター
| パラメーター       | 型    |説明|
|:---------------|:--------|:----------|:---|
|Key|string|新しい設定のキーです。|
|value|(任意)[]|新しい設定の値です。|

#### <a name="returns"></a>戻り値
[Setting](setting.md)

### <a name="getcount"></a>getCount()
コレクション内にある Setting の数を取得します。

#### <a name="syntax"></a>構文
```js
settingCollectionObject.getCount();
```

#### <a name="parameters"></a>パラメーター
なし

#### <a name="returns"></a>戻り値
int

### <a name="getitemkey-string"></a>getItem(key: string)
キーから Setting エントリを取得します。

#### <a name="syntax"></a>構文
```js
settingCollectionObject.getItem(key);
```

#### <a name="parameters"></a>パラメーター
| パラメーター       | 型    |説明|
|:---------------|:--------|:----------|:---|
|Key|string|設定のキーです。|

#### <a name="returns"></a>戻り値
[Setting](setting.md)

### <a name="getitemornullobjectkey-string"></a>getItemOrNullObject(key: string)
キーから Setting エントリを取得します。Setting が存在しない場合は null オブジェクトを返します。

#### <a name="syntax"></a>構文
```js
settingCollectionObject.getItemOrNullObject(key);
```

#### <a name="parameters"></a>パラメーター
| パラメーター       | 型    |説明|
|:---------------|:--------|:----------|:---|
|Key|string|設定のキーです。|

#### <a name="returns"></a>戻り値
[Setting](setting.md)
