# OfficeExtension.Error オブジェクト (JavaScript API for OneNote)

_適用対象:OneNote Online_


OneNote JavaScript API の使用時に発生するエラーを表します。

## プロパティ
| プロパティ     | 型   |説明
|:---------------|:--------|:----------|
|code|string|エラーの種類を示す値を取得します。 有効な値は、"InvalidArgument"、"GeneralException"、"ItemNotFound"、または "UnsupportedOperationForObjectType" です。 |
|debugInfo|string|エラーが発生したときに何が起こったかを示す値を取得します。この値は、開発中またはデバッグ中のみに使用することが想定されています。  |
|message |string| エラー コードに対応する、人間が判読できるローカライズされた文字列を取得します。|
|name |string| 常に "OfficeExtension.Error" である値を取得します。 |
|traceMessages |string[]| Context.trace(); を使用して設定するインストルメンテーション メッセージに対応する値の配列を取得します。 |

## メソッド

| メソッド           | 戻り値の型    |説明|
|:---------------|:--------|:----------|
|[toString()](#tostring)|string|次の形式でエラー コードとメッセージの値を返します: "{0}: {1}", コード, メッセージ。|

## メソッドの詳細

### toString()
次の形式でエラー コードとメッセージの値を返します: "{0}: {1}", コード, メッセージ。

#### 構文
```js
error.toString()
```

#### パラメーター
なし。

#### 戻り値
string
