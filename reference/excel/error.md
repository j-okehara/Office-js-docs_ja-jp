# OfficeExtension.Error オブジェクト (JavaScript API for Excel)

Excel JavaScript API の使用時に発生するエラーを表します。

_適用対象: Excel 2016、Excel Online、Excel for iOS、Office 2016_

## プロパティ
| プロパティ     | 型   |説明
|:---------------|:--------|:----------|
|code|string|エラーの種類を示す値を取得します。次の値をとることができます。"AccessDenied"、"ActivityLimitReached"、"BadPassword"、"GeneralException"、"InsertDeleteConflict"、"InvalidArgument"、"InvalidBinding"、"InvalidOperation"、"InvalidReference"、"InvalidSelection"、"ItemAlreadyExists"、"ItemNotFound"、"NotImplemented"、"UnsupportedOperation"。 |
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
