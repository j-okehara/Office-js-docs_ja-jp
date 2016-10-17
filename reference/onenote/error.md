# <a name="officeextension.error-object-(javascript-api-for-onenote)"></a>OfficeExtension.Error オブジェクト (JavaScript API for OneNote)

_適用対象:OneNote Online_


OneNote JavaScript API の使用時に発生するエラーを表します。

## <a name="properties"></a>プロパティ
| プロパティ     | 型   |説明
|:---------------|:--------|:----------|
|code|string|エラーの種類を示す値を取得します。有効な値は、"InvalidArgument"、"GeneralException"、"ItemNotFound"、または "UnsupportedOperationForObjectType" です。 |
|debugInfo|string|エラーが発生したときに何が起こったかを示す値を取得します。この値は、開発中またはデバッグ中のみに使用することが想定されています。  |
|message |string| エラー コードに対応する、人間が判読できるローカライズされた文字列を取得します。|
|name |string| 常に "OfficeExtension.Error" である値を取得します。 |
|traceMessages |string[]| Context.trace(); を使用して設定するインストルメンテーション メッセージに対応する値の配列を取得します。 |

## <a name="methods"></a>メソッド

| メソッド           | 戻り値の型    |説明|
|:---------------|:--------|:----------|
|[toString()](#tostring)|string|次の形式でエラー コードとメッセージの値を返します: "{0}: {1}", コード, メッセージ。|

## <a name="method-details"></a>メソッドの詳細

### <a name="tostring()"></a>toString()
次の形式でエラー コードとメッセージの値を返します: "{0}: {1}", コード, メッセージ。

#### <a name="syntax"></a>構文
```js
error.toString()
```

#### <a name="parameters"></a>パラメーター
なし。

#### <a name="returns"></a>戻り値
string
