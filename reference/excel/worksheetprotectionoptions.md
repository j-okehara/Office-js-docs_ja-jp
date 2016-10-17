# <a name="worksheetprotectionoptions-object-(javascript-api-for-excel)"></a>WorksheetProtectionOptions オブジェクト (JavaScript API for Excel)

_適用対象:Excel 2016、Excel Online、Excel for iOS、Office 2016_

シート保護のオプションを表します。

## <a name="properties"></a>プロパティ

| プロパティ     | 型   |説明
|:---------------|:--------|:----------|
|allowAutoFilter|bool|自動フィルター機能の使用を可能にするワークシート保護オプションを表します。|
|allowDeleteColumns|bool|列の削除を可能にするワークシート保護オプションを表します。|
|allowDeleteRows|bool|行の削除を可能にするワークシート保護オプションを表します。|
|allowFormatCells|bool|セルの書式設定を可能にするワークシート保護オプションを表します。|
|allowFormatColumns|bool|列の書式設定を可能にするワークシート保護オプションを表します。|
|allowFormatRows|bool|行の書式設定を可能にするワークシート保護オプションを表します。|
|allowInsertColumns|bool|列の挿入を可能にするワークシート保護オプションを表します。|
|allowInsertHyperlinks|bool|ハイパーリンクの挿入を可能にするワークシート保護オプションを表します。|
|allowInsertRows|bool|行の挿入を可能にするワークシート保護オプションを表します。|
|allowPivotTables|bool|ピボット テーブル機能の使用を可能にするワークシート保護オプションを表します。|
|allowSort|bool|並び替え機能の使用を可能にするワークシート保護オプションを表します。|

_プロパティのアクセスの[例を参照してください。](#examples)_

## <a name="relationships"></a>関係
なし


## <a name="methods"></a>メソッド

| メソッド           | 戻り値の型    |説明|
|:---------------|:--------|:----------|
|[load(param: object)](#loadparam-object)|void|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|

## <a name="method-details"></a>メソッドの詳細


### <a name="load(param:-object)"></a>load(param: object)
JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。

#### <a name="syntax"></a>構文
```js
object.load(param);
```

#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|param|object|省略可能。パラメーター名とリレーションシップ名を、区切られた文字列または 1 つの配列として受け入れます。あるいは、[loadOption](loadoption.md) オブジェクトを提供します。|

#### <a name="returns"></a>戻り値
void

#### <a name="examples"></a>例
この例は、作業中のワークシートの保護オプションを読み込みます。
```js
Excel.run(function (ctx) {
    var worksheet = ctx.workbook.worksheets.getActiveWorksheet();
    worksheet.protection.load();            
    return ctx.sync()
        .then(function () {
            console.log("Active worksheet's protection options: " + worksheet.protection.options);
        });
})
.catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```
