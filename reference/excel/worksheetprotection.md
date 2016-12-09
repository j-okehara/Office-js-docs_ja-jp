# <a name="worksheetprotection-object-javascript-api-for-excel"></a>WorksheetProtection オブジェクト (JavaScript API for Excel)

シート オブジェクトの保護を表します。

## <a name="properties"></a>プロパティ

| プロパティ     | 型   |説明| 要件セット|
|:---------------|:--------|:----------|:----|
|protected|bool|ワークシートが保護されているかどうかを示します。読み取り専用。読み取り専用です。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="relationships"></a>リレーションシップ
| リレーションシップ | 型   |説明| 要件セット|
|:---------------|:--------|:----------|:----|
|options|[WorksheetProtectionOptions](worksheetprotectionoptions.md)|シートの保護のオプション。読み取り専用です。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>メソッド

| メソッド           | 戻り値の型    |説明| 要件セット|
|:---------------|:--------|:----------|:----|
|[load(param: object)](#loadparam-object)|(非推奨)|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[protect(options:WorksheetProtectionOptions)](#protectoptions-worksheetprotectionoptions)|void|ワークシートを保護します。ワークシートが保護されている場合は失敗します。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[unprotect()](#unprotect)|void|ワークシートの保護を解除します。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>メソッドの詳細


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

### <a name="protectoptions-worksheetprotectionoptions"></a>protect(options:WorksheetProtectionOptions)
ワークシートを保護します。ワークシートが保護されている場合は失敗します。

#### <a name="syntax"></a>構文
```js
worksheetProtectionObject.protect(options);
```

#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|:---|
|オプション|WorksheetProtectionOptions|省略可能。シートの保護のオプション。|

#### <a name="returns"></a>戻り値
void

#### <a name="examples"></a>例
```js
Excel.run(function (ctx) { 
    var sheet = ctx.workbook.worksheets.getItem("Sheet1");
    var range = sheet.getRange("A1:B3").format.protection.locked = false;
    sheet.protection.protect({allowInsertRows:true});
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});

```
### <a name="unprotect"></a>unprotect()
ワークシートの保護を解除します。

#### <a name="syntax"></a>構文
```js
worksheetProtectionObject.unprotect();
```

#### <a name="parameters"></a>パラメーター
なし

#### <a name="returns"></a>戻り値
void
