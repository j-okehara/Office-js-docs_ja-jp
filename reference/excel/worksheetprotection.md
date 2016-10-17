# <a name="worksheetprotection-object-(javascript-api-for-excel)"></a>WorksheetProtection オブジェクト (JavaScript API for Excel)

_適用対象:Excel 2016、Excel Online、Excel for iOS、Office 2016_

シート オブジェクトの保護を表します。

## <a name="properties"></a>プロパティ

| プロパティ     | 型   |説明
|:---------------|:--------|:----------|
|protected|bool|ワークシートが保護されているかどうかを示します。読み取り専用。|

## <a name="relationships"></a>リレーションシップ
| リレーションシップ | 型   |説明|
|:---------------|:--------|:----------|
|options|[WorksheetProtectionOptions](worksheetprotectionoptions.md)|シートの保護のオプション。読み取り専用。|

## <a name="methods"></a>メソッド

| メソッド           | 戻り値の型    |説明|
|:---------------|:--------|:----------|
|[load(param: object)](#loadparam-object)|void|プロキシ オブジェクトにシートの保護の詳細を設定します。|
|[protect(options:WorksheetProtectionOptions)](#protectoptions-worksheetprotectionoption)|void|ワークシートを保護します。ワークシートが保護されている場合はスローします。|
|[unprotect()](#unprotect)|void|ワークシートの保護を解除します。|

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
この例は、アクティブなワークシートの保護の詳細を読み込みます。
```js
Excel.run(function (ctx) {
    var worksheet = ctx.workbook.worksheets.getActiveWorksheet();
    worksheet.protection.load();            
    return ctx.sync()
        .then(function () {
            console.log("Active worksheet's protection status: " + worksheet.protection.protected);
        });
})
.catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

### <a name="protect(options:-worksheetprotectionoptions)"></a>protect(options:WorksheetProtectionOptions)
オプションの保護ポリシーを使用してワークシートを保護します。ワークシートが保護されている場合は例外をスローします。 

オプションが指定されている場合は、個々のポリシーの有効/無効を切り替えられます。ポリシーが指定されていない場合、既定で有効になります。 

#### <a name="syntax"></a>構文
```js
worksheetProtectionObject.protect(options);
```

#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
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
### <a name="unprotect()"></a>unprotect()
ワークシートの保護を解除します。 

#### <a name="syntax"></a>構文
```js
worksheetProtectionObject.unprotect();
```

#### <a name="parameters"></a>パラメーター
なし

#### <a name="returns"></a>戻り値
void

#### <a name="examples"></a>例
```js
Excel.run(function (ctx) { 
    var sheet = ctx.workbook.worksheets.getItem("Sheet1");  
    sheet.protection.unprotect();
    return ctx.sync(); 
}).catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```