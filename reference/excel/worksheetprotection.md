# WorksheetProtection オブジェクト (JavaScript API for Excel)

_適用対象: Excel 2016、Excel Online、Excel for iOS、Office 2016_

シート オブジェクトの保護を表します。

## プロパティ

| プロパティ     | 型   |説明
|:---------------|:--------|:----------|
|protected|bool|ワークシートが保護されているかどうかを示します。読み取り専用。|

## リレーションシップ
| リレーションシップ | 型   |説明|
|:---------------|:--------|:----------|
|オプション|[WorksheetProtectionOptions](worksheetprotectionoptions.md)|シートの保護のオプション。読み取り専用。|

## メソッド

| メソッド           | 戻り値の型    |説明|
|:---------------|:--------|:----------|
|[load(param: object)](#loadparam-object)|void|プロキシ オブジェクトにシートの保護の詳細を設定します。|
|[protect(options:WorksheetProtectionOptions)](#protectoptions-worksheetprotectionoptions)|void|ワークシートを保護します。ワークシートが保護されている場合はスローします。|
|[unprotect()](#unprotect)|void|ワークシートの保護を解除します。|

## メソッドの詳細


### load(param: object)
JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。

#### 構文
```js
object.load(param);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|param|object|省略可能。パラメーター名とリレーションシップ名を、区切られた文字列または 1 つの配列として受け入れます。あるいは、[loadOption](loadoption.md) オブジェクトを提供します。|

#### 戻り値
void

#### 例
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

### protect(options:WorksheetProtectionOptions)
オプションの保護ポリシーを使用してワークシートを保護します。ワークシートが保護されている場合は例外をスローします。 

オプションが指定されている場合は、個々のポリシーの有効/無効を切り替えられます。ポリシーが指定されていない場合、既定で有効になります。 

#### 構文
```js
worksheetProtectionObject.protect(options);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|オプション|WorksheetProtectionOptions|省略可能。シートの保護のオプション。|


#### 戻り値
void

#### 例
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
### unprotect()
ワークシートの保護を解除します。 

#### 構文
```js
worksheetProtectionObject.unprotect();
```

#### パラメーター
なし

#### 戻り値
void

#### 例
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