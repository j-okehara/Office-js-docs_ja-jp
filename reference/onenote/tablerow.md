# <a name="tablerow-object-(javascript-api-for-onenote)"></a>TableRow オブジェクト (JavaScript API for OneNote)

_適用対象:OneNote Online_  


表の行を表します。

## <a name="properties"></a>プロパティ

| プロパティ     | 型   |説明|フィードバック|
|:---------------|:--------|:----------|:-------|
|cellCount|int|行のセルの数を取得します。読み取り専用です。|[移動](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableRow-cellCount)|
|id|string|行の ID を取得します。読み取り専用です。|[移動](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableRow-id)|
|rowIndex|int|親テーブル内の行のインデックスを取得します。読み取り専用です。|[移動](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableRow-rowIndex)|

_プロパティのアクセスの[例を参照してください。](#property-access-examples)_

## <a name="relationships"></a>関係
| リレーションシップ | 型   |説明| フィードバック|
|:---------------|:--------|:----------|:-------|
|cells|[TableCellCollection](tablecellcollection.md)|行のセルを取得します。読み取り専用です。|[移動](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableRow-cells)|
|parentTable|[テーブル](table.md)|親テーブルを取得します。読み取り専用です。|[移動](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableRow-parentTable)|

## <a name="methods"></a>メソッド

| メソッド           | 戻り値の型    |説明| フィードバック|
|:---------------|:--------|:----------|:-------|
|[clear()](#clear)|void|行の内容をクリアします。|[移動](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableRow-clear)|
|[insertRowAsSibling(insertLocation: string, values: string[])](#insertrowassiblinginsertlocation-string-values-string)|[TableRow](tablerow.md)|現在の行の前後に行を挿入します。|[移動](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableRow-insertRowAsSibling)|
|[load(param: object)](#loadparam-object)|(非推奨)|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|[移動](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableRow-load)|
|[setShadingColor(colorCode: string)](#setshadingcolorcolorcode-string)|void|行のすべてのセルの網かけの色を設定します。|[移動](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableRow-setShadingColor)|

## <a name="method-details"></a>メソッドの詳細


### <a name="clear()"></a>clear()
行の内容をクリアします。

#### <a name="syntax"></a>構文
```js
tableRowObject.clear();
```

#### <a name="parameters"></a>パラメーター
なし

#### <a name="returns"></a>戻り値
void

### <a name="insertrowassibling(insertlocation:-string,-values:-string[])"></a>insertRowAsSibling(insertLocation: string, values: string[])
現在の行の前後に行を挿入します。

#### <a name="syntax"></a>構文
```js
tableRowObject.insertRowAsSibling(insertLocation, values);
```

#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|insertLocation|string|現在の行からの相対位置で、新しい行を挿入する場所。使用可能な値は次のとおりです。Before、After|
|values|string[]|省略可能。配列として指定された、新しい行に挿入する文字列。現在の行内のセルよりも多くのセル数にすることはできません。省略可能。|

#### <a name="returns"></a>戻り値
[TableRow](tablerow.md)

#### <a name="examples"></a>例
```js
OneNote.run(function(ctx) {
    var app = ctx.application;
    var outline = app.getActiveOutline();
    
    // Queue a command to load outline.paragraphs and their types.
    ctx.load(outline, "paragraphs, paragraphs/type");
    
    // Run the queued commands, and return a promise to indicate task completion.
    return ctx.sync().then(function () {
        var paragraphs = outline.paragraphs;
        
        // for each table, get table rows.
        for (var i = 0; i < paragraphs.items.length; i++) {
            var paragraph = paragraphs.items[i];
            if (paragraph.type == "Table") {
                var table = paragraph.table;
                
                // Queue a command to load table.rows.
                ctx.load(table, "rows");
                
                // Run the queued commands
                return ctx.sync().then(function() {
                    var rows = table.rows;
                    rows.items[1].insertRowAsSibling("Before", ["cell0", "cell1"]);
                    return ctx.sync();
                });
            }
        }
    })
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

### <a name="load(param:-object)"></a>load(param: object)
JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。

#### <a name="syntax"></a>構文
```js
object.load(param);
```

#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|param|object|省略可能。パラメーターとリレーションシップ名を、区切られた文字列または 1 つの配列として受け入れます。あるいは、[loadOption](loadoption.md) オブジェクトを提供します。|

#### <a name="returns"></a>戻り値
void

### <a name="setshadingcolor(colorcode:-string)"></a>setShadingColor(colorCode: string)
行のすべてのセルの網かけの色を設定します。

#### <a name="syntax"></a>構文
```js
tableRowObject.setShadingColor(colorCode);
```

#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|colorCode|string|セルに設定する色コード /param|

#### <a name="returns"></a>戻り値
void
### <a name="property-access-examples"></a>プロパティのアクセスの例
**id、cellCount、rowIndex**
```js
OneNote.run(function(ctx) {
    var app = ctx.application;
    var outline = app.getActiveOutline();
    
    // Queue a command to load outline.paragraphs and their types.
    ctx.load(outline, "paragraphs, paragraphs/type");
    
    // Run the queued commands, and return a promise to indicate task completion.
    return ctx.sync().then(function () {
        var paragraphs = outline.paragraphs;
        
        // for each table, get table rows.
        for (var i = 0; i < paragraphs.items.length; i++) {
            var paragraph = paragraphs.items[i];
            if (paragraph.type == "Table") {
                var table = paragraph.table;
                
                // Queue a command to load table.rows.
                ctx.load(table, "rows");
                return ctx.sync().then(function() {
                    var rows = table.rows;
                    
                    // for each table row, log cell count and row index.
                    for (var i = 0; i < rows.items.length; i++) {
                        console.log("Row " + i + " Id: " + rows.items[i].id);
                        console.log("Row " + i + " Cell Count: " + rows.items[i].cellCount);
                        console.log("Row " + i + " Row Index: " + rows.items[i].rowIndex);
                    }
                    return ctx.sync();
                });
            }
        }
    })
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

**parentTable、cells**
```js
OneNote.run(function(ctx) {
    var app = ctx.application;
    var outline = app.getActiveOutline();
    
    // Queue a command to load outline.paragraphs and their types.
    ctx.load(outline, "paragraphs, paragraphs/type");
    
    // Run the queued commands, and return a promise to indicate task completion.
    return ctx.sync().then(function () {
        var paragraphs = outline.paragraphs;
        
        // for each table, get table rows.
        for (var i = 0; i < paragraphs.items.length; i++) {
            var paragraph = paragraphs.items[i];
            if (paragraph.type == "Table") {
                var table = paragraph.table;
                
                // Queue a command to load parentTable and cells of each row in the table.
                ctx.load(table, "rows/parentTable, rows/cells");
                return ctx.sync().then(function() {
                    var rows = table.rows;
                    
                    // for each row, log parentTable and cells
                    for (var i = 0; i < rows.items.length; i++) {
                        console.log("Row " + i + " Parent Table Id: " + rows.items[i].parentTable.id);
                        var cells = rows.items[i].cells;
                        for (var j = 0 ; j < cells.items.length; j++) {
                            console.log("Row " + i + " Cell " + j + " Id: " + cells.items[j].id);
                        }
                    }
                    return ctx.sync();
                });
            }
        }
    })
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

