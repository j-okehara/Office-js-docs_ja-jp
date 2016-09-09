
# Document.setSelectedDataAsync メソッド
ドキュメント内の現在の選択範囲にデータを書き込みます。

|||
|:-----|:-----|
|**ホスト:** Access、Excel、PowerPoint、Project、Word、Word Online|**アドインの種類:** コンテンツ、作業ウィンドウ|
|**[要件セット](../../docs/overview/specify-office-hosts-and-api-requirements.md)に指定できるもの**|選択内容|
|**最終変更バージョン**|1.1|

```js
Office.context.document.setSelectedDataAsync(data [, options], callback(asyncResult));
```


## パラメーター

|**名前**|**型**|**説明**|**サポートのメモ**|
|:-----|:-----|:-----|:-----|
| _data_|data には次のいずれかのデータ型を使用できます。<ul><li><b>文字列</b> (Office.CoercionType.Text) - Excel、Excel Online、PowerPoint、PowerPoint Online、Word、Word Online のみに適用されます。</li><li><b>配列</b> の配列 (Office.CoercionType.Matrix) - Excel、Word、Word Online のみに適用されます。</li><li>[TableData](../../reference/shared/tabledata.md) (Office.CoercionType.Table) - Access、Excel、Word、Word Online のみ</li><li><b>HTML</b> (Office.CoercionType.Html) - Word と Word Online のみに適用されます。</li><li><b>Office Open XML</b> (Office.CoercionType.Ooxml) - Word と Word Online のみに適用されます。</li><li><b>base64 エンコードのイメージ ストリーム</b> (Office.CoercionType.Image) - Excel、PowerPoint、Word、Word Online のみに適用されます。</li></ul>|現在の選択範囲に設定するデータ。 必須です。|**変更対象:** 1.1。Access 用コンテンツ アドインのサポートには、 **Selection** 要件セット 1.1 以上が必要です。イメージ データの設定のサポートには、 **ImageCoercion** 要件セット 1.1 以降が必要です。アプリのアクティブ化のためにこれを設定するには、以下を使用します。<br/><br/>`<Requirements>`<br/>&nbsp;&nbsp;`<Sets DefaultMinVersion="1.1">`<br/>&nbsp;&nbsp;&nbsp;&nbsp;`<Set Name="ImageCoercion"/>`<br/>&nbsp;&nbsp;`</Sets>`<br/>`</Requirements>`<br/><br/>ImageCoercion 機能の実行時の検出は、次のコードで実行できます。<br/><br/>`if (Office.context.requirements.isSetSupported('ImageCoercion', '1.1')) {)) {`<br/>&nbsp;&nbsp;&nbsp;&nbsp;`// insertViaImageCoercion();`<br/>`} else {`<br/>&nbsp;&nbsp;&nbsp;&nbsp;`// insertViaOoxml();`<br/>`}`|
| _options_|**object**|[オプション パラメーター](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)のセットを指定します。 options オブジェクトには、オプションを設定するために次のプロパティを含めることができます。<br/><ul><li>coercionType (<b><a href="735eaab6-5e31-4bc2-add5-9d378900a31b.htm">CoercionType</a></b>) - 設定するデータに強制する型を指定します。 このオプションを指定しないと、既定の coercionType 値である Office.CoercionType.Text が使用されます。</li><li>tableOptions (<b>object</b> ) - 挿入された表の場合に、<a href="http://msdn.microsoft.com/library/46b05707-b350-41be-b6b8-311799c71a33(Office.15).aspx" target="_blank">表の書式設定オプション</a> (ヘッダー行、集計行、縞模様行など) を指定するキー/値ペアのリスト。 </li><li>cellFormat (<b>object</b> ) - 挿入された表の場合に、特定の範囲の列、行、またはセルを指定し、その範囲に適用する<a href="http://msdn.microsoft.com/library/46b05707-b350-41be-b6b8-311799c71a33(Office.15).aspx" target="_blank">セルの書式設定</a>を指定するキー/値ペアのリスト。 </li><li>imageLeft (<b>数値</b> ) - このオプションは画像を挿入する場合に適用されます。PowerPoint のスライドの左端に関連した挿入位置と、Excel で現在選択されているセルとその位置の関係を示します。Word ではこの値は無視されます。この値はポイント単位で指定されます。</li><li>imageTop (<b>数値</b> ) - このオプションは画像を挿入する場合に適用されます。PowerPoint のスライドの上端に関連した挿入位置と、Excel で現在選択されているセルとその位置の関係を示します。Word ではこの値は無視されます。この値はポイント単位で指定されます。</li><li>imageWidth (<b>数値</b> ) - このオプションは画像を挿入する場合に適用されます。画像の幅を示します。このオプションだけを指定して、imageHeight を指定しない場合、画像は指定した幅の値に一致するように拡大/縮小されます。画像の幅と画像の高さの両方を指定すると、画像はそのサイズに変更されます。画像の幅と画像の高さのどちらも指定しない場合は、画像の既定のサイズと縦横比が使用されます。この値はポイント単位で指定されます。</li><li>imageHeight  (<b>数値</b> ) - このオプションは画像を挿入する場合に適用されます。画像の高さを示します。このオプションだけを指定して、imageWidth を指定しない場合、画像は指定した高さの値に一致するように拡大/縮小されます。画像の幅と画像の高さの両方を指定すると、画像はそのサイズに変更されます。画像の幅と画像の高さのどちらも指定しない場合は、画像の既定のサイズと縦横比が使用されます。この値はポイント単位で指定されます。</li><li>asyncContext (<b>object \| value</b>) - <a href="540c114f-0398-425c-baf3-7363f2f6bc47.htm">AsyncResult</a> オブジェクトの asyncContext プロパティで取得できるユーザー定義のオブジェクト。 これは、コールバックが名前付き関数の場合に、<b>AsyncResult</b> にオブジェクトまたは値を提供するために使用します。</li></ul>|_tableOptions_ オプションおよび _cellFormat_ オプションは、v1.1 に追加され、Excel 2013 と Excel Online でサポートされています。<br/><br/>_ImageLeft_ オプションと _ImageTop_ オプションは、Excel と PowerPoint でサポートされています。|
| _callback_|**object**|コールバックが戻るときに呼び出される関数で、唯一のパラメーターは  **AsyncResult** 型です。||

## コールバック値

_callback_ パラメーターに渡した関数が実行されると、[AsyncResult](../../reference/shared/asyncresult.md) オブジェクトを受け取ります。このオブジェクトには、コールバック関数の唯一のパラメーターからアクセスできます。

**setSelectedDataAsync** メソッドに渡されたコールバック関数内で、[AsyncResult.value](../../reference/shared/asyncresult.value.md) プロパティは常に **undefined** を返します。取得するオブジェクトまたはデータがないためです。


## 注釈

_data_ に渡される値には、現在の選択範囲に書き込まれるデータが含まれます。 値に応じて次のような処理が行われます。


-  **文字列:**プレーン テキスト、または **string** に強制的に変換できるその他の値が挿入されます。
    
    
    
    また Excel では、選択したセルに追加する有効な数式として _data_ を指定できます。 たとえば、_data_ を `"=SUM(A1:A5)"` と設定すると、指定の範囲内の値が集計されます。 ただし、バインドされたセルで数式を設定する場合、その後、バインドされたセルからは追加された数式 (または既存の数式) を読み取ることができません。 選択したセルで [Document.getSelectedDataAsync](../../reference/shared/document.getselecteddataasync.md) メソッドを呼び出してそのデータを読み取ると、このメソッドは (数式の結果である) セルに表示されたデータのみを返します。
    
-  **配列の配列 ("matrix"):** ヘッダーなしの表形式データが挿入されます。たとえば、3 行 2 列のデータを書き込むには、 `[["R1C1", "R1C2"], ["R2C1", "R2C2"], ["R3C1", "R3C2"]]` という配列を渡します。3 行 1 列のデータを書き込むには、 `[["R1C1"], ["R2C1"], ["R3C1"]]` という配列を渡します。
    
    
    
    Excel では、_data_ を有効な数式を含む、配列の配列として指定して、選択したセルに追加することができます。 たとえば、他に上書きされるデータがない場合、_data_ を `[["=SUM(A1:A5)","=AVERAGE(A1:A5)"]]` に設定すると、この 2 つの数式が選択範囲に追加されます。 数式を単一のセルで "テキスト" として設定する場合と同様に、設定後、追加した数式 (または既存の任意の数式) を読み取ることはできません。数式の結果を読み取ることができるだけです。
    
-  **[TableData](../../reference/shared/tabledata.md) オブジェクト:** ヘッダー付きのテーブルが挿入されます。
    
    
    
     **Note:** In Excel, if you specify formulas in the **TableData** object you pass for the _data_ parameter, you might not get the results you expect due to the "calculated columns" feature of Excel, which automatically duplicates formulas within a column. To work around this when you want to write _data_ that contains formulas to a selected table, try specifying the data as an array of arrays (instead of a **TableData** object), and specify the _coercionType_ as **Microsoft.Office.Matrix** or "matrix".
    
 **Application-specific behaviors**

さらに、選択範囲にデータを書き込むときには、次のアプリケーション固有の処理が適用されます。

 **Word の場合**


- 選択範囲がなく、挿入ポイントが有効な場所にある場合、指定された  _data_ は、次の規則に従って、挿入ポイントに挿入されます。
    
      - If  _data_ is a string, the specified text is inserted.
    
  - _data_ が配列の配列 ("matrix") または **TableData** オブジェクトの場合、Word の新しいテーブルが挿入されます。
    
  - _data_ が HTML の場合、指定された HTML が挿入されます。
    
     >**重要**: 挿入する HTML に無効な HTML が含まれている場合も、Word でエラーは発生しません。 HTML はできる限り挿入され、無効なデータは省略されます。
  - _data_ が Office Open XML の場合、指定した XML が挿入されます。
    
  - _data_ が base64 エンコードのイメージ ストリームの場合、指定した画像が挿入されます。
    
- 選択範囲がある場合は、上記と同じ規則に従って、指定された  _data_ で置き換えられます。
    
-  **画像の挿入**: 挿入された画像はインラインで配置されます。 **imageLeft** パラメーターと **imageTop** パラメーターは無視されます。画像の縦横比は常に固定されます。 **imageWidth** と **imageHeight** パラメーターのいずれか一方が指定されている場合、もう一方の値は、元の縦横比が維持されるように自動調整されます。
    
 **Excel の場合**


- 1 つのセルが選択されている場合は、次のような処理が行われます。
    
      - If  _data_ is a string, the specified text is inserted as the value of the current cell.
    
  - _data_ が配列の配列 ("matrix") の場合、指定された行と列が挿入されます (周囲のセルに含まれるデータが上書きされる場合は除く)。
    
  - _data_ が **TableData** オブジェクトの場合、指定された行とヘッダー付きの Excel の新しいテーブルが挿入されます (周囲のセルに含まれるデータが上書きされる場合は除く)。
    
- 複数のセルが選択され、形状が  _data_ の形状と一致しない場合、エラーが返されます。
    
- 複数のセルが選択され、選択範囲の形状が  _data_ の形状と一致する場合、選択されたセルの値は _data_ の値に基づいて更新されます。
    
-  **画像の挿入**: 挿入された画像は浮動になります。位置パラメーターの **imageLeft** と **imageTop** は、現在選択されているセルからの相対位置になります。 **imageLeft** と **imageTop** は負の値にすることもでき、その場合は、画像がワークシート内に収まるようにするために Excel によって再調整される可能性があります。画像の縦横比は、 **imageWidth** と **imageHeight** パラメーターの両方が指定されない限り固定されます。 **imageWidth** パラメーターと **imageHeight** パラメーターのいずれか一方のみが指定された場合、もう一方の値は、元の縦横比が維持されるように自動調整されます。
    
上記以外の場合は、エラーが返されます。

 **Excel Online の場合**

Excel に関する上記の説明に加えて、Excel Online にデータを書き込む場合に、次の制限が適用されます。 


- このメソッドに対する単一の呼び出しで、 _data_ パラメーターを使用してワークシートに書き込むセルの総数が 20,000 を超えることはできません。
    
- _cellFormat_ パラメーターに渡される _書式設定グループ_ の数が 100 を超えることはできません。1 つの書式設定グループは、指定のセル範囲に適用される書式設定のセットから構成されます。たとえば、次の呼び出しは、2 つの書式設定グループを _cellFormat_ に渡します。
    

```js
  Office.context.document.setSelectedDataAsync(
    {cellFormat:[{cells: {row: 1}, format: {fontColor: "yellow"}}, 
        {cells: {row: 3, column: 4}, format: {borderColor: "white", fontStyle: "bold"}}]}, 
    function (asyncResult){});
```

 **PowerPoint の場合**

挿入された画像は浮動になります。位置パラメーターの  **imageLeft** と **imageTop** は省略可能ですが、指定する場合は両方を指定する必要があります。1 つの値しか指定されない場合、それは無視されます。 **imageLeft** と **imageTop** は負の値にすることもでき、その場合は、画像がスライドの外に配置される可能性があります。オプションのパラメーターが指定されず、スライドにプレースホルダがある場合は、画像によってスライドのプレースホルダが置き換えられます。 画像の縦横比は、 **imageWidth** パラメーターと **imageHeight** パラメーターの両方が指定されない限り固定されます。 **imageWidth** パラメーターと **imageHeight** パラメーターのいずれか一方が指定されている場合、もう一方の値は、元の縦横比を維持するように自動調整されます。


## 例

次の例では、選択されたテキストまたはセルを "Hello World!" と設定します。エラーが発生した場合は、[error.message](../../reference/shared/error.message.md) プロパティの値を表示します。


```js
function writeText() {
    Office.context.document.setSelectedDataAsync("Hello World!",
        function (asyncResult) {
            var error = asyncResult.error;
            if (asyncResult.status === Office.AsyncResultStatus.Failed){
                 write(error.name + ": " + error.message);
            }
        });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```



省略可能な _coercionType_ パラメーターを指定すると、選択範囲に書き込むデータの種類を指定できます。 次の例では、データを 3 行 2 列の配列として書き込み、データ構造として _coercionType_ を `"matrix"` に設定します。エラーが発生した場合は、[error.message](../../reference/shared/error.message.md) プロパティの値を表示します。




```js
function writeMatrix() {
    Office.context.document.setSelectedDataAsync([["Red", "Rojo"], ["Green", "Verde"], ["Blue", "Azul"]], {coercionType: Office.CoercionType.Matrix}
        function (asyncResult) {
            var error = asyncResult.error;
            if (asyncResult.status === Office.AsyncResultStatus.Failed){
                write(error.name + ": " + error.message);
            }
        });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```



次の例では、データをヘッダーと 4 つの行で構成される 1 列の表として書き込み、データ構造として、_coercionType_ を `"table"` に設定します。エラーが発生した場合は、[error.message](../../reference/shared/error.message.md) プロパティの値を表示します。




```js
function writeTable() {
    // Build table.
    var myTable = new Office.TableData();
    myTable.headers = [["Cities"]];
    myTable.rows = [['Berlin'], ['Roma'], ['Tokyo'], ['Seattle']];

    // Write table.
    Office.context.document.setSelectedDataAsync(myTable, {coercionType: Office.CoercionType.Table},
        function (result) {
            var error = result.error
            if (result.status === Office.AsyncResultStatus.Failed) {
                write(error.name + ": " + error.message);
            }
    });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```



 Word で、選択範囲に HTML を書き込む場合は、次の例に示すように、 _coercionType_ パラメーターを `"html"` に指定します。この例では、HTML の `<b>` タグを使用して "Hello" を太字に設定します。




```js
function writeHtmlData() {
    Office.context.document.setSelectedDataAsync("<b>Hello</b> World!", {coercionType: Office.CoercionType.Html}, function (asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            write('Error: ' + asyncResult.error.message);
        }
    });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

Word、PowerPoint、Excel で選択範囲に画像を書き込むには、次の例に示すとおり、 _coercionType_ パラメーターを `"image"` と指定します。なお、Word では、imageLeft と imageTop は無視されます。




```js
function insertPictureAtSelection(base64EncodedImageStr) {

    Office.context.document.setSelectedDataAsync(base64EncodedImageStr, {
       coercionType: Office.CoercionType.Image,
       imageLeft: 50,
       imageTop: 50,
       imageWidth: 100,
       imageHeight: 100
       },
       function (asyncResult) {
           if (asyncResult.status === Office.AsyncResultStatus.Failed) {
               console.log("Action failed with error: " + asyncResult.error.message);
           }
       });
}
```


## サポートの詳細


次の表でチェック マーク (![チェック記号](../../images/mod_off15_checkmark.png)) は、このメソッドが対応する Office ホスト アプリケーションでサポートされていることを示します。 空のセルは、Office ホスト アプリケーションでこのメソッドをサポートしないことを示します。

Office ホスト アプリケーションとサーバーの要件の詳細については、「[Office アドインを実行するための要件](../../docs/overview/requirements-for-running-office-add-ins.md)」をご覧ください。


**サポートされるホスト (プラットフォーム別)**

||**Windows デスクトップ版 Office**|**Office Online (ブラウザー)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Access**|![チェック記号](../../images/mod_off15_checkmark.png)|||
|**Excel**|![チェック記号](../../images/mod_off15_checkmark.png)|![チェック マーク](../../images/mod_off15_checkmark.png)|![チェック記号](../../images/mod_off15_checkmark.png)|
|**PowerPoint**|![チェック記号](../../images/mod_off15_checkmark.png)|![チェック マーク](../../images/mod_off15_checkmark.png)|![チェック記号](../../images/mod_off15_checkmark.png)|
|**Word**|![チェック記号](../../images/mod_off15_checkmark.png)|![チェック マーク](../../images/mod_off15_checkmark.png)|![チェック マーク](../../images/mod_off15_checkmark.png)|


|||
|:-----|:-----|
|**要件セットに指定できるもの**|選択内容|
|**最小限のアクセス許可レベル**|[WriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**アプリの種類**|コンテンツ、作業ウィンドウ|
|**ライブラリ**|Office.js|
|**名前空間**|Office|

## サポート履歴




|**変更内容**|**1.1**|
|:-----|:-----|
|1.1|Word と Word Online では、base64 エンコードのイメージ ストリームとしてデータを書き込むためのサポートが追加されました。|
|1.1|Word Online では、配列の  _array_ および **TableData** (テーブル) として、 **data** の書き込みに関するサポートが追加されました。|
|1.1|Office for iPad の Excel、PowerPoint、および Word で、Windows デスクトップの Excel、PowerPoint、および Word と同じレベルのサポートが追加されました。|
|1.1|Word Online では、 _string_ (テキスト) としての **data** の書き込みに関するサポートが追加されました。|
|1.1|Excel 用アドインで、[テーブルを挿入するときの書式設定](../../docs/excel/format-tables-in-add-ins-for-excel.md)がサポートされるようになりました。 _tableOptions_ および _cellFormat_ パラメーター (省略可能) を使用します。|
|1.1|Access 用アドインでのテーブル データの書き込みのサポートが追加されました。|
|1.0|導入|
