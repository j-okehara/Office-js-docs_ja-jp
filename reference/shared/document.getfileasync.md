
# <a name="document.getfileasync-method"></a>Document.getFileAsync メソッド
ドキュメント ファイル全体を、最大で 4194304 バイト (4 MB) のスライスに分割して返します。iOS 用アドインの場合は、最大 65536 バイト (64KB) のファイル スライスがサポートされます。許可されている制限を超えてスライス サイズを指定すると、"内部エラー" が発生しますのでご注意ください。 

|||
|:-----|:-----|
|**ホスト:**|Excel、PowerPoint、Word|
|**[要件セット](../../docs/overview/specify-office-hosts-and-api-requirements.md)に指定できるもの**|ファイル|
|**ファイルの最終変更**|1.1|

```js
Office.context.document.getFileAsync(fileType [, options], callback);
```


## <a name="parameters"></a>パラメーター



|**名前**|**型**|**説明**|**サポートのメモ**|
|:-----|:-----|:-----|:-----|
| _fileType_|[FileType](../../reference/shared/filetype-enumeration.md)|ファイルが返される形式を指定します。必須。<br/><table><tr><th>Host</th><th>サポート対象 fileType</th></tr><tr><td>Excel Online</td><td>Office.FileType.Compressed</td></tr><tr><td>Windows デスクトップの PowerPoint</td><td>Office.FileType.Compressed、Office.FileType.Pdf</td></tr><tr><td>Windows デスクトップの Word、MAC および iPad</td><td>Office.FileType.Compressed、Office.FileType.Pdf、Office.FileType.Text</td></tr><tr><td>Word Online</td><td>Office.FileType.Compressed、Office.FileType.Pdf、Office.FileType.Text</td></tr><tr><td>PowerPoint Online</td><td>Office.FileType.Compressed、Office.FileType.Pdf</td></tr></table>|**変更対象:** 1.1。「[サポート履歴](#support-history)」をご覧ください|
| _options_|**object**|次の[オプションのパラメーター](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)のいずれかを指定します||
| _sliceSize_|**number**|目的のスライス サイズをバイト単位で指定します。最大は 4194304 バイト (4 MB) です。指定しない場合は、既定のスライス サイズである 4194304 バイト (4 MB) が使用されます。 ||
| _asyncContext_|**array**、**boolean**、**null**、**number**、**object**、**string**、または **undefined**|変更されずに **AsyncResult** オブジェクトで返される任意の型のユーザー定義項目。||
| _callback_|**object**|コールバックが戻るときに呼び出される関数で、唯一のパラメーターは **AsyncResult** 型です。||

## <a name="callback-value"></a>コールバック値

_callback_ パラメーターに渡した関数が実行されると、[AsyncResult](../../reference/shared/asyncresult.md) オブジェクトを受け取ります。このオブジェクトには、コールバック関数の唯一のパラメーターからアクセスできます。

**getFileAsync** メソッドに渡されたコールバック関数で、**AsyncResult** オブジェクトのプロパティを使用して次の情報を戻せます。



|**プロパティ**|**使用目的**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|[File](../../reference/shared/file.md) オブジェクトにアクセスします。|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|操作の成功または失敗を判断します。|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|操作が失敗した場合、エラーに関する情報を提供する [Error](../../reference/shared/error.md) オブジェクトにアクセスします。|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|ユーザー定義の **object** または値を _asyncContext_ パラメーターとして渡した場合、そのオブジェクトまたは値にアクセスします。|

## <a name="remarks"></a>注釈

Office for iOS 以外の Office ホスト アプリケーションで実行するアドインの場合、**getFileAsync** メソッドは最大 4194304 バイト (4MB) にスライスしたファイルの取得をサポートします。iOS アプリの Office で実行するアドインの場合は、**getFileAsync** メソッドは最大 65536 バイト (64KB) にスライスしたファイルの取得をサポートします。

_fileType_ パラメーターは、次の列挙型またはテキスト値を使用して指定できます。


**FileType 列挙型**


|**列挙**|**値**|**説明**|
|:-----|:-----|:-----|
|"compressed"|ドキュメント全体 (.docx、.pptx、.xslx) を Office Open XML (OOXML) 形式でバイト配列として返します。|Office.FileType.Pdf|
|"pdf"|PDF 形式のドキュメント全体をバイト配列として返します。|Office.FileType.Text|
|Office.FileType.Text|"text"|ドキュメントのテキストのみを  **string** として返します。 |
2 つを超えるドキュメントがメモリに存在する場合、**getFileAsync** 操作は失敗します。ファイルを使い終わったら、[File.closeAsync](../../reference/shared/file.closeasync.md) メソッドを使用してファイルを閉じてください。


## <a name="example---get-a-document-in-office-open-xml-("compressed")-format"></a>例 - Office Open XML ("圧縮") 形式でドキュメントを取得する

次の使用例では、Office Open XML ("圧縮") 形式のドキュメントを 65536 バイト (64KB) のスライスで取得しています。注意: この例での  `app.showNotification` の実装は、Office アドイン用の Visual Studio テンプレートに由来します。


```js
function getDocumentAsCompressed() {
    Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 65536 /*64 KB*/ }, 
        function (result) {
            if (result.status == "succeeded") {
            // If the getFileAsync call succeeded, then
            // result.value will return a valid File Object.
            var myFile = result.value;
            var sliceCount = myFile.sliceCount;
            var slicesReceived = 0, gotAllSlices = true, docdataSlices = [];
            app.showNotification("File size:" + myFile.size + " #Slices: " + sliceCount);

            // Get the file slices.
            getSliceAsync(myFile, 0, sliceCount, gotAllSlices, docdataSlices, slicesReceived);
            }
            else {
            app.showNotification("Error:", result.error.message);
            }
    });
}

function getSliceAsync(file, nextSlice, sliceCount, gotAllSlices, docdataSlices, slicesReceived) {
    file.getSliceAsync(nextSlice, function (sliceResult) {
        if (sliceResult.status == "succeeded") {
            if (!gotAllSlices) { // Failed to get all slices, no need to continue.
                return;
            }

            // Got one slice, store it in a temporary array.
            // (Or you can do something else, such as
            // send it to a third-party server.)
            docdataSlices[sliceResult.value.index] = sliceResult.value.data;
            if (++slicesReceived == sliceCount) {
               // All slices have been received.
               file.closeAsync();
               onGotAllSlices(docdataSlices);
            }
            else {
                getSliceAsync(file, ++nextSlice, sliceCount, gotAllSlices, docdataSlices, slicesReceived);
            }
        }
            else {
                gotAllSlices = false;
                file.closeAsync();
                app.showNotification("getSliceAsync Error:", sliceResult.error.message);
            }
    });
}

function onGotAllSlices(docdataSlices) {
    var docdata = [];
    for (var i = 0; i < docdataSlices.length; i++) {
        docdata = docdata.concat(docdataSlices[i]);
    }

    var fileContent = new String();
    for (var j = 0; j < docdata.length; j++) {
        fileContent += String.fromCharCode(docdata[j]);
    }

    // Now all the file content is stored in 'fileContent' variable,
    // you can do something with it, such as print, fax...
}

```


## <a name="example---get-a-document-in-pdf-format"></a>例 - PDF 形式でドキュメントを取得する

次の例では、PDF 形式でドキュメントを取得します。


```js
Office.context.document.getFileAsync(Office.FileType.Pdf,
    function(result) {
        if (result.status == "succeeded") {
            var myFile = result.value;
            var sliceCount = myFile.sliceCount;
            app.showNotification("File size:" + myFile.size + " #Slices: " + sliceCount);
            // Now, you can call getSliceAsync to download the files, as described in the previous code segment (compressed format).
            
            myFile.closeAsync();
        }
        else {
            app.showNotification("Error:", result.error.message);
        }
}
);


```


## <a name="support-details"></a>サポートの詳細


次の表で、大文字 Y は、このメソッドは、対応する Office ホスト アプリケーションでサポートされていることを示します。空のセルは、Office ホスト アプリケーションでこのメソッドをサポートしないことを示します。

Office ホスト アプリケーションとサーバーの要件の詳細については、「[Office アドインを実行するための要件](../../docs/overview/requirements-for-running-office-add-ins.md)」をご覧ください。


**サポートされるホスト (プラットフォーム別)**


||**Windows デスクトップ版 Office**|**Office Online (ブラウザー)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**||Y||
|**PowerPoint**|Y|Y|Y|
|**Word**|Y|Y|Y|

|||
|:-----|:-----|
|**要件セットに指定できるもの**|ファイル|
|**最小限のアクセス許可レベル**|[ReadAllDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**アドインの種類**|コンテンツ、作業ウィンドウ|
|**ライブラリ**|Office.js|
|**名前空間**|Office|

## <a name="support-history"></a>サポート履歴


|**バージョン**|**変更内容**|
|:-----|:-----|
|1.1| PowerPoint Online で、**fileType** パラメーターとして _Office.FileType.Pdf_ のサポートが追加されました。|
|1.1| PowerPoint Online で、**fileType** パラメーターとして _Office.FileType.Compressed_ のサポートが追加されました。|
|1.1| Word Online で、_fileType_ パラメーターとして **Office.FileType.Text** のサポートが追加されました。|
|1.1| Excel Online で、_fileType_ パラメーターとして **Office.FileType.Compressed** のサポートが追加されました。|
|1.1| Word Online で、_fileType_ パラメーターとして **Office.FileType.Compressed** および **Office.FileType.Pdf** のサポートが追加されました。|
|1.1|Office for iPad 用の PowerPoint および Word で、_fileType_ パラメーターとしてすべての **FileType** 値のサポートが追加されました。|
|1.1|Windows デスクトップ用の Word および PowerPoint で、_fileType_ パラメーターとして **Office.FileType.Pdf** のサポートが追加されました。|
|1.0|導入|
