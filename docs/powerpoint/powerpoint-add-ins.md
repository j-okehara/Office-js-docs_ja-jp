
# PowerPoint 用のコンテンツ アドインと作業ウィンドウ アドインを作成する

この記事のコード例は、PowerPoint コンテンツ アドインの開発におけるいくつかの基本的なタスクを示しています。情報を表示する場合、これらの例は Visual StudioOffice アドイン プロジェクト テンプレートに含まれる  `app.showNotification` 関数に依存しています。アドインの開発に Visual Studio を使用しない場合は、 `showNotification` 関数を独自のコードに置き換える必要があります。これらの例のいくつかは、これらの関数の範囲外で宣言されている `globals` オブジェクト ( `var globals = {activeViewHandler:0, firstSlideId:0};`) にも依存しています。

これらのコード例では、プロジェクトが [Office.js v1.1 以降のライブラリを参照](../../docs/develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md)している必要があります。


## プレゼンテーションのアクティブ ビューの検出と ActiveViewChanged イベントの処理を行う

`getFileView` 関数は [Document.getActiveViewAsync](../../reference/shared/document.getactiveviewasync.md) メソッドを呼び出して、プレゼンテーションの現在のビューが "編集" ビュー (**[標準]** や **[アウトライン表示]** などの、スライドを編集できるビュー) なのか "読み取り" ビュー (**[スライド ショー]** や **[閲覧表示]**) なのかを返します。


```js
function getFileView() {
    //Gets whether the current view is edit or read.
    Office.context.document.getActiveViewAsync(function (asyncResult) {
        if (asyncResult.status == "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
        else {
            app.showNotification(asyncResult.value);
        }
    });
}
```

`registerActiveViewChanged` 関数は [addHandlerAsync](../../reference/shared/document.addhandlerasync.md) メソッドを呼び出して、[Document.ActiveViewChanged](../../reference/shared/document.activeviewchanged.md) イベントのハンドラーを登録します。この関数を実行した後は、プレゼンテーションのビューを変更するときに、`app.showNotification` の通知でアクティブなビュー モード ("読み取り" または "編集") が表示されるようになります。




```js
function registerActiveViewChanged() {
    Globals.activeViewHandler = function (args) {
        app.showNotification(JSON.stringify(args));
    }

    Office.context.document.addHandlerAsync(Office.EventType.ActiveViewChanged, Globals.activeViewHandler, 
        function (asyncResult) {
            if (asyncResult.status == "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
            else {
            app.showNotification(asyncResult.status);
            }
        });
}
```


## プレゼンテーションの URL を取得する

`getFileUrl` 関数は [Document.getFileProperties](../../reference/shared/document.getfilepropertiesasync.md) メソッドを呼び出して、プレゼンテーション ファイルの URL を取得します。


```js
function getFileUrl() {
    //Get the URL of the current file.
    Office.context.document.getFilePropertiesAsync(function (asyncResult) {
        var fileUrl = asyncResult.value.url;
        if (fileUrl == "") {
            app.showNotification("The file hasn't been saved yet. Save the file and try again");
        }
        else {
            app.showNotification(fileUrl);
        }
    });
}
```


## プレゼンテーションの特定のスライドに移動する

`getSelectedRange` 関数は [Document.getSelectedDataAsync](../../reference/shared/document.getselecteddataasync.md) メソッドを呼び出して、`asyncResult.value` から返される JSON オブジェクトを取得します。そのオブジェクトには、選択範囲のスライド (または現在のスライドのみ) の ID、タイトル、インデックスが入った "slides" という名前の配列が含まれています。この関数はまた、選択範囲の最初のスライドの ID をグローバル変数に保存します。


```js
function getSelectedRange() {
    // Get the id, title, and index of the current slide (or selected slides) and store the first slide id */
    Globals.firstSlideId = 0;

    Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange, function (asyncResult) {
        if (asyncResult.status == "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
        else {
            Globals.firstSlideId = asyncResult.value.slides[0].id;
            app.showNotification(JSON.stringify(asyncResult.value));
        }
    });
}
```

`goToFirstSlide` 関数は [Document.goToByIdAsync](../../reference/shared/document.gotobyidasync.md) メソッドを呼び出して、上記の `getSelectedRange` 関数が格納した最初のスライドの ID に移動します。




```js
function goToFirstSlide() {
    Office.context.document.goToByIdAsync(Globals.firstSlideId, Office.GoToType.Slide, function (asyncResult) {
        if (asyncResult.status == "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
        else {
            app.showNotification("Navigation successful");
        }
    });
}
```


## プレゼンテーション内のスライド間を移動する

`goToSlideByIndex` 関数は **Document.goToByIdAsync** メソッドを呼び出して、プレゼンテーションの次のスライドに移動します。


```js
function goToSlideByIndex() {
    var goToFirst = Office.Index.First;
    var goToLast = Office.Index.Last;
    var goToPrevious = Office.Index.Previous;
    var goToNext = Office.Index.Next;

    Office.context.document.goToByIdAsync(goToNext, Office.GoToType.Index, function (asyncResult) {
        if (asyncResult.status == "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
        else {
            app.showNotification("Navigation successful");
        }
    });
}
```




## その他のリソース

- [コンテンツ アドインおよび作業ウィンドウ アドインで、ドキュメントごとにアドインの状態と設定を保存する方法](../../docs/develop/persisting-add-in-state-and-settings.md#how-to-save-add-in-state-and-settings-per-document-for-content-and-task-pane-add-ins)

- [ドキュメントまたはスプレッドシート内のアクティブな選択範囲へのデータの読み取りおよび書き込み](../../docs/develop/read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)
    
- [PowerPoint または Word 用アドインからドキュメント全体を取得する](../../docs/develop/get-the-whole-document-from-an-add-in-for-powerpoint-or-word.md)
    
- [PowerPoint アドインでドキュメントのテーマを使用する](../powerpoint/use-document-themes-in-your-powerpoint-add-ins.md)
    
