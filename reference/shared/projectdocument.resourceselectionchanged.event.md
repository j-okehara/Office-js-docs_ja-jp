

# ProjectDocument.ResourceSelectionChanged イベント
アクティブ プロジェクト内でリソースの選択が変更されるときに発生します。

|||
|:-----|:-----|
|**ホスト:**|Project|
|**[要件セット](../../docs/overview/specify-office-hosts-and-api-requirements.md)に指定できるもの**|選択内容|
|**で追加**|1.0|

```js
Office.EventType.ResourceSelectionChanged
```


## 注釈

 **ResourceSelectionChanged** は、[EventType](../../reference/shared/eventtype-enumeration.md) 列挙定数で、イベント ハンドラーを追加または削除するために [ProjectDocument.addHandlerAsync](../../reference/shared/projectdocument.addhandlerasync.md) および [ProjectDocument.removeHandlerAsync](../../reference/shared/projectdocument.removehandlerasync.md) メソッドで使用できます。


## 例

次のコード例では、 **ResourceSelectionChanged** イベントのハンドラーを追加します。ドキュメント内のリソース選択が変更されると、このハンドラーは、選択されたリソースの GUID を取得します。

この例では、アプリに jQuery ライブラリへの参照が指定されており、ページ本文の content div で次のページ コントロールが定義されていることを想定しています。




```HTML
<span id="message"></span>
```




```js
(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {

            // After the DOM is loaded, add-in-specific code can run.
            Office.context.document.addHandlerAsync(
                Office.EventType.ResourceSelectionChanged,
                getResourceGuid);
        });
    };

    // Get the GUID of the selected resource and display it in the add-in.
    function getResourceGuid() {
        Office.context.document.getSelectedResourceAsync(
            function (result) {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    onError(result.error);
                }
                else {
                    $('#message').html(result.value);
                }
            }
        );
    }

    function onError(error) {
        $('#message').html(error.name + ' ' + error.code + ': ' + error.message);
    }
})();
```

Project アドインでの **ResourceSelectionChanged** イベント ハンドラーの使用方法を示す完全なコード サンプルについては、「[テキスト エディターを使用して Project 2013 用の作業ウィンドウ アドインを初めて作成する](../../docs/project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md)」を参照してください。


## サポートの詳細


次の表で、大文字 Y は、このイベントは、対応する Office ホスト アプリケーションでサポートされていることを示します。空のセルは、Office ホスト アプリケーションがこのイベントをサポートしないことを示します。

Office ホスト アプリケーションとサーバーの要件の詳細については、「[Office アドインを実行するための要件](../../docs/overview/requirements-for-running-office-add-ins.md)」をご覧ください。


||**Windows デスクトップ版 Office**|**Office Online (ブラウザー)**|
|:-----|:-----|:-----|
|**Project**|Y||

|||
|:-----|:-----|
|**要件セットに指定できるもの**|選択内容|
|**最小限のアクセス許可レベル**|[ReadDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**アプリの種類**|作業ウィンドウ|
|**ライブラリ**|Office.js|
|**名前空間**|Office|

## サポート履歴

|**変更内容**|**1.1**|
|:-----|:-----|
|1.0|導入|

## 関連項目



#### その他の技術情報


[テキスト エディターを使用して Project 2013 用の作業ウィンドウ アドインを初めて作成する](../../docs/project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md)
[EventType 列挙](../../reference/shared/eventtype-enumeration.md)
[ProjectDocument.addHandlerAsync メソッド](../../reference/shared/projectdocument.addhandlerasync.md)
[ProjectDocument.removeHandlerAsync メソッド](../../reference/shared/projectdocument.removehandlerasync.md)
[ProjectDocument オブジェクト](../../reference/shared/projectdocument.projectdocument.md)
