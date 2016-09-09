
# ProjectDocument.getSelectedTaskAsync メソッド
タスク ビューで選択されているタスクの GUID を非同期に取得します。

|||
|:-----|:-----|
|**ホスト:**|Project|
|**[要件セット](../../docs/overview/specify-office-hosts-and-api-requirements.md)に指定できるもの**|選択内容|
|**で追加**|1.0|

```
Office.context.document.getSelectedTaskAsync([options,] [callback]);
```


## パラメーター



|**名前**|**型**|**説明**|**サポートのメモ**|
|:-----|:-----|:-----|:-----|
| _オプション_|**object**|次の[オプションのパラメーター](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)のいずれかを指定します。||
| _asyncContext_|**array**、 **boolean**、 **null**、 **number**、 **object** 、 **string**、または  **undefined**|変更されずに  **AsyncResult** オブジェクトで返される任意の型のユーザー定義項目。||
| _callback_|**object**|コールバックが戻るときに呼び出される関数で、唯一のパラメーターは  **AsyncResult** 型です。||

## コールバック値

_callback_ 関数が実行されるとき、その関数は [AsyncResult](../../reference/shared/asyncresult.md) オブジェクトを受け取ります。このオブジェクトは、コールバック関数のパラメーターからアクセスできます。

**getSelectedTaskAsync** メソッドの場合、返される [AsyncResult](../../reference/shared/asyncresult.md) オブジェクトには次のプロパティが含まれています。


****


|**名前**|**説明**|
|:-----|:-----|
|[asyncContext](../../reference/shared/asyncresult.asynccontext.md)|省略可能な  _asyncContext_ に入れて渡されたデータ (このパラメーターが使用された場合)。|
|[エラー](../../reference/shared/asyncresult.error.md)|**status** プロパティが **failed** と等しい場合に、エラーに関する情報。|
|[status](../../reference/shared/asyncresult.status.md)|非同期呼び出しの  **succeeded** または **failed** 状態。|
|[value](../../reference/shared/asyncresult.value.md)|**string** としての、選択されたタスクの GUID。|

## 注釈

Project アドインでは、タスクの ID 番号を使用するよりも、タスクの GUID を使用した方が便利です (たとえば、ガント チャートの最初のタスクの ID は **1** です)。タスクの GUID は Project タスク情報 (Project Server と認識モードで同期する SharePoint プロジェクトのタスクなど) にアクセスするために使用できます。また、タスク GUID をローカル変数に保存し、それを [getTaskAsync](../../reference/shared/projectdocument.gettaskasync.md) および [getTaskFieldAsync](../../reference/shared/projectdocument.gettaskfieldasync.md) メソッドで使用することもできます。

アクティブ ビューがタスク ビュー (ガント チャート ビューまたはタスク分配状況ビューなど) でない場合、またはタスク ビューでタスクが選択されていない場合は、**getSelectedTaskAsync** は 5001 エラー (内部エラー) を返します。[ViewSelectionChanged](../../reference/shared/projectdocument.addhandlerasync.md) イベントと [getSelectedViewAsync](../../reference/shared/projectdocument.viewselectionchanged.event.md) メソッドを使用して、アクティブ ビューの種類に基づいてボタンをアクティブにする例については、「[addHandlerAsync メソッド](../../reference/shared/projectdocument.getselectedviewasync.md)」をご覧ください。


## 例

次のコード例では、タスク ビューで現在選択されているタスクの GUID を、**getSelectedTaskAsync** メソッドを使用して取得します。その後、[getTaskAsync](../../reference/shared/projectdocument.gettaskasync.md) を呼び出すことによってタスクのプロパティを取得します。

この例では、アドインに jQuery ライブラリへの参照が指定されており、ページ本文の 内容 div で次のページ コントロールが定義されていることを想定しています。




```HTML
<input id="get-info" type="button" value="Get info" /><br />
<span id="message"></span>
```




```js
(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {

            // After the DOM is loaded, add-in-specific code can run.
            $('#get-info').click(getTaskInfo);
        });
    };

    // // Get the GUID of the task, and then get local task properties.
    function getTaskInfo() {
        getTaskGuid().then(
            function (data) {
                getTaskProperties(data);
            }
        );
    }

    // Get the GUID of the selected task.
    function getTaskGuid() {
        var defer = $.Deferred();
        Office.context.document.getSelectedTaskAsync(
            function (result) {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    onError(result.error);
                }
                else {
                    defer.resolve(result.value);
                }
            }
        );
        return defer.promise();
    }

    // Get local properties for the selected task, and then display it in the add-in.
    function getTaskProperties(taskGuid) {
        Office.context.document.getTaskAsync(
            taskGuid,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    onError(result.error);
                }
                else {
                    var taskInfo = result.value;
                    var output = String.format(
                        'Name: {0}<br/>GUID: {1}<br/>SharePoint task ID: {2}<br/>Resource names: {3}',
                        taskInfo.taskName, taskGuid, taskInfo.wssTaskId, taskInfo.resourceNames);
                    $('#message').html(output);
                }
            }
        );
    }

    function onError(error) {
        $('#message').html(error.name + ' ' + error.code + ': ' + error.message);
    }
})();

    


```


## サポートの詳細


次の表で、大文字 Y は、このメソッドは、対応する Office ホスト アプリケーションでサポートされていることを示します。空のセルは、Office ホスト アプリケーションでこのメソッドをサポートしないことを示します。

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



****


|**変更内容**|**1.1**|
|:-----|:-----|
|1.0|導入|

## 関連項目



#### その他の技術情報


[getTaskAsync メソッド](../../reference/shared/projectdocument.gettaskasync.md)

[AsyncResult オブジェクト](../../reference/shared/asyncresult.md)

[ProjectDocument オブジェクト](../../reference/shared/projectdocument.projectdocument.md)
