

# <a name="projectdocument.gettaskfieldasync-method"></a>ProjectDocument.getTaskFieldAsync メソッド
指定したタスクの指定したフィールドの値を非同期に取得します。

|||
|:-----|:-----|
|**ホスト:**|Project|
|**[要件セット](../../docs/overview/specify-office-hosts-and-api-requirements.md)に指定できるもの**|選択内容|
|**追加されたバージョン**|1.0|

```js
Office.context.document.getTaskFieldAsync(taskId, fieldId[, options][, callback]);
```


## <a name="parameters"></a>パラメーター



|**名前**|**型**|**説明**|**サポートのメモ**|
|:-----|:-----|:-----|:-----|
| _taskId_|**string**|タスクの GUID。必須。||
| _fieldId_|[ProjectTaskFields](../../reference/shared/projecttaskfields-enumeration.md)|ターゲット フィールドの ID。必須。||
| _options_|**object**|次の[オプションのパラメーター](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)のいずれかを指定します。||
| _asyncContext_|**array**、**boolean**、**null**、**number**、**object**、**string**、または **undefined**|変更されずに **AsyncResult** オブジェクトで返される任意の型のユーザー定義項目。||
| _callback_|**object**|コールバックが戻るときに呼び出される関数で、唯一のパラメーターは **AsyncResult** 型です。||

## <a name="callback-value"></a>コールバック値

_callback_ 関数が実行されるとき、その関数は [AsyncResult](../../reference/shared/asyncresult.md) オブジェクトを受け取ります。このオブジェクトは、コールバック関数のパラメーターからアクセスできます。

**getTaskFieldAsync** メソッドの場合、返される [AsyncResult](../../reference/shared/asyncresult.md) オブジェクトには次のプロパティが含まれています。



|**名前**|**説明**|
|:-----|:-----|
|[asyncContext](../../reference/shared/asyncresult.asynccontext.md)|省略可能な _asyncContext_ パラメーターに入れて渡されたデータ (このパラメーターが使用された場合)。|
|[error](../../reference/shared/asyncresult.error.md)|**status** プロパティが **failed** と等しい場合に、エラーに関する情報。|
|[status](../../reference/shared/asyncresult.status.md)|非同期呼び出しの **succeeded** または **failed** 状態。|
|[value](../../reference/shared/asyncresult.value.md)|**fieldValue** プロパティが含まれます。これは、指定したフィールドの値を表します。|

## <a name="remarks"></a>注釈

最初に [getSelectedTaskAsync](../../reference/shared/projectdocument.getselectedtaskasync.md) メソッドを呼び出してタスク GUID を取得し、_taskId_ 引数として **getTaskFieldAsync** に渡します。アクティブ ビューがタスク ビュー (ガント チャート ビューやタスク配分状況ビューなど) ではない場合、あるいはタスク ビューでタスクが選択されていない場合、[getSelectedTaskAsync](../../reference/shared/projectdocument.getselectedtaskasync.md) は 5001 エラー (内部エラー) を返します。[ViewSelectionChanged](../../reference/shared/projectdocument.viewselectionchanged.event.md) イベントや [getSelectedViewAsync](../../reference/shared/projectdocument.getselectedviewasync.md) メソッドを利用し、アクティブなビューの種類に基づいてボタンをアクティブ化する例については、「[addHandlerAsync メソッド](../../reference/shared/projectdocument.addhandlerasync.md)」を参照してください。


## <a name="example"></a>例

次のコード例では、タスク ビューで現在選択されているタスクの GUID を、[getSelectedResourceAsync](../../reference/shared/projectdocument.getselectedtaskasync.md) メソッドを使用して取得します。その後、 **getTaskFieldAsync** を再帰的に呼び出すことにより、2つのタスク フィールド値を取得します。

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

    // Get the GUID of the task, and then get the task fields.
    function getTaskInfo() {
        getTaskGuid().then(
            function (data) {
                getTaskFields(data);
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

    // Get the specified fields for the selected task.
    function getTaskFields(taskGuid) {
        var output = '';
        var targetFields = [Office.ProjectTaskFields.Priority, Office.ProjectTaskFields.PercentComplete];
        var fieldValues = ['Priority: ', '% Complete: '];
        var index = 0;
        getField();

        // Get each field, and then display the field values in the add-in.
        function getField() {
            if (index == targetFields.length) {
                for (var i = 0; i < fieldValues.length; i++) {
                    output += fieldValues[i] + '<br />';
                }
                $('#message').html(output);
            }

            // Get the field value. If the call is successful, then get the next field.
            else {
                Office.context.document.getTaskFieldAsync(
                    taskGuid,
                    targetFields[index],
                    function (result) {
                        if (result.status === Office.AsyncResultStatus.Succeeded) {
                            fieldValues[index] += result.value.fieldValue;
                            getField(index++);
                        }
                        else {
                            onError(result.error);
                        }
                    }
                );
            }
        }
    }

    function onError(error) {
        $('#message').html(error.name + ' ' + error.code + ': ' + error.message);
    }
})();

```


## <a name="support-details"></a>サポートの詳細


次の表で、大文字 Y は、このメソッドは、対応する Office ホスト アプリケーションでサポートされていることを示します。空のセルは、Office ホスト アプリケーションでこのメソッドをサポートしないことを示します。

Office ホスト アプリケーションとサーバーの要件の詳細については、「[Office アドインを実行するための要件](../../docs/overview/requirements-for-running-office-add-ins.md)」をご覧ください。


||**Windows デスクトップ版 Office**|**Office Online (ブラウザー)**|
|:-----|:-----|:-----|
|**Project**|Y||

|||
|:-----|:-----|
|**要件セットに指定できるもの**||
|**最小限のアクセス許可レベル**|[ReadDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**アドインの種類**|作業ウィンドウ|
|**ライブラリ**|Office.js|
|**名前空間**|Office|

## <a name="support-history"></a>サポート履歴



|**バージョン**|**変更内容**|
|:-----|:-----|
|1.0|導入|

## <a name="see-also"></a>関連項目



#### <a name="other-resources"></a>その他の技術情報


[getSelectedTaskAsync メソッド](../../reference/shared/projectdocument.getselectedresourceasync.md)
[AsyncResult オブジェクト](../../reference/shared/asyncresult.md)
[ProjectTaskFields 列挙](../../reference/shared/projecttaskfields-enumeration.md)
[ProjectDocument オブジェクト](../../reference/shared/projectdocument.projectdocument.md)
