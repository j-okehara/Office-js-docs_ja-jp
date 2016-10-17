
# <a name="projectdocument.settaskfieldasync-method-(javascript-api-for-office-v1.1)"></a>ProjectDocument.setTaskFieldAsync メソッド (JavaScript API for Office v1.1)
指定したタスクの指定したフィールドの値を非同期に設定します。 **重要:**この API は、Wndows デスクトップ上の Project 2016 でのみ動作します。

|||
|:-----|:-----|
|**ホスト:**|Project|
|**[要件セット](../../docs/overview/specify-office-hosts-and-api-requirements.md)に指定できるもの**|選択内容|
|**追加されたバージョン**|1.1|

```js
Office.context.document.setTaskFieldAsync(taskId, fieldId, fieldValue[, options][, callback]);
```


## <a name="parameters"></a>パラメーター


_taskId_<br/>&nbsp;&nbsp;&nbsp;&nbsp;タスクの GUID。必須。<br/><br/>_fieldId_<br/>&nbsp;&nbsp;&nbsp;&nbsp;ターゲット フィールドの ID ([ProjectTaskFields](../../reference/shared/projecttaskfields-enumeration.md) 定数またはそれに対応する整数値)。必須。<br/><br/>_fieldValue_<br/>&nbsp;&nbsp;&nbsp;&nbsp;ターゲット フィールドの値 (**string**、**number**、**boolean**、**object**)。必須。<br/><br/>_options_<br/>&nbsp;&nbsp;&nbsp;&nbsp;次の[オプションのパラメーター](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods):<br/><br/>

&nbsp;&nbsp;&nbsp;&nbsp;_asyncContext_<br/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;型: **array、boolean、null、number、object、string、undefined**<br/></br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;変更されずに [AsyncResult](../../reference/shared/asyncresult.md) オブジェクトで返される任意の型のユーザー定義項目。省略可能。</br></br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;たとえば、_asyncContext_ 引数を渡すことができます。形式として `{asyncContext: 'Some text'}` または `{asyncContext: <object>}` を使用します。<br/><br/>_callback_<br/>&nbsp;&nbsp;&nbsp;&nbsp;型: **function**<br/><br/>&nbsp;&nbsp;&nbsp;&nbsp;メソッド コールが戻るときに呼び出される関数で、唯一のパラメーターは [AsyncResult](../../reference/shared/asyncresult.md) 型です。省略可能。
    

## <a name="callback-value"></a>コールバック値

_callback_ 関数が実行されるとき、その関数は [AsyncResult](../../reference/shared/asyncresult.md) オブジェクトを受け取ります。このオブジェクトは、コールバック関数のパラメーターからアクセスできます。

**setTaskFieldAsync** メソッドの場合、返される [AsyncResult](../../reference/shared/asyncresult.md) オブジェクトには次のプロパティが含まれています。



|**名前**|**説明**|
|:-----|:-----|
|[asyncContext](../../reference/shared/asyncresult.asynccontext.md)|省略可能な _asyncContext_ パラメーターに入れて渡されたデータ (このパラメーターが使用された場合)。|
|[error](../../reference/shared/asyncresult.error.md)|**status** プロパティが **failed** と等しい場合に、エラーに関する情報。|
|[status](../../reference/shared/asyncresult.status.md)|非同期呼び出しの **succeeded** または **failed** 状態。|
|[value](../../reference/shared/asyncresult.value.md)|このメソッドは値を返しません。|

## <a name="remarks"></a>注釈

最初に [getSelectedTaskAsync](../../reference/shared/projectdocument.getselectedtaskasync.md) メソッドまたは [getTaskByIndexAsync](../../reference/shared/projectdocument.settaskfieldasync.md) メソッドを使ってタスク GUID を取得し、次に GUID を _taskId_ 引数として **setTaskFieldAsync** に渡します。各非同期の呼び出しで更新できるのは、1 つのタスクの 1 つのフィールドだけです。


## <a name="example"></a>例

次のコード例では、タスク ビューで現在選択されているタスクの GUID を、[getSelectedResourceAsync](../../reference/shared/projectdocument.getselectedtaskasync.md) メソッドを使用して取得します。その後、 **setTaskFieldAsync** を再帰的に呼び出すと 2 つのタスク フィールド値が設定されます。

例で使われている [getSelectedTaskAsync](../../reference/shared/projectdocument.getselectedtaskasync.md) メソッドでは、タスク ビュー (たとえば、タスク配分状況) がアクティブなビューで、タスクが選ばれている必要があります。アクティブなビューの種類に基づいてボタンをアクティブ化する例については、[addHandlerAsync](../../reference/shared/projectdocument.addhandlerasync.md) メソッドをご覧ください。

この例では、アドインに jQuery ライブラリへの参照が指定されており、ページ本文の 内容 div で次のページ コントロールが定義されていることを想定しています。




```HTML
<input id="set-info" type="button" value="Set info" /><br />
<span id="message"></span>
```




```js
(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            
            // After the DOM is loaded, add-in-specific code can run.
            app.initialize();
            $('#set-info').click(setTaskInfo);
        });
    };

    // Get the GUID of the task, and then get the task fields.
    function setTaskInfo() {
        getTaskGuid().then(
            function (data) {
                setTaskFields(data);
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

    // Set the specified fields for the selected task.
    function setTaskFields(taskGuid) {
        var targetFields = [Office.ProjectTaskFields.Active, Office.ProjectTaskFields.Notes];
        var fieldValues = [true, 'Notes for the task.'];

        // Set the field value. If the call is successful, set the next field.
        for (var i = 0; i < targetFields.length; i++) {
            Office.context.document.setTaskFieldAsync(
                taskGuid,
                targetFields[i],
                fieldValues[i],
                function (result) {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        i++;
                    }
                    else {
                        onError(result.error);
                    }
                }
            );
        }
        $('#message').html('Field values set');
    }

    function onError(error) {
        app.showNotification(error.name + ' ' + error.code + ': ' + error.message);
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
|**最小限のアクセス許可レベル**|[WriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**アドインの種類**|作業ウィンドウ|
|**ライブラリ**|Office.js|
|**名前空間**|Office|

## <a name="support-history"></a>サポート履歴



|**バージョン**|**変更内容**|
|:-----|:-----|
|1.1|導入|

## <a name="see-also"></a>関連項目



#### <a name="other-resources"></a>その他の技術情報


[getSelectedTaskAsync メソッド](../../reference/shared/projectdocument.getselectedresourceasync.md)[getTaskByIndexAsync](../../reference/shared/projectdocument.settaskfieldasync.md)[AsyncResult オブジェクト](../../reference/shared/asyncresult.md)[ProjectTaskFields 列挙型](../../reference/shared/projecttaskfields-enumeration.md)[ProjectDocument オブジェクト](../../reference/shared/projectdocument.projectdocument.md)
