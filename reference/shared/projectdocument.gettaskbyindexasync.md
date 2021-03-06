﻿

# <a name="projectdocument.gettaskbyindexasync-method"></a>ProjectDocument.getTaskByIndexAsync メソッド
タスクのコレクション内に指定のインデックスがあるタスクの GUID を非同期に取得します。

**重要:**この API は、Wndows デスクトップ上の Project 2016 でのみ動作します。

|||
|:-----|:-----|
|**ホスト:**|Project|
|**[要件セット](../../docs/overview/specify-office-hosts-and-api-requirements.md)に指定できるもの**|選択内容|
|**追加されたバージョン**|1.1|

```js
Office.context.document.getTaskByIndexAsync(taskIndex[, options][, callback]);
```


## <a name="parameters"></a>パラメーター

_taskIndex_<br/>&nbsp;&nbsp;&nbsp;&nbsp;型: **number**

&nbsp;&nbsp;&nbsp;&nbsp;プロジェクトのタスクのコレクションにあるタスクのインデックス。必須。

    
_options_<br/>&nbsp;&nbsp;&nbsp;&nbsp;次の[オプションのパラメーター](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods):


&nbsp;&nbsp;&nbsp;&nbsp;_asyncContext_<br/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;型: **array、boolean、null、number、object、string、undefined**<br/></br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;変更されずに [AsyncResult](../../reference/shared/asyncresult.md) オブジェクトで返される任意の型のユーザー定義項目。省略可能。</br></br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;たとえば、_asyncContext_ 引数を渡すことができます。形式として `{asyncContext: 'Some text'}` または `{asyncContext: <object>}` を使用します。

_callback_<br/>&nbsp;&nbsp;&nbsp;&nbsp;型: **function**

&nbsp;&nbsp;&nbsp;&nbsp;メソッド コールが戻るときに呼び出される関数で、唯一のパラメーターは [AsyncResult](../../reference/shared/asyncresult.md) 型です。省略可能。


## <a name="callback-value"></a>コールバック値

_callback_ 関数が実行されるとき、その関数は [AsyncResult](../../reference/shared/asyncresult.md) オブジェクトを受け取ります。このオブジェクトは、コールバック関数のパラメーターからアクセスできます。

**getTaskByIndexAsync** メソッドの場合、返される [AsyncResult](../../reference/shared/asyncresult.md) オブジェクトには次のプロパティが含まれています。


|**名前**|**説明**|
|:-----|:-----|
|[asyncContext](../../reference/shared/asyncresult.asynccontext.md)|省略可能な _asyncContext_ パラメーターに入れて渡されたデータ (このパラメーターが使用された場合)。|
|[error](../../reference/shared/asyncresult.error.md)|**status** プロパティが **failed** と等しい場合に、エラーに関する情報。|
|[status](../../reference/shared/asyncresult.status.md)|非同期呼び出しの **succeeded** または **failed** 状態。|
|[value](../../reference/shared/asyncresult.value.md)|**string** としてのタスクの GUID。|

## <a name="remarks"></a>注釈

プロジェクトのタスクのコレクションにある最大インデックスを取得するには、[getMaxTaskIndexAsync](../../reference/shared/projectdocument.getmaxtaskindexasync.md) メソッドを使います。0 インデックス タスクは、プロジェクトのサマリー タスクを表します。


## <a name="example"></a>例

次のコード例は、[getMaxTaskIndexAsync](../../reference/shared/projectdocument.getmaxtaskindexasync.md) を呼び出してプロジェクトのタスク コレクションの最大インデックスを取得し、 **getTaskByIndexAsync** を呼び出して各タスクの GUID を取得します。

この例では、アドインに jQuery ライブラリへの参照が指定されており、ページ本文の 内容 div で次のページ コントロールが定義されていることを想定しています。




```HTML
<input id="get-info" type="button" value="Get info" /><br />
<span id="message"></span>
```




```js
(function () {
    "use strict";
    var taskGuids = [];

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {

            // After the DOM is loaded, add-in-specific code can run.
            app.initialize();
            $('#get-info').click(getTaskInfo);
        });
    };

    // Get the maximum task index, and then get the task GUIDs.
    function getTaskInfo() {
        getMaxTaskIndex().then(
            function (data) {
                getTaskGuids(data);
            }
        );
    }

    // Get the maximum index of the tasks for the current project.
    function getMaxTaskIndex() {
        var defer = $.Deferred();
        Office.context.document.getMaxTaskIndexAsync(
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

    // Get each task GUID, and then display the GUIDs in the add-in.
    function getTaskGuids(maxTaskIndex) {
        var defer = $.Deferred();
        for (var i = 0; i <= maxTaskIndex; i++) {
            getTaskGuid(i);
        }
        return defer.promise();
        function getTaskGuid(index) {
            Office.context.document.getTaskByIndexAsync(index,
                function (result) {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        taskGuids.push(result.value);
                        if (index == maxTaskIndex) {
                            defer.resolve();
                            $('#message').html(taskGuids.toString());
                        }
                    }
                    else {
                        onError(result.error);
                    }
                }
            );
        }
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
|**最小限のアクセス許可レベル**|[ReadDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**アドインの種類**|作業ウィンドウ|
|**ライブラリ**|Office.js|
|**名前空間**|Office|

## <a name="support-history"></a>サポート履歴

|**バージョン**|**変更内容**|
|:-----|:-----|
|1.1|導入|

## <a name="see-also"></a>関連項目



#### <a name="other-resources"></a>その他の技術情報


[getMaxTaskIndexAsync](../../reference/shared/projectdocument.getmaxtaskindexasync.md)[AsyncResult オブジェクト](../../reference/shared/asyncresult.md)[ProjectDocument オブジェクト](../../reference/shared/projectdocument.projectdocument.md)
