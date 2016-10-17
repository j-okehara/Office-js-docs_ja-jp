
# <a name="projectdocument.getresourcebyindexasync-method-(javascript-api-for-office-v1.1)"></a>ProjectDocument.getResourceByIndexAsync メソッド (JavaScript API for Office v1.1)
リソースのコレクション内に指定のインデックスがあるリソースの GUID を非同期に取得します。

 **重要:**この API は、Wndows デスクトップ上の Project 2016 でのみ動作します。

|||
|:-----|:-----|
|**ホスト:**|Project|
|**[要件セット](../../docs/overview/specify-office-hosts-and-api-requirements.md)に指定できるもの**|選択内容|
|**追加されたバージョン**|1.1|

```js
Office.context.document.getResourceByIndexAsync(resourceIndex[, options][, callback]);
```


## <a name="parameters"></a>パラメーター

_resourceIndex_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;型: **number**<br/><br/>
&nbsp;&nbsp;&nbsp;&nbsp;プロジェクトのリソースのコレクションにあるリソースのインデックス。必須。
    
_options_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;次の[オプションのパラメーター](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods):<br/><br/>
&nbsp;&nbsp;&nbsp;&nbsp;_asyncContext_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;型: **array、boolean、null、number、object、string、undefined**<br/><br/>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;変更されずに [AsyncResult](../../reference/shared/asyncresult.md) オブジェクトで返される任意の型のユーザー定義項目。省略可能。<br/><br/>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;たとえば、_asyncContext_ 引数を渡すことができます。形式として `{asyncContext: 'Some text'}` または `{asyncContext: <object>}` を使用します。

_callback_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;型: **function**<br/><br/>
&nbsp;&nbsp;&nbsp;&nbsp;メソッド コールが戻るときに呼び出される関数で、唯一のパラメーターは [AsyncResult](../../reference/shared/asyncresult.md) 型です。省略可能。
    

## <a name="callback-value"></a>コールバック値

_callback_ 関数が実行されるとき、その関数は [AsyncResult](../../reference/shared/asyncresult.md) オブジェクトを受け取ります。このオブジェクトは、コールバック関数のパラメーターからアクセスできます。

**getResourceByIndexAsync** メソッドの場合、返される [AsyncResult](../../reference/shared/asyncresult.md) オブジェクトには次のプロパティが含まれています。



|**名前**|**説明**|
|:-----|:-----|
|[asyncContext](../../reference/shared/asyncresult.asynccontext.md)|省略可能な _asyncContext_ パラメーターに入れて渡されたデータ (このパラメーターが使用された場合)。|
|[error](../../reference/shared/asyncresult.error.md)|**status** プロパティが **failed** と等しい場合に、エラーに関する情報。|
|[status](../../reference/shared/asyncresult.status.md)|非同期呼び出しの **succeeded** または **failed** 状態。|
|[value](../../reference/shared/asyncresult.value.md)|**string** としてのリソースの GUID。|

## <a name="remarks"></a>注釈

プロジェクトのリソースのコレクションにある最大インデックスを取得するには、[getMaxResourceIndexAsync](../../reference/shared/projectdocument.getmaxresourceindexasync.md) メソッドを使います。リソース コレクションには、0 インデックスの位置にあるリソースは含まれません。


## <a name="example"></a>例

次のコード例は、[getMaxResourceIndexAsync](../../reference/shared/projectdocument.getmaxresourceindexasync.md) を呼び出してプロジェクトのリソース コレクションにある最大インデックスを取得し、 **getResourceByIndexAsync** を呼び出して各リソースの GUID を呼び出します。

この例では、アドインに jQuery ライブラリへの参照が指定されており、ページ本文の 内容 div で次のページ コントロールが定義されていることを想定しています。




```HTML
<input id="get-info" type="button" value="Get info" /><br />
<span id="message"></span>
```




```js
(function () {
    "use strict";
    var resourceGuids = [];

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {

            // After the DOM is loaded, add-in-specific code can run.
            app.initialize();
            $('#get-info').click(getResourceInfo);
        });
    };

    // Get the maximum resource index, and then get the resource GUIDs.
    function getResourceInfo() {
        getMaxResourceIndex().then(
            function (data) {
                getResourceGuids(data);
            }
        );
    }

    // Get the maximum index of the resources for the current project.
    function getMaxResourceIndex() {
        var defer = $.Deferred();
        Office.context.document.getMaxResourceIndexAsync(
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

    // Get each resource GUID, and then display the GUIDs in the add-in.
    // There is no 0 index for resources, so start with index 1.
    function getResourceGuids(maxResourceIndex) {
        var defer = $.Deferred();
        for (var i = 1; i <= maxResourceIndex; i++) {
            getResourceGuid(i);
        }
        return defer.promise();
        function getResourceGuid(index) {
            Office.context.document.getResourceByIndexAsync(index,
                function (result) {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        resourceGuids.push(result.value);
                        if (index == maxResourceIndex) {
                            defer.resolve();
                            $('#message').html(resourceGuids.toString());
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


[getMaxResourceIndexAsync](../../reference/shared/projectdocument.getmaxresourceindexasync.md)

[AsyncResult オブジェクト](../../reference/shared/asyncresult.md)

[ProjectDocument オブジェクト](../../reference/shared/projectdocument.projectdocument.md)
