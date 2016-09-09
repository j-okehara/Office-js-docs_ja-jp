

# ProjectDocument.getSelectedViewAsync メソッド
ドキュメント内のアクティブ ビューのビューの種類と名前を非同期に取得します。

|||
|:-----|:-----|
|**ホスト:**|Project|
|**[要件セット](../../docs/overview/specify-office-hosts-and-api-requirements.md)に指定できるもの**|選択内容|
|**で追加**|1.0|

```js
Office.context.document.getSelectedViewAsync([options,] [callback]);
```


## パラメーター



|**名前**|**型**|**説明**|**サポートのメモ**|
|:-----|:-----|:-----|:-----|
| _オプション_|**object**|次の[オプションのパラメーター](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)のいずれかを指定します。||
| _asyncContext_|**array**、 **boolean**、 **null**、 **number**、 **object** 、 **string**、または  **undefined**|変更されずに  **AsyncResult** オブジェクトで返される任意の型のユーザー定義項目。||
| _callback_|**object**|コールバックが戻るときに呼び出される関数で、唯一のパラメーターは  **AsyncResult** 型です。||

## コールバック値

_callback_ 関数が実行されるとき、その関数は [AsyncResult](../../reference/shared/asyncresult.md) オブジェクトを受け取ります。このオブジェクトは、コールバック関数のパラメーターからアクセスできます。

**getSelectedViewAsync** メソッドの場合、返される [AsyncResult](../../reference/shared/asyncresult.md) オブジェクトには次のプロパティが含まれています。


****


|**名前**|**説明**|
|:-----|:-----|
|[asyncContext](../../reference/shared/asyncresult.asynccontext.md)|省略可能な  _asyncContext_ に入れて渡されたデータ (このパラメーターが使用された場合)。|
|[エラー](../../reference/shared/asyncresult.error.md)|**status** プロパティが **failed** と等しい場合に、エラーに関する情報。|
|[status](../../reference/shared/asyncresult.status.md)|非同期呼び出しの  **succeeded** または **failed** 状態。|
|[value](../../reference/shared/asyncresult.value.md)|次のプロパティが含まれます。<br/><br/><div>* **viewName** - ビューの名前 ([ProjectViewTypes](../../reference/shared/projectviewtypes-enumeration.md) 定数)。<br/>* **viewType** - ビューの種類 ([ProjectViewTypes](../../reference/shared/projectviewtypes-enumeration.md) 定数の整数値)。</div>|

## 例

次のコード例では、ドキュメント内のアクティブ ビューの名前と種類を取得する  [getSelectedViewAsync](../../reference/shared/projectdocument.viewselectionchanged.event.md) を呼び出す**ViewSelectionChanged** イベント ハンドラーを追加します。

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
                Office.EventType.ViewSelectionChanged,
                getActiveView);
            getActiveView();
        });
    };

    // Get the active view's name and type.
    function getActiveView() {
        Office.context.document.getSelectedViewAsync(
            function (result) {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    onError(result.error);
                }
                else {
                    var output = String.format(
                        'View name: {0}<br/>View type: {1}',
                        result.value.viewName, viewType);
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


[ProjectViewTypes 列挙](../../reference/shared/projectviewtypes-enumeration.md)
[AsyncResult オブジェクト](../../reference/shared/asyncresult.md)
[ViewSelectionChanged イベント](../../reference/shared/projectdocument.viewselectionchanged.event.md)
[ProjectDocument オブジェクト](../../reference/shared/projectdocument.projectdocument.md)
