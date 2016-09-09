
# ProjectDocument.getSelectedDataAsync メソッド
ガント チャート ビュー内で現在選択されている 1 つ以上のセルに含まれるデータのテキスト値を非同期に取得します。

|||
|:-----|:-----|
|**ホスト:**|Project|
|**[要件セット](../../docs/overview/specify-office-hosts-and-api-requirements.md)に指定できるもの**|選択内容|
|**で追加**|1.0|

```
Office.context.document.getSelectedDataAsync(coercionType[, options][, callback]);
```


## パラメーター



|**名前**|**型**|**説明**|**サポートのメモ**|
|:-----|:-----|:-----|:-----|
| _coercionType_|[CoercionType](../../reference/shared/coerciontype-enumeration.md)|返されるデータ構造の種類です。必須です。<br/>Project 2013 は **Office.CoercionType.Text** または `"text"` のみをサポートします。||
| _オプション_|**object**|次の[オプションのパラメーター](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)のいずれかを指定します。||
| _valueFormat_|[ValueFormat](../../reference/shared/valueformat-enumeration.md)|数値または日付値に対して使用する書式設定。<br/>Project 2013 ではこのパラメーターは無視され、内部的に `unformatted` に設定されます。||
| _filterType_|[FilterType](../../reference/shared/filtertype-enumeration.md)|表示されているデータのみを含めるか、すべてのデータを含めるかを指定します。 <br/>Project 2013 ではこのパラメーターは無視され、内部的に `all` に設定されます。||
| _asyncContext_|**array**、 **boolean**、 **null**、 **number**、 **object** 、 **string**、または  **undefined**|変更されずに  **AsyncResult** オブジェクトで返される任意の型のユーザー定義項目。||
| _callback_|**object**|コールバックが戻るときに呼び出される関数で、唯一のパラメーターは  **AsyncResult** 型です。||

## コールバック値

_callback_ 関数が実行されるとき、その関数は [AsyncResult](../../reference/shared/asyncresult.md) オブジェクトを受け取ります。このオブジェクトは、コールバック関数のパラメーターからアクセスできます。

**getSelectedDataAsync** メソッドの場合、返される [AsyncResult](../../reference/shared/asyncresult.md) オブジェクトには次のプロパティが含まれています。


****


|**名前**|**説明**|
|:-----|:-----|
|[asyncContext](../../reference/shared/asyncresult.asynccontext.md)|省略可能な  _asyncContext_ に入れて渡されたデータ (このパラメーターが使用された場合)。|
|[エラー](../../reference/shared/asyncresult.error.md)|**status** プロパティが **failed** と等しい場合に、エラーに関する情報。|
|[status](../../reference/shared/asyncresult.status.md)|非同期呼び出しの  **succeeded** または **failed** 状態。|
|[value](../../reference/shared/asyncresult.value.md)|選択されたセルのテキスト値。|

## 注釈

**ProjectDocument.getSelectedDataAsync** メソッドは、[Document.getSelectedDataAsync](../../reference/shared/document.getselecteddataasync.md) メソッドをオーバーライドし、ガント チャート ビュー内の 1 つ以上セルで選択されているデータのテキスト値を返します。**ProjectDocument.getSelectedDataAsync** でサポートされるのは、[CoercionType](../../reference/shared/coerciontype-enumeration.md) としてのテキスト形式のみです。`matrix`、`table`、またはその他の形式はサポートされません。


## 例

次のコード例では、選択されたセルの値を取得します。また、オプションの  _asyncContext_ パラメーターを使用して、コールバック関数にテキストを渡します。

この例では、アプリに jQuery ライブラリへの参照が指定されており、ページ本文の 内容 div で次のページ コントロールが定義されていることを想定しています。




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
            $('#get-info').click(getSelectedText);
        });
    };

    // Get the text from the selected cells in the document, and display it in the add-in.
    function getSelectedText() {
        Office.context.document.getSelectedDataAsync(
            Office.CoercionType.Text,
            {asyncContext: 'Some related info'},
            function (result) {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    onError(result.error);
                }
                else {
                    var output = String.format(
                        'Selected text: {0}<br/>Passed info: {1}',
                        result.value, result.asyncContext);
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


[AsyncResult オブジェクト](../../reference/shared/asyncresult.md)

[Office.CoercionType](../../reference/shared/coerciontype-enumeration.md)

[ProjectDocument オブジェクト](../../reference/shared/projectdocument.projectdocument.md)
