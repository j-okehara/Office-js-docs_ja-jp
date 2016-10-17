
# <a name="projectdocument.getprojectfieldasync-method"></a>ProjectDocument.getProjectFieldAsync メソッド
アクティブなプロジェクトの指定したフィールドの値を非同期に取得します。

|||
|:-----|:-----|
|**ホスト:**|Project|
|**[要件セット](../../docs/overview/specify-office-hosts-and-api-requirements.md)に指定できるもの**|選択内容|
|**追加されたバージョン**|1.0|

```
Office.context.document.getProjectFieldAsync(fieldId[, options][, callback]);
```


## <a name="parameters"></a>パラメーター



|**名前**|**型**|**説明**|
|:-----|:-----|:-----|:-----|
| _fieldId_|[ProjectProjectFields](../../reference/shared/projectprojectfields-enumeration.md)|ターゲット フィールドの ID。必須。|
| _options_|**object**|次の[オプションのパラメーター](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)のいずれかを指定します。|
| _asyncContext_|**array**、**boolean**、**null**、**number**、**object**、**string**、または **undefined**|変更されずに **AsyncResult** オブジェクトで返される任意の型のユーザー定義項目。|
| _callback_|**object**|コールバックが戻るときに呼び出される関数で、唯一のパラメーターは **AsyncResult** 型です。|

## <a name="callback-value"></a>コールバック値

_callback_ 関数が実行されるとき、その関数は [AsyncResult](../../reference/shared/asyncresult.md) オブジェクトを受け取ります。このオブジェクトは、コールバック関数のパラメーターからアクセスできます。

**getProjectFieldAsync** メソッドの場合、返される [AsyncResult](../../reference/shared/asyncresult.md) オブジェクトには次のプロパティが含まれています。


****


|**名前**|**説明**|
|:-----|:-----|
|[asyncContext](../../reference/shared/asyncresult.asynccontext.md)|省略可能な _asyncContext_ パラメーターに入れて渡されたデータ (このパラメーターが使用された場合)。|
|[error](../../reference/shared/asyncresult.error.md)|**status** プロパティが **failed** と等しい場合に、エラーに関する情報。|
|[status](../../reference/shared/asyncresult.status.md)|非同期呼び出しの **succeeded** または **failed** 状態。|
|[value](../../reference/shared/asyncresult.value.md)|**fieldValue** プロパティが含まれます。これは、指定したフィールドの値を表します。|

## <a name="example"></a>例

次のコード例では、アクティブなプロジェクトの、指定された 3 つのフィールドの値を取得した後、その値をアプリで表示します。

この例では、前の呼び出しが正常に戻った後、 **getProjectFieldAsync** を再帰的に呼び出します。また、呼び出しを追跡して、すべての呼び出しの送信時刻を判別します。

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

            // Get information for the active project.
            getProjectInformation();
        });
    };

    // Get the specified fields for the active project.
    function getProjectInformation() {
        var fields =
            [Office.ProjectProjectFields.Start, Office.ProjectProjectFields.Finish, Office.ProjectProjectFields.GUID];
        var fieldValues = ['Start: ', 'Finish: ', 'GUID: '];
        var index = 0; 
        getField();

        // Get each field, and then display the field values in the add-in.
        function getField() {
            if (index == fields.length) {
                var output = '';
                for (var i = 0; i < fieldValues.length; i++) {
                    output += fieldValues[i] + '<br />';
                }
                $('#message').html(output);
            }
            else {
                Office.context.document.getProjectFieldAsync(
                    fields[index],
                    function (result) {

                        // If the call is successful, get the field value and then get the next field.
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



****


|**バージョン**|**変更内容**|
|:-----|:-----|
|1.0|導入|

## <a name="see-also"></a>関連項目



#### <a name="other-resources"></a>その他の技術情報


[ProjectProjectFields 列挙型](../../reference/shared/projectprojectfields-enumeration.md)

[AsyncResult オブジェクト](../../reference/shared/asyncresult.md)

[ProjectDocument オブジェクト](../../reference/shared/projectdocument.projectdocument.md)
