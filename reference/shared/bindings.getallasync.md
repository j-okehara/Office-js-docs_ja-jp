
# Bindings.getAllAsync メソッド
以前に作成されたバインドをすべて取得します。

|||
|:-----|:-----|
|**ホスト:**|Access、Excel、Word|
|**[要件セット](../../docs/overview/specify-office-hosts-and-api-requirements.md)に指定できるもの**|MatrixBindings, TableBindings, TextBindings|
|**最終変更**|1.1|

```
bindingsObj.getAllAsync([options], callback);
```


## パラメーター



|**名前**|**型**|**説明**|**サポートのメモ**|
|:-----|:-----|:-----|:-----|
| _オプション_|**object**|次の[オプションのパラメーター](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)のいずれかを指定します||
| _asyncContext_|**array**、 **boolean**、 **null**、 **number**、 **object** 、 **string**、または  **undefined**|変更されずに  **AsyncResult** オブジェクトで返される任意の型のユーザー定義項目。||
| _callback_|**object**|コールバックが戻るときに呼び出される関数で、唯一のパラメーターは  **AsyncResult** 型です。||

## コールバック値

_callback_ パラメーターに渡した関数が実行されると、[AsyncResult](../../reference/shared/asyncresult.md) オブジェクトを受け取ります。このオブジェクトには、コールバック関数の唯一のパラメーターからアクセスできます。

コールバック関数が  **getAllAsync** メソッドに渡された場合、 **AsyncResult** オブジェクトのプロパティを使用して、次の情報を返すことができます。



|**プロパティ**|**使用目的**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|参照される [Bindings](../../reference/shared/bindings.bindings.md) オブジェクトに対して作成される各バインディングを格納する配列にアクセスします。|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|操作の成功または失敗を判断します。|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|操作が失敗した場合、エラーに関する情報を提供する [Error](../../reference/shared/error.md) オブジェクトにアクセスします。|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|ユーザー定義の  **object** または値を _asyncContext_ パラメーターとして渡した場合、そのオブジェクトまたは値にアクセスします。|

## 例




```js
function displayAllBindingNames() {
    Office.context.document.bindings.getAllAsync(function (asyncResult) {
        var bindingString = '';
        for (var i in asyncResult.value) {
            bindingString += asyncResult.value[i].id + '\n';
        }
        write('Existing bindings: ' + bindingString);
    });
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```




## サポートの詳細


次の表で、大文字 Y は、このメソッドは、対応する Office ホスト アプリケーションでサポートされていることを示します。空のセルは、Office ホスト アプリケーションでこのメソッドをサポートしないことを示します。

Office ホスト アプリケーションとサーバーの要件の詳細については、「[Office アドインを実行するための要件](../../docs/overview/requirements-for-running-office-add-ins.md)」をご覧ください。


||**Windows デスクトップ版 Office**|**Office Online (ブラウザー)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||Y||
|**Excel**|Y|Y|Y|
|**Word**|Y||Y|

|||
|:-----|:-----|
|**要件セットに指定できるもの**|MatrixBindings, TableBindings, TextBindings|
|**最小限のアクセス許可レベル**|[Restricted](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**アプリの種類**|コンテンツ、作業ウィンドウ|
|**ライブラリ**|Office.js|
|**名前空間**|Office|

## サポート履歴




|**変更内容**|**1.1**|
|:-----|:-----|
|1.1|Office for iPad における Excel と Word のサポートが追加されました。|
|1.1|
            Access 用コンテンツ アプリにおけるテーブルのバインドのサポートが追加されました。 |
|1.0|導入|
