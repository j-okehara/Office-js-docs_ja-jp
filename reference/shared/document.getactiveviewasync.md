
# Document.getActiveViewAsync メソッド
 プレゼンテーションの現在のビューの状態を返します (編集または読み取り)。

|||
|:-----|:-----|
|**ホスト:** Excel、PowerPoint、Word|**アドインの種類:** コンテンツ、作業ウィンドウ|
|**[要件セット](../../docs/overview/specify-office-hosts-and-api-requirements.md)に指定できるもの**|ActiveView|
|**ActiveView で追加**|1.1|

```
Office.context.document.getActiveViewAsync([,options], callback);
```


## パラメーター



|**名前**|**型**|**説明**|**サポートのメモ**|
|:-----|:-----|:-----|:-----|
| _オプション_|**object**|次の[オプションのパラメーター](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)のいずれかを指定します||
| _asyncContext_|**array**、 **boolean**、 **null**、 **number**、 **object** 、 **string**、または  **undefined**|変更されずに  **AsyncResult** オブジェクトで返される任意の型のユーザー定義項目。||
| _callback_|**object**|コールバックが戻るときに呼び出される関数で、唯一のパラメーターは  **AsyncResult** 型です。||

## コールバック値

_callback_ パラメーターに渡した関数が実行されると、[AsyncResult](../../reference/shared/asyncresult.md) オブジェクトを受け取ります。このオブジェクトには、コールバック関数の唯一のパラメーターからアクセスできます。

**getActiveViewAsync** メソッドに渡されるコールバック関数で、[AsyncResult.value](../../reference/shared/asyncresult.value.md) プロパティは、プレゼンテーションの現在のビューの状態を返します。戻り値は、`edit` または `read` のいずれかです。`edit` は、スライドを編集できるいずれかのビュー、たとえば **[標準]** または **[アウトライン表示]** などに該当します。`read` は、**[スライド ショー]** または **[読み取り表示]** のいずれかに該当します。


## 注釈

ビューが変更されたときにイベントをトリガーできます。


## 例

現在のプレゼンテーションのビューを取得するには、その値を返すコールバック関数を記述する必要があります。次の例は、その方法を示しています。


-  ビューの種類を返す **匿名のコールバック関数を** _getActiveViewAsync_ メソッドの **callback**パラメーターに渡します。
    
-  アドインのページに **値を表示します**。
    

```js
function getFileView() {
    // Get whether the current view is edit or read.
    Office.context.document.getActiveViewAsync(function (asyncResult) {
        if (asyncResult.status == "failed") {
            showMessage("Action failed with error: " + asyncResult.error.message);
        }
        else {
            showMessage(asyncResult.value);
        }
    });
}
```




## サポートの詳細


次の表で、大文字 Y は、このメソッドは、対応する Office ホスト アプリケーションでサポートされていることを示します。空のセルは、Office ホスト アプリケーションでこのメソッドをサポートしないことを示します。

Office ホスト アプリケーションとサーバーの要件の詳細については、「[Office アドインを実行するための要件](../../docs/overview/requirements-for-running-office-add-ins.md)」をご覧ください。


**サポートされるホスト (プラットフォーム別)**


||**Windows デスクトップ版 Office**|**Office Online (ブラウザー)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|||Y|
|**PowerPoint**|Y|Y|Y|
|**Word**|||Y|

|||
|:-----|:-----|
|**要件セットに指定できるもの**|ActiveView|
|**ActiveView で追加**|1.1|
|**最小限のアクセス許可レベル**|[Restricted](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**アプリの種類**|コンテンツ、作業ウィンドウ|
|**ライブラリ**|Office.js|
|**名前空間**|Office|

## サポート履歴





****


|**変更内容**|**1.1**|
|:-----|:-----|
|1.1|Office for iPad で Excel、PowerPoint、および Word のサポートが追加されました。|
|1.1|導入。|
