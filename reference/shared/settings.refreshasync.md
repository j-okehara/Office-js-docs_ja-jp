

# Settings.refreshAsync メソッド
ドキュメントに保持されている設定をすべて読み取って、メモリ内に保持されているこれらの設定のコンテンツまたは作業ウィンドウ アドインのコピーを更新します。

|||
|:-----|:-----|
|**ホスト:**|Access、Excel、PowerPoint、Word|
|**[要件セット](../../docs/overview/specify-office-hosts-and-api-requirements.md)に指定できるもの**|設定値|
|**最終変更バージョン**|1.1|

```js
Office.context.document.settings.refreshAsync(callback);
```


## パラメーター

_callback_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;型: **object**

&nbsp;&nbsp;&nbsp;&nbsp;コールバックが戻るときに呼び出される関数で、唯一のパラメーターは **AsyncResult** 型です。

    



## コールバック値

_callback_ パラメーターに渡した関数が実行されると、[AsyncResult](../../reference/shared/asyncresult.md) オブジェクトを受け取ります。このオブジェクトには、コールバック関数の唯一のパラメーターからアクセスできます。

**refreshAsync** メソッドに渡されるコールバック関数では、**AsyncResult** オブジェクトのプロパティを使用して次の情報を返すことができます。



|**プロパティ**|**使用目的**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|更新された値を持つ [Settings](../../reference/shared/settings.md) オブジェクトにアクセスします。|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|操作の成功または失敗を判断します。|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|操作が失敗した場合、エラーに関する情報を提供する [Error](../../reference/shared/error.md) オブジェクトにアクセスします。|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|ユーザー定義の  **object** または値を _asyncContext_ パラメーターとして渡した場合、そのオブジェクトまたは値にアクセスします。|

## 注釈

このメソッドは、Word および PowerPoint の共同編集のシナリオで、同じアドインの複数のインスタンスから同じドキュメントを操作する場合に有効です。各アドインは、ユーザーがドキュメントを開いたときにドキュメントから読み込まれる設定のコピー (メモリ内に保持される) を操作するため、ユーザー間で設定値が一致しないことがあります。こうした状況は、アドインのインスタンスから [Settings.saveAsync](../../reference/shared/settings.saveasync.md) メソッドを呼び出して、その特定のユーザーのすべての設定をドキュメントに保存すると発生する可能性があります。すべてのユーザーの設定値を更新するには、アドインの **settingsChanged** イベントのイベント ハンドラーから [refreshAsync](../../reference/shared/settings.settingschangedevent.md) メソッドを呼び出します。

**refreshAsync** メソッドは、Excel 用に作成されたアドインから呼び出せますが、Excel は共同編集をサポートしていないため、このメソッドを呼び出すことはありません。


## 例




```js
function refreshSettings() {
    Office.context.document.settings.refreshAsync(function (asyncResult) {
        write('Settings refreshed with status: ' + asyncResult.status);
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
|**PowerPoint**|Y|Y|Y|
|**Word**|Y|Y|Y|

|||
|:-----|:-----|
|**要件セットに指定できるもの**|設定値|
|**最小限のアクセス許可レベル**|[Restricted](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**アプリの種類**|コンテンツ、作業ウィンドウ|
|**ライブラリ**|Office.js|
|**名前空間**|Office|

## サポート履歴




|**変更内容**|**1.1**|
|:-----|:-----|
|1.1|PowerPoint Online のサポートが追加されました。|
|1.1|Office for iPad で Excel、PowerPoint、および Word のサポートが追加されました。|
|1.1|Access 用コンテンツ アドインにおけるカスタム設定のサポートが追加されました。|
|1.0|導入|
