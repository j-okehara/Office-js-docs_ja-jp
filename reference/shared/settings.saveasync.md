
# <a name="settings.saveasync-method"></a>Settings.saveAsync メソッド
設定プロパティ バッグのメモリ内コピーをドキュメントに保持します。

|||
|:-----|:-----|
|**ホスト:**|Access、Excel、PowerPoint、Word|
|**[要件セット](../../docs/overview/specify-office-hosts-and-api-requirements.md)に指定できるもの**|設定値|
|**最終変更バージョン**|1.1|

```js
Office.context.document.settings.saveAsync(callback);
```


## <a name="parameters"></a>パラメーター



_callback_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;型: **object**

&nbsp;&nbsp;&nbsp;&nbsp;コールバックが戻るときに呼び出される関数で、唯一のパラメーターは **AsyncResult** 型です。省略可能。

    



## <a name="callback-value"></a>コールバック値

_callback_ パラメーターに渡した関数が実行されると、[AsyncResult](../../reference/shared/asyncresult.md) オブジェクトを受け取ります。このオブジェクトには、コールバック関数の唯一のパラメーターからアクセスできます。

**saveAsync** メソッドに渡されたコールバック関数で、**AsyncResult** オブジェクトのプロパティを使用し、次の情報を返すことができます。



|**プロパティ**|**使用目的**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|取得するオブジェクトまたはデータがないため、常に **undefined** を返します。|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|操作の成功または失敗を判断します。|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|操作が失敗した場合、エラーに関する情報を提供する [Error](../../reference/shared/error.md) オブジェクトにアクセスします。|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|ユーザー定義の **object** または値を _asyncContext_ パラメーターとして渡した場合、そのオブジェクトまたは値にアクセスします。|

## <a name="remarks"></a>注釈

アドイン によって過去に保存された設定は、アプリの初期化時に読み込まれます。したがって、セッションの実行中に [set](../../reference/shared/settings.set.md) および [get](../../reference/shared/settings.get.md) メソッドを使用し、設定プロパティ バッグのメモリ内コピーで作業できます。これらの設定を アドイン の次回使用時にも使用できるように保存するときは、**saveAsync** メソッドを使用します。


 >**メモ**:  **saveAsync** メソッドでは、メモリ内設定プロパティをドキュメント ファイルに保持します。ただし、ドキュメント ファイル自体への変更は、ユーザー (または **AutoRecover** 設定) がファイル システムにドキュメントを保存する場合にのみ保存されます。

同じ アドイン の他のインスタンスが設定を変更する可能性があり、その変更がすべてのインスタンスで利用できるようにする必要がある場合、[refreshAsync](../../reference/shared/settings.refreshasync.md) メソッドは共同編集のシナリオ (Word でのみサポートされる) でのみ有効です。


## <a name="example"></a>例




```js
function persistSettings() {
    Office.context.document.settings.saveAsync(function (asyncResult) {
        write('Settings saved with status: ' + asyncResult.status);
    });
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```




## <a name="support-details"></a>サポートの詳細


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
|**最小限のアクセス許可レベル**|[制限あり](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**アドインの種類**|コンテンツ、作業ウィンドウ|
|**ライブラリ**|Office.js|
|**名前空間**|Office|

## <a name="support-history"></a>サポート履歴




|**バージョン**|**変更内容**|
|:-----|:-----|
|1.1|PowerPoint Online のサポートが追加されました。|
|1.1|Office for iPad で Excel、PowerPoint、および Word のサポートが追加されました。|
|1.1|Access 用コンテンツ アドインにおけるカスタム設定のサポートが追加されました。|
|1.0|導入|
