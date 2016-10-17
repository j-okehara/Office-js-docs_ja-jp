
# <a name="bindings.addfrompromptasync-method"></a>Bindings.addFromPromptAsync メソッド
 ユーザーがバインド先の選択範囲を指定できるようにするための UI を表示します。

|||
|:-----|:-----|
|**ホスト:**|Access、Excel|
|**[要件セット](../../docs/overview/specify-office-hosts-and-api-requirements.md)に指定できるもの**|セットには指定できない|
|**最終変更**|1.1|

```
_bindingsObj.addFromPromptAsync(bindingType [, options], callback);
```


## <a name="parameters"></a>パラメーター



|**名前**|**型**|**説明**|**サポートのメモ**|
|:-----|:-----|:-----|:-----|
| _bindingType_|[BindingType](../../reference/shared/bindingtype-enumeration.md)|作成するバインディング オブジェクトの種類を指定します。必須。選択したオブジェクトを指定の種類の強制変換できない場合、**null** を返します。||
| _options_|**object**|次の[オプションのパラメーター](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)のいずれかを指定します。||
| _id_|**string**|新しいバインド オブジェクトの一意の識別名を指定します。_id_ パラメーターに引数が渡されない場合は、[Binding.id](../../reference/shared/binding.id.md) が自動生成されます。||
| _promptText_|**string**|ユーザーに選択するものを示すプロンプトを UI に表示する文字列を指定します。200 文字以下の制限があります。_promptText_ 引数を渡さないと、"選択範囲を指定してください" と表示されます。||
| _sampleData_|[TableData](../../reference/shared/tabledata.md)|アドインによってバインドできるフィールド (列) の種類の例としてプロンプト UI に表示するサンプル データのテーブルを指定します。**TableData** オブジェクトで提供されるヘッダーは、フィールド選択 UI で使用されるラベルを指定します。省略可能。**注:**このパラメーターは、Access 用アドインでのみ使用されます。Excel 用アドインのメソッドを呼び出すときに指定した場合は無視されます。||
| _asyncContext_|**array**、**boolean**、**null**、**number**、**object**、**string**、または **undefined**|変更されずに **AsyncResult** オブジェクトで返される任意の型のユーザー定義項目。||
| _callback_|**object**|コールバックが戻るときに呼び出される関数で、唯一のパラメーターは **AsyncResult** 型です。||

## <a name="callback-value"></a>コールバック値

_callback_ パラメーターに渡した関数が実行されると、[AsyncResult](../../reference/shared/asyncresult.md) オブジェクトを受け取ります。このオブジェクトには、コールバック関数の唯一のパラメーターからアクセスできます。

**addFromPromptAsync** メソッドに渡されたコールバック関数で、**AsyncResult** オブジェクトのプロパティを使用し、次の情報を返すことができます。



|**プロパティ**|**使用目的**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|ユーザーによって指定された選択範囲を表す [Binding](../../reference/shared/binding.md) オブジェクトにアクセスします。|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|操作の成功または失敗を判断します。|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|操作が失敗した場合、エラーに関する情報を提供する [Error](../../reference/shared/error.md) オブジェクトにアクセスします。|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|ユーザー定義の **object** または値を _asyncContext_ パラメーターとして渡した場合、そのオブジェクトまたは値にアクセスします。|

## <a name="remarks"></a>解説

指定された型のバインド オブジェクトを [Bindings](../../reference/shared/bindings.bindings.md) コレクションに追加します。このバインド オブジェクトは、提供される _id_ で識別できるようになります。指定された選択範囲をバインドできない場合、メソッドは失敗します。


## <a name="example"></a>例




```js
function addBindingFromPrompt() {

    Office.context.document.bindings.addFromPromptAsync(Office.BindingType.Text, { id: 'MyBinding', promptText: 'Select text to bind to.' }, function (asyncResult) {
        write('Added new binding with type: ' + asyncResult.value.type + ' and id: ' + asyncResult.value.id);
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


|**Windows デスクトップ版 Office**|**Office Online (ブラウザー)**|**Office for iPad**|
|:-----|:-----|:-----|
|**Access**||Y||
|**Excel**|Y|Y|Y|

|||
|:-----|:-----|
|**要件セットに指定できるもの**|セットには指定できない|
|**最小限のアクセス許可レベル**|[ReadDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**アドインの種類**|コンテンツ、作業ウィンドウ|
|**ライブラリ**|Office.js|
|**名前空間**|Office|

## <a name="support-history"></a>サポート履歴




|**バージョン**|**変更内容**|
|:-----|:-----|
|1.1|Office for iPad の Excel のサポートが追加されました。|
|1.1|([**挿入**] > [**テーブル**] > [**テーブル**] または [**ホーム**] > [**スタイル**] > [**テーブルとして書式設定**] コマンドを使って) データがテーブルとしてスプレッドシートに追加されなかった場合でも、Excel 用アプリで (_bindingType_ を **Office.BindingType.Table** として渡すことによって) 表形式データが含まれる一定範囲のセルのテーブル バインドを作成できます。|
|1.1|
            Access 用コンテンツ アプリにおけるテーブルのバインドのサポートが追加されました。 |
|1.1|Excel 用アプリにおいてテーブルのバインドとしてマトリックス データにバインドするサポートが追加されました。|
|1.0|導入|
