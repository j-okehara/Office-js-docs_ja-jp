
# <a name="bindings.addfromselectionasync-method"></a>Bindings.addFromSelectionAsync メソッド
ドキュメント内の現在の選択範囲にバインドを追加します。

|||
|:-----|:-----|
|**ホスト:**|Access、Excel、Word|
|**[要件セット](../../docs/overview/specify-office-hosts-and-api-requirements.md)に指定できるもの**|MatrixBindings, TableBindings, TextBindings|
|**最終変更**|1.1|

```
bindingsObj.addFromSelectionAsync(bindingType [, options], callback);
```


## <a name="parameters"></a>パラメーター



|**名前**|**型**|**説明**|**サポートのメモ**|
|:-----|:-----|:-----|:-----|
| _bindingType_|[BindingType](../../reference/shared/bindingtype-enumeration.md)|作成するバインディング オブジェクトの種類を指定します。必須。選択したオブジェクトを指定の種類の強制変換できない場合、**null** を返します。||
| _options_|**object**|次の[オプションのパラメーター](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)のいずれかを指定します。||
| _id_|**string**|新しいバインド オブジェクトの一意の識別名を指定します。_id_ パラメーターに引数が渡されない場合は、[Binding.id](../../reference/shared/binding.id.md) が自動生成されます。||
| _asyncContext_|**array**、**boolean**、**null**、**number**、**object**、**string**、または **undefined**|変更されずに **AsyncResult** オブジェクトで返される任意の型のユーザー定義項目。||
| _callback_|**object**|コールバックが戻るときに呼び出される関数で、唯一のパラメーターは **AsyncResult** 型です。||

## <a name="callback-value"></a>コールバック値

_callback_ パラメーターに渡した関数が実行されると、[AsyncResult](../../reference/shared/asyncresult.md) オブジェクトを受け取ります。このオブジェクトには、コールバック関数の唯一のパラメーターからアクセスできます。

コールバック関数が **addFromSelectionAsync** メソッドに渡された場合、**AsyncResult** オブジェクトのプロパティを使用して、次の情報を返すことができます。



|**プロパティ**|**使用目的**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|ユーザーによって指定された選択範囲を表す [Binding](../../reference/shared/binding.md) オブジェクトにアクセスします。|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|操作の成功または失敗を判断します。|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|操作が失敗した場合、エラーに関する情報を提供する [Error](../../reference/shared/error.md) オブジェクトにアクセスします。|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|ユーザー定義の **object** または値を _asyncContext_ パラメーターとして渡した場合、そのオブジェクトまたは値にアクセスします。|

## <a name="remarks"></a>解説

指定された種類のバインド オブジェクトを  **Bindings** コレクションに追加します。このバインド オブジェクトは、指定された _id_ で識別されます。


 >**メモ**  Excel では、既存のバインドの **Binding.id** に渡す **addFromSelectionAsync** メソッドを呼び出した場合、そのバインドの [Binding.type](../../reference/shared/binding.type.md) が使用され、その種類は _bindingType_ パラメーターに別の値を指定しても変更できません。既存の _id_ を使用して _bindingType_ を変更する必要がある場合は、最初に [Bindings.releaseByIdAsync](../../reference/shared/bindings.releasebyidasync.md) メソッドを呼び出してバインドを解放し、次に **addFromSelectionAsync** メソッドを呼び出してバインドを新しい種類で再確立します。


## <a name="example"></a>例

'MyBinding' の [Binding.id](../../reference/shared/binding.textbinding.md) により、現在の選択範囲に**TextBinding** を追加します。


```js
function addBindingFromSelection() {
    Office.context.document.bindings.addFromSelectionAsync(Office.BindingType.Text, { id: 'MyBinding' }, 
        function (asyncResult) {
        write('Added new binding with type: ' + asyncResult.value.type + ' and id: ' + asyncResult.value.id);
        }
    );
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
|**Word**|Y|Y|Y|

|||
|:-----|:-----|
|**要件セットに指定できるもの**|MatrixBindings, TableBindings, TextBindings|
|**最小限のアクセス許可レベル**|[ReadDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**アドインの種類**|コンテンツ、作業ウィンドウ|
|**ライブラリ**|Office.js|
|**名前空間**|Office|

## <a name="support-history"></a>サポート履歴



****


|**バージョン**|**変更内容**|
|:-----|:-----|
|1.1|Office for iPad における Excel と Word のサポートが追加されました。|
|1.1|([ _挿入_]  >  [ **テーブル**]  >  [ **テーブル**] または [ **ホーム**]  >  [ **スタイル**]  >  [ **テーブルとして書式設定**] コマンドを使って) データがテーブルとしてスプレッドシートに追加されなかった場合でも、Excel 用アプリで ( **bindingType** を **Office.BindingType.Table** として渡すことによって) 表形式データが含まれる一定範囲のセルのテーブル バインドを作成できます。|
|1.1|
            Access 用コンテンツ アプリにおけるテーブルのバインドのサポートが追加されました。 |
|1.0|導入|
