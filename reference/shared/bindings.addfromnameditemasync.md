
# <a name="bindings.addfromnameditemasync-method"></a>Bindings.addFromNamedItemAsync メソッド
ドキュメント内の名前付きの項目にバインドを追加します。

|||
|:-----|:-----|
|**ホスト:**|Access、Excel、Word|
|**[要件セット](../../docs/overview/specify-office-hosts-and-api-requirements.md)に指定できるもの**|MatrixBindings, TableBindings, TextBindings|
|**最終変更**|1.1|

```
Office.context.document.bindings.addFromNamedItemAsync(itemName, bindingType [, options], callback);
```


## <a name="parameters"></a>パラメーター



|**名前**|**型**|**説明**|**サポートのメモ**|
|:-----|:-----|:-----|:-----|
| _itemName_|**string**|名前付きの項目の名前。必須。||
| _bindingType_|[BindingType](../../reference/shared/bindingtype-enumeration.md)|作成するバインディング オブジェクトの種類を指定します。必須。選択したオブジェクトを指定の種類の強制変換できない場合、**null** を返します。||
| _options_|**object**|次の[オプションのパラメーター](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)のいずれかを指定します。||
| _id_|**string**|新しいバインド オブジェクトの一意の識別名を指定します。_id_ パラメーターに引数が渡されない場合は、[Binding.id](../../reference/shared/binding.id.md) が自動生成されます。||
| _asyncContext_|**array**、**boolean**、**null**、**number**、**object**、**string**、または **undefined**|変更されずに **AsyncResult** オブジェクトで返される任意の型のユーザー定義項目。||
| _callback_|**object**|コールバックが戻るときに呼び出される関数で、唯一のパラメーターは **AsyncResult** 型です。||

## <a name="callback-value"></a>コールバック値

_callback_ パラメーターに渡した関数が実行されると、[AsyncResult](../../reference/shared/asyncresult.md) オブジェクトを受け取ります。このオブジェクトには、コールバック関数の唯一のパラメーターからアクセスできます。

**addFromNamedItemAsync** メソッドに渡されるコールバック関数では、**AsyncResult** オブジェクトのプロパティを使用して次の情報を返すことができます。



|**プロパティ**|**使用目的**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|指定された名前のアイテムを表す [Binding](../../reference/shared/binding.md) オブジェクトにアクセスします。|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|操作の成功または失敗を判断します。|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|操作が失敗した場合、エラーに関する情報を提供する [Error](../../reference/shared/error.md) オブジェクトにアクセスします。|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|ユーザー定義の **object** または値を _asyncContext_ パラメーターとして渡した場合、そのオブジェクトまたは値にアクセスします。|

## <a name="remarks"></a>解説

 **Excel の場合**、_itemName_ パラメーターは、名前付きの範囲またはテーブルを参照できます。

既定では、Excel のテーブルを追加すると、最初に追加したテーブルには "Table1"、次に追加したテーブルには "Table2" という名前が割り当てられます。Excel UI で意味のあるテーブル名を割り当てるには、リボンの **[テーブル ツール | デザイン]** タブの **[テーブル名]** プロパティを使用します。


 >**注**: Excel では、テーブルを名前付きアイテムとして指定する場合、`"Sheet1!Table1"` の形式で完全修飾名を指定して、テーブルの名前にワークシートの名前を含める必要があります。

 **Word の場合**、_itemName_ パラメーターは、**リッチテキスト** コンテンツ コントロールの **[タイトル]** プロパティを参照します。(**リッチテキスト** コンテンツ コントロール以外のコンテンツ コントロールにバインドすることはできません。)

既定では、コンテンツ コントロールには [ **タイトル**] 値が割り当てられません。Word UI で意味のあるテーブル名を割り当てるには、リボンの [ **開発者**] タブの [ **コントロール**] グループから [ **リッチ テキスト**] コンテンツ コントロールを挿入した後、[ **コントロール**] グループの [ **プロパティ**] コマンドを使用して [ **コンテンツ コントロールのプロパティ**] ダイアログ ボックスを表示します。次に、コンテンツ コントロールの [ **タイトル**] プロパティに、コードから参照する名前を設定します。


 >**メモ**  Word では、同じ **[タイトル]** プロパティ値 (名前) を持つ複数の **リッチテキスト** コンテンツ コントロールが存在する場合に、このメソッドを使用して (名前を _itemName_ パラメーターに指定して)、これらのコンテンツ コントロールを 1 つにまとめようとすると、エラーが発生します。


## <a name="example"></a>例

次の例では、Excel の `myRange` という名前の項目にマトリックス ("matrix") バインドを追加し、そのバインドの [id](../../reference/shared/binding.id.md) に `myMatrix` を割り当てます。


```js
function bindNamedItem() {
    Office.context.document.bindings.addFromNamedItemAsync("myRange", "matrix", {id:'myMatrix'}, function (result) {
        if (result.status == 'succeeded'){
            write('Added new binding with type: ' + result.value.type + ' and id: ' + result.value.id);
            }
        else
            write('Error: ' + result.error.message);
    });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

次の例では、Excel の `Table1` という名前の項目にテーブル ("table") バインドを追加し、そのバインドの **id** に `myTable` を割り当てます。




```js
function bindNamedItem() {
    Office.context.document.bindings.addFromNamedItemAsync("Table1", "table", {id:'myTable'}, function (result) {
        if (result.status == 'succeeded'){
            write('Added new binding with type: ' + result.value.type + ' and id: ' + result.value.id);
            }
        else
            write('Error: ' + result.error.message);
    });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

次の例では、 `"FirstName"` という名前のリッチ テキスト コンテンツ コントロールに Word のテキスト バインドを作成し、 **id**`"firstName"` を割り当て、その情報を表示します。




```js
function bindContentControl() {
    Office.context.document.bindings.addFromNamedItemAsync('FirstName', 
        Office.BindingType.Text, {id:'firstName'},
        function (result) {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                write('Control bound. Binding.id: '
                    + result.value.id + ' Binding.type: ' + result.value.type);
            } else {
                write('Error:', result.error.message);
            }
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

## <a name="see-also"></a>関連項目



#### <a name="other-resources"></a>その他の技術情報


[ドキュメントまたはスプレッドシート内の領域へのバインド](../../docs/develop/bind-to-regions-in-a-document-or-spreadsheet.md#add-a-binding-to-a-named-item)
