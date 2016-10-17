
# <a name="document.getselecteddataasync-method"></a>Document.getSelectedDataAsync メソッド
ドキュメントの現在の選択範囲に含まれるデータを読み取ります。

|||
|:-----|:-----|
|**ホスト:**|Access、Excel、PowerPoint、Project、Word|
|**要件セットに指定できるもの**|選択内容|
|**選択内容の最終変更**|1.1|

```js
Office.context.document.getSelectedDataAsync(coercionType [, options], callback); 
```


## <a name="parameters"></a>パラメーター



|**名前**|**型**|**説明**|**サポートのメモ**|
|:-----|:-----|:-----|:-----|
| _coercionType_|[CoercionType](../../reference/shared/coerciontype-enumeration.md)<br/><table><tr><td></td><td><b>ホスト サポート</b></td></tr><tr><td><b>Office.CoercionType.Text</b> (文字列)</td><td>Excel、Excel Online、PowerPoint、PowerPoint Online、Word、および Word Online のみ</td></tr><tr><td><b>Office.CoercionType.Matrix</b> (配列の配列)</td><td>Excel、Word、および Word Online のみ</td></tr><tr><td><b>Office.CoercionType.Table</b> ([TableData](../../reference/shared/tabledata.md) オブジェクト)</td><td>Access、Excel、Word、および Word Online のみ</td></tr><tr><td><b>Office.CoercionType.Html</b></td><td>Word のみ。</td></tr><tr><td><b>Office.CoercionType.Ooxml</b> (Office Open XML)</td><td>Word および Word Online のみ</td></tr><tr><td><b>Office.CoercionType.SlideRange</b></td><td>PowerPoint および PowerPoint Online のみ</td></tr></table>|返されるデータ構造の種類です。必須。||
| _options_|**object**<br/><table><tr><td><i>valueFormat</i></td><td><b>[ValueFormat](../../reference/shared/valueformat-enumeration.md)</b></td><td>結果を数値で返すか、または書式設定ありまたはなしの日付値で返すかを指定します。</td><td></td></tr><tr><td><i>filterType</i></td><td>[FilterType](../../reference/shared/filtertype-enumeration.md)</td><td>データの取得時にフィルターを適用するかどうかを指定します。省略可能。</td><td>このパラメーターは Word 文書では無視されます。</td></tr><tr><td><i>asyncContext</i></td><td><b>array</b>、<b>boolean</b>、<b>null</b>、<b>number</b>、<b>object</b>、<b>string</b>、または <b>undefined</b></td><td>変更されずに <b>AsyncResult</b> オブジェクトで返される任意の型のユーザー定義項目。</td><td></td></tr></table>|次の[オプションのパラメーター](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)のいずれかを指定します||
| _callback_|**object**|コールバックが戻るときに呼び出される関数で、唯一のパラメーターは **AsyncResult** 型です。||

## <a name="callback-value"></a>コールバック値

_callback_ パラメーターに渡した関数が実行されると、[AsyncResult](../../reference/shared/asyncresult.md) オブジェクトを受け取ります。このオブジェクトには、コールバック関数の唯一のパラメーターからアクセスできます。

**getSelectedDataAsync** メソッドに渡されるコールバック関数では、**AsyncResult** オブジェクトのプロパティを使用して、次の情報を返すことができます。



|**プロパティ**|**使用目的**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|現在選択されている値にアクセスします。この値は、 _coercionType_ パラメーターで指定したデータ構造またはデータ形式で返されます (データの強制型変換の詳細については、「 **コメント** 」を参照してください)。|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|操作の成功または失敗を判断します。|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|操作が失敗した場合、エラーに関する情報を提供する [Error](../../reference/shared/error.md) オブジェクトにアクセスします。|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|ユーザー定義の **object** または値を _asyncContext_ パラメーターとして渡した場合、そのオブジェクトまたは値にアクセスします。|

## <a name="remarks"></a>注釈

作業ウィンドウ アドインまたはコンテンツ アドインで、 **getSelectedDataAsync** メソッドを使用して、ドキュメント、スプレッドシート、プレゼンテーション、またはプロジェクトのユーザーの選択内容からデータを読み取るスクリプトを記述します。たとえば、ユーザーが Word ドキュメントのコンテンツを選択したら、 **getSelectedDataAsync** メソッドを使用して、その選択を読み取り、それをクエリしたりその他の何らかの操作としたりして、Web サービスに送信することができます。

選択の読み取り後、[Document](../../reference/shared/document.setselecteddataasync.md) オブジェクトの [setSelectedDataAsync](../../reference/shared/document.addhandlerasync.md) および **addHandlerAsync** メソッドを使用して、[選択に書き戻すか、イベント ハンドラーを追加して](../../docs/develop/read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)、ユーザーが選択を変更しているかどうかを検出することもできます。

**getSelectedDataAsync** メソッドは、選択がアクティブである場合に限り、選択から読み取ることができます。Word および Excel のアドインで、ユーザーが選択した内容の読み取りと書き込みのために、永続的な関連付けを作成する必要がある場合、代わりに [Bindings.addFromSelectionAsync](../../reference/shared/bindings.addfromselectionasync.md) メソッドを使用して、[その選択にバインドします](../../docs/develop/bind-to-regions-in-a-document-or-spreadsheet.md)。

**getSelectedDataAsync** メソッドの _coercionType_ パラメーターを使用して、データ構造や選択したデータの読み方を指定します。



|**指定した _coercionType_**|**返されるデータ**|**Office ホスト アプリケーションのサポート**|
|:-----|:-----|:-----|
|**Office.CoercionType.Text** または `"text"`|文字列。|Word、Excel、PowerPoint、および Project。<br/><br/> **メモ** Excel で、セルのサブセットが選択されている場合でも、セルのコンテンツ全体が返されます。|
|**Office.CoercionType.Matrix** または `"matrix"`|配列の配列。たとえば、2 行 X 2 列の選択の場合、` [['a','b'], ['c','d']]` となります。|Word および Excel。|
|**Office.CoercionType.Table** または `"table"`|ヘッダーのあるテーブルを読み取るための [TableData](../../reference/shared/tabledata.md) オブジェクト。|Word および Excel。|
|**Office.CoercionType.Html** または `"html"`|HTML 形式。|Word のみ。|
|**Office.CoercionType.Ooxml** または `"ooxml"`|Open Office XML (OpenXML) 形式。|Word のみ。<br/><br/> **ヒント**アドインのコードを開発する場合、_getSelectedDataAsync_ メソッドの `"ooxml"`**coercionType** を使用して、Word 文書で選択したコンテンツが OpenXML タグでどのように定義されているか確認できます。次に、[Document.setSelectedDataAsync](../../reference/shared/document.setselecteddataasync.md) メソッドのデータ パラメーターでそれらのタグを使用して、その形式や構造でコンテンツをドキュメントに書き込みます。たとえば、OpenXML として[イメージをドキュメントに挿入](http://blogs.msdn.com/b/officeapps/archive/2012/10/26/inserting-images-with-apps-for-office.aspx)できます。|
|**Office.CoercionType.SlideRange** または "slideRange"|選ばれているスライドの ID、タイトル、インデックスが含まれている「スライド」という名前の配列が含まれる JSON オブジェクト。  **メモ:** 複数のスライドを選択するには、ユーザーは [ **標準** ]、[ **アウトライン表示**]、または [ **スライド一覧表示**] 表示でプレゼンテーションを編集する必要があります。また、このメソッドは [ **マスター表示** ] ではサポートされていません。たとえば、2 つのスライドが選ばれている場合は  `{"slides":[{"id":257,"title":"Slide 2","index":2},{"id":256,"title":"Slide 1","index":1}]}`。|PowerPoint のみ。|
選択内容のデータ構造が指定した  _coercionType_ に一致しない場合、 **getSelectedDataAsync** メソッドはデータをその型または構造にデータを強制型変換しようとします。選択内容を、指定した **Office.CoercionType** に強制型変換できない場合、 **AsyncResult.status** プロパティは `"failed"` を返します。


## <a name="example"></a>例

現在の選択内容の値を取得するには、選択を読み取るコールバック関数を記述する必要があります。次の例は、その方法を示しています。


-  現在の選択の値を読み取る **匿名のコールバック関数を** _getSelectedDataAsync_ メソッドの **callback**パラメーターに渡します。
    
-  書式なしでフィルター処理されていないテキストとして**選択を読み取ります**。
    
-  アドインのページに **値を表示します**。
    

```js
function getText() {
    Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, 
        { valueFormat: "unformatted", filterType: "all" },
        function (asyncResult) {
            var error = asyncResult.error;
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                write(error.name + ": " + error.message);
            } 
            else {
                // Get selected data.
                var dataValue = asyncResult.value; 
                write('Selected data is ' + dataValue);
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


**サポートされるホスト (プラットフォーム別)**


||**Windows デスクトップ版 Office**|**Office Online (ブラウザー)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||Y||
|**Excel**|Y|Y|Y|
|**PowerPoint**|Y|Y|Y|
|**Project**|Y|||
|**Word**|Y|Y|Y|

|||
|:-----|:-----|
|**要件セットに指定できるもの**|選択内容|
|**最小限のアクセス許可レベル**|[ReadDocument (Office Open XML を取得するために必要な ReadAllDocument)](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**アドインの種類**|コンテンツ、作業ウィンドウ|
|**ライブラリ**|Office.js|
|**名前空間**|Office|

## <a name="support-history"></a>サポート履歴



|**バージョン**|**変更内容**|
|:-----|:-----|
|1.1|PowerPoint Online のサポートが追加されました。|
|1.1| Word Online で、 **coercionType** パラメーターとして **Office.CoercionType.Matrix** と _Office.CoercionType.Table_ のサポートが追加されました。|
|1.1|Office for iPad の Excel、PowerPoint、および Word で、Windows デスクトップの Excel、PowerPoint、および Word と同じレベルのサポートが追加されました。|
|1.1| Word Online で、 **coercionType** パラメーターとして _Office.CoercionType.Text_ のサポートが追加されました。|
|1.1|PowerPoint 用コンテンツ アドインで、**Office.CoercionType.SlideRange** を _getSelectedDataAsync_ メソッドの **coercionType** パラメーターとして渡すことによって、選択範囲のスライドの ID、タイトル、およびインデックスを取得できます。この値を使って現在選ばれているスライドに移動する方法の例については、[Document.goToByIdAsync](../../reference/shared/document.gotobyidasync.md) メソッドのトピックをご覧ください。|
|1.0|導入|
