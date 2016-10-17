
# <a name="tablebinding.rowcount-property"></a>TableBinding.rowCount プロパティ
テーブルの行数を整数値で取得します。

|||
|:-----|:-----|
|**ホスト:**|Access、Excel、Word|
|**[要件セット](../../docs/overview/specify-office-hosts-and-api-requirements.md)に指定できるもの**|TableBindings|
|**選択内容の最終変更**|1.1|

```
var rowCount = bindingObj.rowCount;
```


## <a name="return-value"></a>戻り値

指定された [TableBinding](../../reference/shared/binding.tablebinding.md) オブジェクト内の行数。


## <a name="remarks"></a>注釈

Excel 2013 および Excel Online で 1 行を選択して空のテーブルを挿入すると (**[挿入]** タブの **[テーブル]** を使用)、両方の Office ホスト アプリケーションで 1 行のヘッダーと空の 1 行を作成します。ただし、アプリのスクリプトで、この新規挿入されたテーブルのバインドを作成し (たとえば、[addFromSelectionAsync](../../reference/shared/bindings.addfromselectionasync.md) メソッドを使用)、**rowCount** プロパティの値を確認した場合、返される値はスプレッドシートを Excel 2013 で開いているか Excel Online で開いているかによって異なります。


- Excel 2013 の場合、 **rowCount** は 0 を返します (ヘッダーに続く空の行は計算されません)。
    
- Excel Online の場合、 **rowCount** は 1 を返します (ヘッダーに続く空の行は計算されます)。
    
スクリプトでこの違いを回避するには、 `rowCount == 1` かどうかを確認し、これが真の場合、行に含まれている文字列がすべて空であるかどうかを確認します。

Access 用コンテンツ アプリでは、パフォーマンス上の理由から  **rowCount** プロパティは常に -1 を返します。


## <a name="example"></a>例




```js
function showBindingRowCount() {
    Office.context.document.bindings.getByIdAsync("myBinding", function (asyncResult) {
        write("Rows: " + asyncResult.value.rowCount);
    });
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```




## <a name="support-details"></a>サポートの詳細


次の表で、大文字 Y は、このプロパティは、対応する Office ホスト アプリケーションでサポートされていることを示します。空のセルは、Office ホスト アプリケーションでこのプロパティをサポートしないことを示します。

Office ホスト アプリケーションとサーバーの要件の詳細については、「[Office アドインを実行するための要件](../../docs/overview/requirements-for-running-office-add-ins.md)」をご覧ください。


**サポートされるホスト (プラットフォーム別)**


||**Windows デスクトップ版 Office**|**Office Online (ブラウザー)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||Y||
|**Excel**|Y|Y|Y|
|**Word**|Y|Y|Y|

|||
|:-----|:-----|
|**要件セットに指定できるもの**|TableBindings|
|**最小限のアクセス許可レベル**|[ReadDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**アドインの種類**|コンテンツ、作業ウィンドウ|
|**ライブラリ**|Office.js|
|**名前空間**|Office|

## <a name="support-history"></a>サポート履歴



****


|**バージョン**|**変更内容**|
|:-----|:-----|
|1.1|Office for iPad における Excel と Word のサポートが追加されました。|
|1.1|Access 用のアドインのサポートが追加されました。|
|1.0|導入|
