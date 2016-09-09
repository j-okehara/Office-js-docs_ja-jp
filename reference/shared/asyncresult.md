
# AsyncResult オブジェクト
要求が失敗した場合の状態やエラー情報など、非同期要求の結果をカプセル化するオブジェクト。

|||
|:-----|:-----|
|**ホスト:**|Access、Excel、Outlook、PowerPoint、Project、Word|
|**最終変更バージョン**|1.1|

```
AsyncResult
```


## メンバー


**プロパティ**


|**名前**|**説明**|
|:-----|:-----|
|**[asyncContext](../../reference/shared/asyncresult.asynccontext.md)**|呼び出されたメソッドの省略可能な  _asyncContext_ パラメーターに渡されたユーザー定義アイテムを、渡されたときと同じ状態で取得します。|
|**[エラー](../../reference/shared/asyncresult.error.md)**|エラーの説明を提供する  **Error** オブジェクトを取得します (エラーが発生した場合)。|
|**[status](../../reference/shared/asyncresult.status.md)**|非同期操作の状態を取得します。|
|**[value](../../reference/shared/asyncresult.value.md)**|この非同期操作のペイロードまたはコンテンツを取得します (ある場合)。|

## 注釈

"Async" メソッドの _callback_ パラメーターに渡す関数が実行されると、[AsyncResult](../../reference/shared/asyncresult.md) オブジェクトを受け取ります。このオブジェクトには、コールバック関数の唯一のパラメーターからアクセスできます。

次に、コンテンツ アドインおよび作業ウィンドウ アドインに適用できる例を示します。次の例は、 [Document](../../reference/shared/document.getselecteddataasync.md) オブジェクトの**getSelectedDataAsync** メソッドの呼び出しを示しています。




```js
Office.context.document.getSelectedDataAsync("text", {valueFormat:"unformatted", filterType:"all"}, 
   function (result) {
      if (result.status === "success")      
         var dataValue = result.value; // Get selected data.
         write('Selected data is ' + dataValue);
      else {            
         var err = result.error; 
         write(err.name + ": " + err.message);
      }
   });
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}

```

_callback_ 引数として渡される匿名関数 ( `function (result){...}`) には、_result_ というパラメーターがあります。この関数が実行された場合、このパラメーターを使用して **AsyncResult** オブジェクトにアクセスできます。**getSelectedDataAsync** メソッドの呼び出しが完了して、コールバック関数が実行されると、次のコード行によって **AsyncResult** オブジェクトの **value** プロパティにアクセスして、ドキュメント内で選択されているデータが返されます。

 `var dataValue = result.value;`

この関数内の他のコード行は、コールバック関数の  _result_ パラメーターを使用して、 **AsyncResult** オブジェクトの **status** および **error** プロパティにアクセスします。

**AsyncResult** オブジェクトは、次のメソッドの _callback_ パラメーターに引数として渡される関数から使用できます。



|**親オブジェクト**|**メソッド**|
|:-----|:-----|
|**Document** (Excel、PowerPoint、Project、および Word のみ)|[getSelectedDataAsync](../../reference/shared/document.getselecteddataasync.md)|
||[setSelectedDataAsync](../../reference/shared/document.setselecteddataasync.md)|
|**Bindings** (Excel および Word のみ)|[addFromPromptAsync](../../reference/shared/bindings.addfrompromptasync.md)|
||[addFromSelectionAsync](../../reference/shared/bindings.addfromselectionasync.md)|
||[getAllAsync](../../reference/shared/bindings.getallasync.md)|
||[getByIdAsync](../../reference/shared/bindings.getbyidasync.md)|
||[releaseByIdAsync](../../reference/shared/bindings.releasebyidasync.md)|
|**Binding** (Excel および Word のみ)|[getDataAsync](../../reference/shared/binding.getdataasync.md)|
||[setDataAsync](../../reference/shared/binding.setdataasync.md)|
||[removeHandlerAsync](../../reference/shared/binding.removehandlerasync.md)|
|**TableBinding** (Excel および Word のみ)||
||[addRowsAsync](../../reference/shared/binding.tablebinding.addrowsasync.md)|
||[deleteAllDataValuesAsync](../../reference/shared/binding.tablebinding.deletealldatavaluesasync.md)|
|**Settings** (Excel、PowerPoint、および Word のみ)|[refreshAsync](../../reference/shared/settings.refreshasync.md)|
||[saveAsync](../../reference/shared/settings.saveasync.md)|
|**CustomXmlNode** (Word のみ)|[getNodesAsync](../../reference/shared/customxmlnode.getnodesasync.md)|
||[getNodeValueAsync](../../reference/shared/customxmlnode.getnodevalueasync.md)|
||[getXmlAsync](../../reference/shared/customxmlnode.getxmlasync.md)|
||[setNodeValueAsync](../../reference/shared/customxmlnode.setnodevalueasync.md)|
||[setXmlAsync](../../reference/shared/customxmlnode.setxmlasync.md)|
|**CustomXmlPart** (Word のみ)|[deleteAsync](../../reference/shared/customxmlpart.deleteasync.md)|
||[getNodesAsync](../../reference/shared/customxmlpart.getnodesasync.md)|
||[getXmlAsync](../../reference/shared/customxmlpart.getxmlasync.md)|
|**CustomXmlParts** (Word のみ)|[addAsync](../../reference/shared/customxmlparts.addasync.md)|
||[getByIdAsync](../../reference/shared/customxmlparts.getbyidasync.md)|
||[getByNamespaceAsync](../../reference/shared/customxmlparts.getbynamespaceasync.md)|
|**CustomXmlPrefixMappings** (Word のみ)|[addNamespaceAsync](../../reference/shared/customxmlprefixmappings.addnamespaceasync.md)|
||[getNamespaceAsync](../../reference/shared/customxmlprefixmappings.getnamespaceasync.md)|
||[getPrefixAsync](../../reference/shared/customxmlprefixmappings.getprefixasync.md)|
|**Mailbox** (Outlook のみ)|[getUserIdentityTokenAsync](http://msdn.microsoft.com/library/c658518b-6867-41a0-99cf-810303e4c539%28Office.15%29.aspx)|
||[makeEwsRequestAsync](http://msdn.microsoft.com/library/2ec380e0-4a67-4146-92a6-6a39f65dc6f2%28Office.15%29.aspx)|
|**CustomProperties** (Outlook のみ)|[saveAsync](http://msdn.microsoft.com/library/690d5aa9-62b5-4e5c-9548-62dfdbb5fa56%28Office.15%29.aspx)|
|**Item** (Outlook のみ)|[loadCustomPropertiesAsync](http://msdn.microsoft.com/library/dfbec151-8ea7-4915-b723-09ea1396a261%28Office.15%29.aspx)|
|**RoamingSettings** (Outlook のみ)|[saveAsync](http://msdn.microsoft.com/library/a616f71c-a447-423f-a0d2-e9d6f1ac32f8%28Office.15%29.aspx)|

## サポートの詳細


次の表で、大文字 Y は、このメソッドは、対応する Office ホスト アプリケーションでサポートされていることを示します。空のセルは、Office ホスト アプリケーションでこのメソッドをサポートしないことを示します。

Office ホスト アプリケーションとサーバーの要件の詳細については、「[Office アドインを実行するための要件](../../docs/overview/requirements-for-running-office-add-ins.md)」をご覧ください。



| |**Windows デスクトップ版 Office**|**Office Online (ブラウザー)**|**Office for iPad**|**デバイス用 OWA**|**Outlook for Mac**|
|:-----|:-----|:-----|:-----|:-----|:-----|
|**Access**||Y||||
|**Excel**|Y|Y|Y|||
|**Outlook**|Y|Y||Y|Y|
|**PowerPoint**|Y|Y|Y|||
|**Project**|Y|||||
|**Word**|Y|Y|Y|||

|||
|:-----|:-----|
|**アドインの種類**|コンテンツ、作業ウィンドウ、Outlook|
|**ライブラリ**|Office.js|
|**名前空間**|Office|

## サポート履歴



****


|**変更内容**|**1.1**|
|:-----|:-----|
|1.1|Office for iPad で Excel、PowerPoint、および Word のサポートが追加されました。|
|1.1|Access 用のアドインのサポートが追加されました。|
|1.0|導入|
