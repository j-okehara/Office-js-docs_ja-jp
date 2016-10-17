
# <a name="customxmlnode.settextasync-method"></a>CustomXmlNode.setTextAsync メソッド
カスタム XML パーツ内の XML ノードのテキストを非同期的に設定します。

|||
|:-----|:-----|
|**ホスト:**|Word|
|**[要件セット](../../docs/overview/specify-office-hosts-and-api-requirements.md)に指定できるもの**|CustomXmlParts|
|**追加されたバージョン**|1.2|

```
customXmlNodeObj.setTextAsync(text, [asyncContext,]callback(asyncResult);
```


## <a name="parameters"></a>パラメーター



|**名前**|**型**|**説明**|
|:-----|:-----|:-----|
| _text_|**string**|必須。XML ノードのテキスト値。|
| _asyncContext_|**object**|オプション。[AsyncResult](../../reference/shared/asyncresult.md) オブジェクトの asyncContext プロパティで取得できるユーザー定義のオブジェクト。これは、コールバックが名前付き関数の場合に、 **AsyncResult** にオブジェクトまたは値を提供するために使用します。|
| _callback_|**object**|省略可能。コールバックが戻るときに呼び出される関数で、唯一のパラメーターは **AsyncResult** 型です。|

## <a name="callback-value"></a>コールバック値

_callback_ パラメーターに渡した関数が実行されると、[AsyncResult](../../reference/shared/asyncresult.md) オブジェクトを受け取ります。このオブジェクトには、コールバック関数の唯一のパラメーターからアクセスできます。

**setTextAsync** メソッドに渡されたコールバック関数で、**AsyncResult** オブジェクトのプロパティを使用して次の情報を戻せます。



|**プロパティ**|**使用目的**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|使用しません。|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|操作の成功または失敗を示します。|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|操作が失敗した場合、エラーに関する情報を提供する [Error](../../reference/shared/error.md) オブジェクトにアクセスします。|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|ユーザー定義の  **object** または値を _asyncContext_ パラメーターとして渡した場合、そのオブジェクトまたは値にアクセスします。このプロパティは、 _asyncContext_ が設定されていない場合に未定義を返します。|

## <a name="example"></a>例

カスタム XML パーツ内のノードのテキスト値を設定する方法について説明します。


```js
// Get the built-in core properties XML part by using its ID. This results in a call to Word.
Office.context.document.customXmlParts.getByIdAsync("{6C3C8BC8-F283-45AE-878A-BAB7291924A1}", function (getByIdAsyncResult) {
    
    // Access the XML part.
    var xmlPart = getByIdAsyncResult.value;
    
    // Add namespaces to the namespace manager. These two calls result in two calls to Word.
    xmlPart.namespaceManager.addNamespaceAsync('cp', 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties', function () {
        xmlPart.namespaceManager.addNamespaceAsync('dc', 'http://purl.org/dc/elements/1.1/', function () {

            // Get XML nodes by using an Xpath expression. This results in a call to the host.
            xmlPart.getNodesAsync("/cp:coreProperties/dc:subject", function (getNodesAsyncResult) {
                
                // Get the first node returned by using the Xpath expression. This will be the subject element in this example.
                var subjectNode = getNodesAsyncResult.value[0];
                
                // Set the text value of the subject node and use the asyncContext. This results in a call to the host. 
                // The results are logged to the browser console. 
                subjectNode.setTextAsync("newSubject", {asyncContext: "StateNormal"}, function (setTextAsyncResult) {
                   console.log("The status of the call: " + setTextAsyncResult.status);
                   console.log("The asyncContext value = " + setTextAsyncResult.asyncContext);
                });
            });
        });
    });
});
```


## <a name="support-details"></a>サポートの詳細


次の表で、大文字 Y は、このメソッドは、対応する Office ホスト アプリケーションでサポートされていることを示します。空のセルは、Office ホスト アプリケーションでこのメソッドをサポートしないことを示します。

Office ホスト アプリケーションとサーバーの要件の詳細については、「[Office アドインを実行するための要件](../../docs/overview/requirements-for-running-office-add-ins.md)」をご覧ください。


||**Windows デスクトップ版 Office**|**Office Online (ブラウザー)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Word**|Y|Y|Y|

|||
|:-----|:-----|
|**要件セットに指定できるもの**|CustomXmlParts|
|**最小限のアクセス許可レベル**|[ReadWriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**アドインの種類**|作業ウィンドウ|
|**ライブラリ**|Office.js|
|**名前空間**|Office|

## <a name="support-history"></a>サポート履歴



****


|**バージョン**|**変更内容**|
|:-----|:-----|
|1.1|setTextAsync を追加しました。|
