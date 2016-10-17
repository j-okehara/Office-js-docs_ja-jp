
# <a name="document.customxmlparts-property"></a>Document.customXmlParts プロパティ
ドキュメント内のカスタム XML パーツを表すオブジェクトを取得します。

|||
|:-----|:-----|
|**ホスト:**|Word|
|**追加されたバージョン**|1.1|

```js
var xmlParts = Office.context.document.customXmlParts;
```


## <a name="return-value"></a>戻り値

[CustomXmlParts](../../reference/shared/customxmlparts.customxmlparts.md) オブジェクト。


## <a name="example"></a>例




```js
function getCustomXmlParts(){
    Office.context.document.customXmlParts.getByNamespaceAsync('http://tempuri.org', function (asyncResult) {
        write('Retrieved ' + asyncResult.value.length + ' custom XML parts');
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
|**Word**|Y||Y|

|||
|:-----|:-----|
|**最小限のアクセス許可レベル**|[制限あり](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**アドインの種類**|作業ウィンドウ|
|**ライブラリ**|Office.js|
|**名前空間**|Office|

## <a name="support-history"></a>サポート履歴



****


|**バージョン**|**変更内容**|
|:-----|:-----|
|1.1|Office for iPad での Word のサポートが追加されました。|
|1.0|導入|
