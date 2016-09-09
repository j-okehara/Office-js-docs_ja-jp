
# Context.displayLanguage プロパティ
Office ホスト アプリケーションの UI 用にユーザーが指定したロケール (言語) を取得します。

|||
|:-----|:-----|
|**ホスト:**|Access、Excel、Outlook、PowerPoint、Project、Word|
|**最終変更バージョン**|1.1|

```
var myDisplayLanguage = Office.context.displayLanguage;
```


## 戻り値

RFC 1766 言語タグ形式の **string** ( `en-US` など) です。


## 解説 

 
**displayLanguage** の値は、Office ホスト アプリケーションの **[ファイル]** 、 **[オプション]** 、 **[言語]** の順に指定した現在の **[表示言語]** 設定に反映されます。 
 

Access Web アプリのコンテンツ アドインで、  **displayLanguage** プロパティはアドインの言語 (たとえば「en-US」) を取得します。


## 例




```js
function sayHelloWithDisplayLanguage() {
    var myDisplayLanguage = Office.context.displayLanguage;
    switch (myDisplayLanguage) {
        case 'en-US':
            write('Hello!');
            break;
        case 'en-NZ':
            write('G\'day mate!');
            break;
    }
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```




## サポートの詳細


次の表で、大文字 Y は、このメソッドは、対応する Office ホスト アプリケーションでサポートされていることを示します。空のセルは、Office ホスト アプリケーションでこのメソッドをサポートしないことを示します。

Office ホスト アプリケーションとサーバーの要件の詳細については、「[Office アドインを実行するための要件](../../docs/overview/requirements-for-running-office-add-ins.md)」をご覧ください。


||**Windows デスクトップ版 Office**|**Office Online (ブラウザー)**|**Office for iPad**|**Outlook for Mac**|
|:-----|:-----|:-----|:-----|:-----|
|**Access**||Y|||
|**Excel**|Y|Y|Y||
|**Outlook**|Y|Y||Y|
|**PowerPoint**|Y|Y|Y||
|**Project**|Y||||
|**Word**|Y|Y|Y||

|||
|:-----|:-----|
|**最小限のアクセス許可レベル**|[Restricted](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**アドインの種類**|コンテンツ、作業ウィンドウ、Outlook|
|**ライブラリ**|Office.js|
|**名前空間**|Office|

## サポート履歴



****


|**変更内容**|**1.1**|
|:-----|:-----|
|1.1|PowerPoint Online のサポートが追加されました。|
|1.1|Office for iPad で Excel、PowerPoint、および Word のサポートが追加されました。|
|1.1|Access 用コンテンツ アドインで、この API へのアクセスが追加されました。|
|1.0|導入|
