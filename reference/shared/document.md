
# <a name="document-object"></a>Document オブジェクト
アドインから対話操作するドキュメントを表す抽象クラス。

|||
|:-----|:-----|
|**ホスト:**|Access、Excel、PowerPoint、Project、Word|
|**追加されたバージョン**|1.0|
|**最終変更バージョン**|1.1|

```
Office.context.document
```


## <a name="members"></a>メンバー


**プロパティ**


|**名前**|**説明**|**サポートのメモ**|
|:-----|:-----|:-----|
|[bindings](../../reference/shared/document.bindings.md)|ドキュメントに定義されているバインドへのアクセスを提供するオブジェクトを取得します。|1.1 で、Access 用コンテンツ アドインのサポートが追加されました。|
|[customXmlParts](../../reference/shared/document.customxmlparts.md)|ドキュメント内のカスタム XML パーツを表すオブジェクトを取得します。||
|[mode](../../reference/shared/document.mode.md)|ドキュメントのモードを取得します。|1.1 で、Access 用コンテンツ アドインのサポートが追加されました。|
|[settings](../../reference/shared/document.settings.md)|現在のドキュメントのコンテンツ アプリまたは作業ウィンドウ アプリの保存されているカスタム設定を表すオブジェクトを取得します。|1.1 で、Access 用コンテンツ アドインのサポートが追加されました。|
|[url](../../reference/shared/document.url.md)|ホスト アプリケーションが現在開いているドキュメントの URL を取得します。|1.1 で、Access 用コンテンツ アドインのサポートが追加されました。|

**メソッド**


|**名前**|**説明**|**サポートのメモ**|
|:-----|:-----|:-----|
|[addHandlerAsync](../../reference/shared/document.addhandlerasync.md)|**Document** オブジェクト イベントのイベント ハンドラーを追加します。||
|[getActiveViewAsync](../../reference/shared/document.getactiveviewasync.md)|プレゼンテーションの現在のビューを返します。|1.1 で、[PowerPoint 用アドイン](../../docs/powerpoint/powerpoint-add-ins.md)のサポートが追加されました。|
|[getFileAsync](../../reference/shared/document.getfileasync.md)|ドキュメント ファイル全体を、最大で 4194304 バイト (4 MB) のスライスに分割して返します。|1.1 で、PowerPoint および Word 用アドインで PDF としてファイルを取得するサポートが追加されました。|
|[getFilePropertiesAsync](../../reference/shared/document.getfilepropertiesasync.md)|現在のドキュメントのファイル プロパティを取得します。このリリースでは、ドキュメントの URL のみが取得できます。|1.1 で、Excel、Word、および PowerPoint のアドインでのドキュメントの URL の取得が追加されました。|
|[getSelectedDataAsync](../../reference/shared/document.getselecteddataasync.md)|ドキュメントの現在の選択範囲に含まれるデータを読み取ります。|1.1 で、PowerPoint 用アドインでの選択範囲のスライドの ID、タイトル、インデックスを取得するためのサポートが追加されました。|
|[goToByIdAsync](../../reference/shared/document.gotobyidasync.md)|ドキュメント内の指定されたオブジェクトまたは場所に移動します。|1.1 で、Excel 用アドインと PowerPoint 用アドインでのドキュメント内のナビゲーションに対するサポートが追加されました。|
|[removeHandlerAsync](../../reference/shared/document.removehandlerasync.md)|**Document** オブジェクト イベントのイベント ハンドラーを削除します。||
|[setSelectedDataAsync](../../reference/shared/document.setselecteddataasync.md)|ドキュメント内の現在の選択範囲にデータを書き込みます。|1.1 で、[Excel 用アドインでデータを書き込む際に、選択しているテーブルの書式を設定する](../../docs/excel/format-tables-in-add-ins-for-excel.md)サポートが追加されました。|

**イベント**


|**名前**|**説明**|**サポートのメモ**||
|:-----|:-----|:-----|:-----|
|[ActiveViewChanged](../../reference/shared/document.activeviewchanged.md)|ユーザーがドキュメントの現在のビューを変更したときに発生します。|1.1 で、PowerPoint 用アドインのサポートが追加されました。||
|[SelectionChanged](../../reference/shared/document.selectionchanged.event.md)|ドキュメント内で選択が変更されるときに発生します。|||

## <a name="remarks"></a>注釈

**Document** オブジェクトをスクリプトで直接インスタンス化することはありません。**Document** オブジェクトのメンバーを呼び出して現在のドキュメントまたはワークシートを操作するには、`Office.context.document` をスクリプトで使用します。


## <a name="example"></a>例

次の例は、 **Document** オブジェクトの **getSelectedDataAsync** メソッドを使用して、ユーザーの現在の選択範囲をテキストとして取得し、それをアプリのページに表示します。


```js

// Display the user's current selection.
function showSelection() {
    Office.context.document.getSelectedDataAsync(
        "text",                        // coercionType
        {valueFormat: "unformatted",   // valueFormat
        filterType: "all"},            // filterType
        function (result) {            // callback
            var dataValue; 
            dataValue = result.value;
            write('Selected data is: ' + dataValue);
        });
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```




## <a name="support-details"></a>要件


**Document** オブジェクトの各 API メンバーのサポートは、Office のホスト アプリケーションの間で異なります。各メンバーのホストのサポート情報のトピックにある「サポートの詳細」セクションを参照してください。

Office ホスト アプリケーションとサーバーの要件の詳細については、「[Office アドインを実行するための要件](../../docs/overview/requirements-for-running-office-add-ins.md)」をご覧ください。


|||
|:-----|:-----|
|**追加されたバージョン**|1.0|
|**最終変更バージョン**|1.1|
|**アドインの種類**|コンテンツ、作業ウィンドウ|
|**ライブラリ**|Office.js|
|**名前空間**|Office|
