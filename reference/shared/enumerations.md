
# <a name="enumerations"></a>列挙型

完全修飾列挙名 ( `Office.CoercionType.Text`) または対応するテキスト値 ( `"text"`) を使用して列挙値を指定できます。たとえば、次のメソッド呼び出しでは列挙名を使用しています。


```js
Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, {valueFormat:Office.ValueFormat.Unformatted, filterType:Office.FilterType.All},
   function (result) {
      if (result.status === Office.AsyncResultStatus.Success)
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


次の例では、同じ呼び出しに列挙テキスト値を使用しています。




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
```


## <a name="reference"></a>参照



|**名前**|**定義**|
|:-----|:-----|
|[ActiveView](activeview-enumeration.md)|ユーザーがドキュメントを編集できるかどうかなど、ドキュメントのアクティブなビューの状態を指定します。|
|[AsyncResultStatus](asyncresultstatus-enumeration.md)|非同期呼び出しの結果を指定します。|
|
  [AttachmentType](http://msdn.microsoft.com/library/83883a47-a937-4afb-a55e-e789057335c4%28Office.15%29.aspx)|電子メール メッセージまたは会議出席依頼の添付ファイルの種類を指定します。Outlook 2013 はこの列挙をサポートしていません。|
|[BindingType](bindingtype-enumeration.md)|返されるバインド オブジェクトの種類を指定します。|
|
  [BodyType](http://msdn.microsoft.com/library/31350fe6-4c42-4cbb-a5b2-4fb2d360fa11%28Office.15%29.aspx)|予定またはメッセージの本文のテキストの種類を指定します。|
|[CoercionType](coerciontype-enumeration.md)|呼び出されたメソッドによって返される、または設定されるデータを強制的に変換する方法を指定します。|
|[CustomXMLNodeType](customxmlnodetype-enumeration.md)|ノードの種類を指定します。|
|[DocumentMode](documentmode-enumeration.md)|関連付けられているアプリケーションのドキュメントを読み取り専用または読み取り/書き込みのどちらかに指定します。 |
|
  [EntityType](http://msdn.microsoft.com/library/0035be38-8a65-4693-bcc4-0a8dd7b1495b%28Office.15%29.aspx)|エンティティの種類を指定します。|
|[EventType](eventtype-enumeration.md)|発生したイベントの種類を指定します。|
|[FileType](filetype-enumeration.md)|ドキュメントを返す形式を指定します。|
|[GoToType](gototype-enumeration.md)|ナビゲートする場所またはオブジェクトの種類を指定します。|
|[FilterType](filtertype-enumeration.md)|データを取得するときにホスト アプリケーションからのフィルタリングを適用するかどうかを指定します。|
|[InitializationReason](initializationreason-enumeration.md)|ドキュメントにアドインが挿入されたばかりであるか、既に含まれていたかを指定します。|
|
  [ItemType](http://msdn.microsoft.com/library/e0bb23fd-f360-4b0f-b72c-1cf08d4cab3f%28Office.15%29.aspx)|アイテムの種類を指定します。|
|
  [notificationMessageType](http://msdn.microsoft.com/library/ff00c89d-0019-4545-a95b-7ed0db712ce9%28Office.15%29.aspx)|予定またはメッセージの通知メッセージを指定します。|
|[ProjectProjectFields](projectprojectfields-enumeration.md)|[getProjectFieldAsync](projectdocument.getprojectfieldasync.md) メソッドのパラメーターとして使用できるプロジェクト フィールドを指定します。|
|[ProjectResourceFields](projectresourcefields-enumeration.md)|[getResourceFieldAsync](projectdocument.gettaskfieldasync.md) メソッドのパラメーターとして使用できるリソース フィールドを指定します。|
|[ProjectTaskFields](projecttaskfields-enumeration.md)|[getTaskFieldAsync](projectdocument.gettaskfieldasync.md) メソッドのパラメーターとして使用できるタスク フィールドを指定します。|
|[ProjectViewTypes](projectviewtypes-enumeration.md)|[getSelectedViewAsync](projectdocument.getselectedviewasync.md) メソッドで認識できるビューの種類を指定します。|
|
  [RecipientType](http://msdn.microsoft.com/library/6e7c4029-6e52-47f6-98d2-4cd3ce7bd8b4%28Office.15%29.aspx)|予定の受信者の種類を指定します。|
|
  [ResponseType](http://msdn.microsoft.com/library/b3e723ca-4be0-4846-ad97-0eecab4355eb%28Office.15%29.aspx)|会議招待への応答を指定します。|
|[SelectionMode](selectionmode-enumeration.md)|移動先の場所を選択 (強調表示) するかどうかを指定します ([Document.goToByIdAsync](document.gotobyidasync.md) メソッドを使用する場合)。|
|
  [SourceProperty](http://msdn.microsoft.com/library/6a209a7f-57cd-4dc3-869e-07b0f5928b28%28Office.15%29.aspx)|呼び出されたメソッドによって返されるデータのソースを指定します。|
|[Table](table-enumeration.md)|[テーブルの書式設定メソッド](../../docs/excel/format-tables-in-add-ins-for-excel.md)の _cellFormat_ パラメーターの `cells:` プロパティに列挙値を指定します。|
|[ValueFormat](valueformat-enumeration.md)|呼び出されたメソッドが返す値 (数字、日付など) を、書式設定して返すかどうかを指定します。|

## <a name="support-details"></a>サポートの詳細


各列挙のサポートは、Office ホスト アプリケーション間で異なります。ホストのサポート情報については、各列挙のトピックの「サポートの詳細」セクションを参照してください。

Office ホスト アプリケーションとサーバーの要件の詳細については、「[Office アドインを実行するための要件](../../docs/overview/requirements-for-running-office-add-ins.md)」をご覧ください。


|||
|:-----|:-----|
|**アドインの種類**|コンテンツ、作業ウィンドウ、Outlook|
|**ライブラリ**|Office.js|
|**名前空間**|Office|
