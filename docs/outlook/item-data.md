
# 読み取りまたは新規作成フォームの Outlook アイテム データを取得および設定する

Office アドイン マニフェスト スキーマのバージョン 1.1 以降、Outlook は、アイテムの表示または作成時にアドインをアクティブ化することができます。アドインが閲覧フォームまたは新規作成フォームのどちらでアクティブ化されるかによって、アイテムでアドインが使用できるプロパティも異なります。たとえば、 [dateTimeCreated](../../reference/outlook/Office.context.mailbox.item.md) プロパティと [dateTimeModified](../../reference/outlook/Office.context.mailbox.item.md) プロパティは送信済みのアイテム (アイテムは、その後閲覧フォームで表示されます) のみに定義され、(新規作成フォームで) メッセージを作成する場合これらのプロパティは定義されません。もう 1 つの例は [bcc](../../reference/outlook/Office.context.mailbox.item.md) プロパティです。このプロパティは、(新規作成フォームで) メッセージを作成する場合にのみ使用でき、閲覧フォームでは使用できません。

表 1 は、メール アドインの閲覧モードと新規作成モードのそれぞれで使用可能な JavaScript API for Office のアイテムレベルのプロパティを示しています。通常、閲覧フォームで使用可能なプロパティは読み取り専用で、新規作成フォームで使用可能なプロパティは値の取得および設定を行えます。ただし、例外として、 [itemId](../../reference/outlook/Office.context.mailbox.item.md) と [conversationId](../../reference/outlook/Office.context.mailbox.item.md) プロパティがあります。これらのプロパティは、フォームに関係なく読み取り専用です。新規作成フォームで使用可能な残りのアイテムレベルのプロパティは、アドインとユーザーが同時に同じプロパティの読み取りまたは書き込みを行う可能性があるため、新規作成モードでこれらのプロパティの取得や設定を行うメソッドは非同期です。このため、これらのプロパティが返すオブジェクトの種類も、新規作成フォームと閲覧フォームとで異なります。新規作成フォームで非同期のメソッドを使用してアイテムレベルのプロパティを取得または設定することについて詳しくは、「 [Outlook で新規作成フォームのアイテム データを取得および設定する](../outlook/get-and-set-item-data-in-a-compose-form.md)」をご覧ください。


**表 1. 新規作成フォームと閲覧フォームで使用できるアイテムのプロパティ**


|**アイテムの種類**|**プロパティ**|**閲覧フォームにおけるプロパティのタイプ**|**新規作成フォームにおけるプロパティのタイプ**|
|:-----|:-----|:-----|:-----|
|予定とメッセージ|[dateTimeCreated](../../reference/outlook/Office.context.mailbox.item.md)|JavaScript  **Date** オブジェクト|このプロパティは使用できません|
|予定とメッセージ|[dateTimeModified](../../reference/outlook/Office.context.mailbox.item.md)|JavaScript  **Date** オブジェクト|このプロパティは使用できません|
|予定とメッセージ|[itemClass](../../reference/outlook/Office.context.mailbox.item.md)|String|このプロパティは使用できません|
|予定とメッセージ|[itemId](../../reference/outlook/Office.context.mailbox.item.md)|String|このプロパティは使用できません|
|予定とメッセージ|[itemType](../../reference/outlook/Office.context.mailbox.item.md)|[ItemType](../../reference/outlook/Office.MailboxEnums.md) 列挙型の文字列|このプロパティは使用できません|
|予定とメッセージ|[添付ファイル](../../reference/outlook/Office.context.mailbox.item.md)|[AttachmentDetails](../../reference/outlook/simple-types.md)|このプロパティは使用できません|
|予定とメッセージ|[body](../../reference/outlook/Office.context.mailbox.item.md)|[本文](../../reference/outlook/Body.md)|[本文](../../reference/outlook/Body.md)|
|予定|[end](../../reference/outlook/Office.context.mailbox.item.md)|JavaScript  **Date** オブジェクト|[時刻](../../reference/outlook/Time.md)|
|予定|[location](../../reference/outlook/Office.context.mailbox.item.md)|String|[Location](../../reference/outlook/Location.md)|
|予定とメッセージ|[normalizedSubject](../../reference/outlook/Office.context.mailbox.item.md)|String|このプロパティは使用できません|
|予定|[optionalAttendees](../../reference/outlook/Office.context.mailbox.item.md)|[EmailAddressDetails](../../reference/outlook/simple-types.md)|[受信者](../../reference/outlook/Recipients.md)|
|予定|[organizer](../../reference/outlook/Office.context.mailbox.item.md)|EmailAddressDetails|このプロパティは使用できません|
|予定|[requiredAttendees](../../reference/outlook/Office.context.mailbox.item.md)|EmailAddressDetails|受信者|
|予定|[リソース](../../reference/outlook/Office.context.mailbox.item.md)|String|このプロパティは使用できません|
|予定|[開始](../../reference/outlook/Office.context.mailbox.item.md)|JavaScript  **Date** オブジェクト|時刻|
|予定とメッセージ|[件名](../../reference/outlook/Office.context.mailbox.item.md)|String|[件名](../../reference/outlook/Subject.md)|
|メッセージ|[bcc](../../reference/outlook/Office.context.mailbox.item.md)|このプロパティは使用できません|受信者|
|メッセージ|[cc](../../reference/outlook/Office.context.mailbox.item.md)|EmailAddressDetails|受信者|
|メッセージ|[conversationId](../../reference/outlook/Office.context.mailbox.item.md)|String|文字列 (読み取り専用)|
|メッセージ|[from](../../reference/outlook/Office.context.mailbox.item.md)|EmailAddressDetails|このプロパティは使用できません|
|メッセージ|[internetMessageId](../../reference/outlook/Office.context.mailbox.item.md)|整数|このプロパティは使用できません|
|メッセージ|[sender](../../reference/outlook/Office.context.mailbox.item.md)|EmailAddressDetails|このプロパティは使用できません|
|メッセージ|[to](../../reference/outlook/Office.context.mailbox.item.md)|EmailAddressDetails|受信者|

## Exchange Server のコールバックのトークンを閲覧アドインから使用する


Outlook アドインが閲覧フォームでアクティブ化される場合、Exchange コールバック トークンを取得できます。このトークンをサーバー側のコードで使用して、Exchange Web Services (EWS) を介してアイテムのすべてにアクセスできます。アドイン マニフェストで  **ReadItem** のアクセス許可を指定すると、 [mailbox.getCallbackTokenAsync](../../reference/outlook/Office.context.mailbox.md) メソッドを使用した Exchange コールバック トークンの取得、 [mailbox.ewsUrl](../../reference/outlook/Office.context.mailbox.md) プロパティを使用したユーザーのメールボックスの EWS エンドポイントの URL の取得、 [item.itemId](../../reference/outlook/Office.context.mailbox.item.md) による選択したアイテムの EWS ID の取得を行えます。その後、コールバック トークン、EWS エンドポイントの URL、EWS アイテム ID をサーバー側のコードに渡して [GetItem](http://msdn.microsoft.com/en-us/library/e3590b8b-c2a7-4dad-a014-6360197b68e4%28Office.15%29.aspx) の操作にアクセスし、アイテムのその他のプロパティを取得することができます。


## 閲覧アドインまたは新規作成アドインからのアクセス



  [mailbox.makeEwsRequestAsync](../../reference/outlook/Office.context.mailbox.md) メソッドを使用すると、Exchange Web Services (EWS) の操作である [GetItem](http://msdn.microsoft.com/en-us/library/e3590b8b-c2a7-4dad-a014-6360197b68e4%28Office.15%29.aspx) および [UpdateItem](http://msdn.microsoft.com/en-us/library/5d027523-e0bc-4da2-b60b-0cb9fc1fdfe4%28Office.15%29.aspx) にアドインから直接アクセスすることもできます。これらの操作を使用して、指定したアイテムの多数のプロパティを取得および設定できます。このメソッドは、アドイン マニフェストで **ReadWriteMailbox** のアクセス許可が指定されている限り、アドインが閲覧フォームまたは新規作成フォームのどちらでアクティブ化されたかに関係なく、Outlook アドインで使用できます。 **makeEwsRequestAsync** を使用した EWS の操作へのアクセスについて詳しくは、「 [Outlook アドインから Web サービスを呼び出す](../outlook/web-services.md)」をご覧ください。


## その他のリソース



- [Outlook アドイン](../outlook/outlook-add-ins.md)
    
- [Outlook で新規作成フォームのアイテム データを取得および設定する](../outlook/get-and-set-item-data-in-a-compose-form.md)
    
- [Outlook アドインから Web サービスを呼び出す](../outlook/web-services.md)
    


