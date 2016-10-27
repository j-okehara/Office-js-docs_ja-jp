

# <a name="item"></a>item

## [Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). item

`item` の名前空間を使用して、現在選択されているメッセージ、会議出席依頼、または予定にアクセスします。[itemType](Office.context.mailbox.item.md#itemtype-officemailboxenumsitemtype) プロパティを使用して、`item` の種類を指定できます。

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](../tutorial-api-requirement-sets.md)| 1.0|
|[最小限のアクセス許可レベル](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| 制限あり|
|適用可能な Outlook のモード| 作成または読み取り|

### <a name="example"></a>例

次の JavaScript のコード例は、Outlook の現在のアイテムの `subject` プロパティにアクセスする方法を示しています。

```JavaScript
// The initialize function is required for all apps.
Office.initialize = function () {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    var item = Office.context.mailbox.item;
    var subject = item.subject;
    // Continue with processing the subject of the current item,
    // which can be a message or appointment.
    });
}
```

### <a name="members"></a>メンバー

#### <a name="attachments-:array.<[attachmentdetails](simple-types.md#attachmentdetails)>"></a>attachments :Array.<[AttachmentDetails](simple-types.md#attachmentdetails)>

アイテムの添付ファイルの配列を取得します。閲覧モードのみ。

##### <a name="type:"></a>型:

*   Array.<[AttachmentDetails](simple-types.md#attachmentdetails)>

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](../tutorial-api-requirement-sets.md)| 1.0|
|[最小限のアクセス許可レベル](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用可能な Outlook のモード| 読み取り|

##### <a name="example"></a>例

次のコードでは、現在のアイテムのすべての添付ファイルの詳細を含む HTML 文字列を作成します。

```JavaScript
var _Item = Office.context.mailbox.item;
var outputString = "";

if (_Item.attachments.length > 0) {
  for (i = 0 ; i < _Item.attachments.length ; i++) {
    var _att = _Item.attachments[i];
    outputString += "<BR>" + i + ". Name: ";
    outputString += _att.name;
    outputString += "<BR>ID: " + _att.id;
    outputString += "<BR>contentType: " + _att.contentType;
    outputString += "<BR>size: " + _att.size;
    outputString += "<BR>attachmentType: " + _att.attachmentType;
    outputString += "<BR>isInline: " + _att.isInline;
  }
}

// Do something with outputString
```

####  <a name="bcc-:[recipients](recipients.md)"></a>bcc :[Recipients](Recipients.md)

メッセージの BCC (ブラインド カーボン コピー) 行の受信者を取得または設定します。新規作成モードのみ。

##### <a name="type:"></a>型:

*   [Recipients](Recipients.md)

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](../tutorial-api-requirement-sets.md)| 1.1|
|[最小限のアクセス許可レベル](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用可能な Outlook のモード| 作成|

##### <a name="example"></a>例

```JavaScript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-:[body](body.md)"></a>body :[Body](Body.md)

アイテムの本文を操作するメソッドを提供するオブジェクトを取得します。

##### <a name="type:"></a>型:

*   [Body](Body.md)

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](../tutorial-api-requirement-sets.md)| 1.1|
|[最小限のアクセス許可レベル](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用可能な Outlook のモード| 作成または読み取り|
####  cc :Array.<[EmailAddressDetails](simple-types.md#emailaddressdetails)>|[Recipients](Recipients.md)

メッセージの CC (カーボン コピー) の受信者を取得または設定します。

##### <a name="read-mode"></a>閲覧モード

`cc` プロパティは、メッセージの **CC** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。

##### <a name="compose-mode"></a>新規作成モード

`cc` プロパティは、メッセージの **CC** 行にある受信者を操作するメソッドを提供する `Recipients` オブジェクトを返します。

##### <a name="type:"></a>型:

*   Array.<[EmailAddressDetails](simple-types.md#emailaddressdetails)> | [Recipients](Recipients.md)

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](../tutorial-api-requirement-sets.md)| 1.0|
|[最小限のアクセス許可レベル](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用可能な Outlook のモード| 作成または読み取り|

##### <a name="example"></a>例

```JavaScript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="(nullable)-conversationid-:string"></a>(nullable) conversationId :String

特定のメッセージが含まれている電子メールの会話の識別子を取得します。

メール アプリを閲覧フォームでアクティブ化するか、新規作成フォームの返信でアクティブ化すると、このプロパティで整数を取得することができます。その後、ユーザーが返信の件名を変更した場合、その返信の送信時にメッセージの会話 ID が変更され、以前に取得した値は適用されなくなります。

新規作成フォームで新しいアイテムに対してこのプロパティに null を取得します。ユーザーが件名を設定し、アイテムを保存する場合、`conversationId` プロパティは値を返します。

##### <a name="type:"></a>型:

*   String

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](../tutorial-api-requirement-sets.md)| 1.0|
|[最小限のアクセス許可レベル](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用可能な Outlook のモード| 作成または読み取り|
#### <a name="datetimecreated-:date"></a>dateTimeCreated :Date

アイテムが作成された日時を取得します。閲覧モードのみ。

##### <a name="type:"></a>型:

*   日付

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](../tutorial-api-requirement-sets.md)| 1.0|
|[最小限のアクセス許可レベル](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用可能な Outlook のモード| 読み取り|

##### <a name="example"></a>例

```JavaScript
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-:date"></a>dateTimeModified :Date

アイテムが最後に変更された日時を取得します。閲覧モードのみ。

##### <a name="type:"></a>型:

*   日付

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](../tutorial-api-requirement-sets.md)| 1.0|
|[最小限のアクセス許可レベル](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用可能な Outlook のモード| 読み取り|

##### <a name="example"></a>例

```JavaScript
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-:date|[time](time.md)"></a>end :Date|[Time](Time.md)

予定が終了する日時を取得または設定します。

`end` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](Office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) メソッドを使用して、end プロパティ値をクライアントのローカル日時に変換することができます。

##### <a name="read-mode"></a>閲覧モード

`end` プロパティは `Date` オブジェクトを返します。

##### <a name="compose-mode"></a>新規作成モード

`end` プロパティは `Time` オブジェクトを返します。

[`Time.setAsync`](Time.md#setasync) メソッドを使用して終了時刻を設定する場合、[`convertToUtcClientTime`](Office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。

##### <a name="type:"></a>型:

*   Date | [Time](Time.md)

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](../tutorial-api-requirement-sets.md)| 1.0|
|[最小限のアクセス許可レベル](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用可能な Outlook のモード| 作成または読み取り|

##### <a name="example"></a>例

次の例では、`Time` オブジェクトの [`setAsync`](Time.md#setasync) メソッドを使用して、新規作成モードで予定の終了時刻を設定します。

```JavaScript
var endTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used
  // in the callback
     asyncContext: {verb:"Set"}
}
Office.context.mailbox.item.end.setAsync(endTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function
    console.debug("End Time " + result.asyncContext.verb);
  }
});
```

#### <a name="from-:[emailaddressdetails](simple-types.md#emailaddressdetails)"></a>from :[EmailAddressDetails](simple-types.md#emailaddressdetails)

メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。

メッセージが代理人から送信された場合を除き、`from` プロパティと [`sender`](Office.context.mailbox.item.md#sender) プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。

##### <a name="type:"></a>型:

*   [EmailAddressDetails](simple-types.md#emailaddressdetails)

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](../tutorial-api-requirement-sets.md)| 1.0|
|[最小限のアクセス許可レベル](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用可能な Outlook のモード| 読み取り|
#### <a name="internetmessageid-:string"></a>internetMessageId :String

電子メール メッセージのインターネット メッセージ ID を取得します。閲覧モードのみ。

##### <a name="type:"></a>型:

*   String

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](../tutorial-api-requirement-sets.md)| 1.0|
|[最小限のアクセス許可レベル](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用可能な Outlook のモード| 読み取り|

##### <a name="example"></a>例

```JavaScript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-:string"></a>itemClass :String

選択されたアイテムの Exchange Web サービスのアイテム クラスを取得します。閲覧モードのみ。

`itemClass` プロパティには、選択したアイテムのメッセージ クラスを指定します。次に、メッセージまたは予定アイテムの既定のメッセージ クラスを示します。

| 種類 | 説明 | アイテム クラス |
| --- | --- | --- |
| 予定アイテム | アイテム クラス `IPM.Appointment` または `IPM.Appointment.Occurence` の予定表アイテムは次のとおりです。 | `IPM.Appointment`<br />`IPM.Appointment.Occurence` |
| メッセージ アイテム | これには、既定のメッセージ クラス `IPM.Note` を持つ電子メール メッセージ、および基本メッセージ クラスとして `IPM.Schedule.Meeting` を使用する会議出席依頼、返信、または取り消しが含まれます。 | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

既定のメッセージ クラスを拡張したカスタム メッセージ クラス (たとえば、カスタム予定表メッセージ クラス`IPM.Appointment.Contoso` など) を作成できます。

##### <a name="type:"></a>型:

*   String

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](../tutorial-api-requirement-sets.md)| 1.0|
|[最小限のアクセス許可レベル](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用可能な Outlook のモード| 読み取り|

##### <a name="example"></a>例

```JavaScript
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="(nullable)-itemid-:string"></a>(nullable) itemId :String

現在のアイテムの Exchange Web サービスのアイテム識別子を取得します。閲覧モードのみ。

`itemId` プロパティから返される識別子は、Exchange Web サービスのアイテム識別子と同じです。`itemId` プロパティは、Outlook のエントリ ID と同じではありません。

`itemId` プロパティは、新規作成モードでストアに保存されていないアイテムの `null` を返します。アイテム識別子が必要な場合、[`saveAsync`](Office.context.mailbox.item.md#saveAsync) メソッドを使用してアイテムをストアに保存できます。そうすると、コールバック関数の [`AsyncResult.value`](simple-types.md#asyncresult) パラメーターでアイテム識別子が返されます。

##### <a name="type:"></a>型:

*   String

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](../tutorial-api-requirement-sets.md)| 1.0|
|[最小限のアクセス許可レベル](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用可能な Outlook のモード| 読み取り|

##### <a name="example"></a>例

次のコードは、アイテム識別子の有無を確認します。`itemId` プロパティが `null` または `undefined` を返す場合、アイテムはストアに保存され、非同期の結果からアイテム識別子が取得されます。

```JavaScript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-:[office.mailboxenums.itemtype](office.mailboxenums.md#itemtype-string)"></a>itemType :[Office.MailboxEnums.ItemType](Office.MailboxEnums.md#itemtype-string)

インスタンスが表しているアイテムの種類を取得します。

`itemType` プロパティは、`ItemType` 列挙値の 1 つを返します。これは `item` オブジェクト インスタンスがメッセージと予定のどちらであるかを示すものです。

##### <a name="type:"></a>型:

*   [Office.MailboxEnums.ItemType](Office.MailboxEnums.md#itemtype-string)

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](../tutorial-api-requirement-sets.md)| 1.0|
|[最小限のアクセス許可レベル](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用可能な Outlook のモード| 作成または読み取り|

##### <a name="example"></a>例

```JavaScript
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-:string|[location](location.md)"></a>location :String|[Location](Location.md)

予定の場所を取得または設定します。

##### <a name="read-mode"></a>閲覧モード

`location` プロパティは、予定の場所を格納した文字列を返します。

##### <a name="compose-mode"></a>新規作成モード

`location` プロパティは予定の場所を取得または設定するために使用するメソッドを提供する `Location` オブジェクトを返します。

##### <a name="type:"></a>型:

*   String | [Location](Location.md)

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](../tutorial-api-requirement-sets.md)| 1.0|
|[最小限のアクセス許可レベル](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用可能な Outlook のモード| 作成または読み取り|

##### <a name="example"></a>例

```JavaScript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-:string"></a>normalizedSubject :String

すべてのプレフィックス (`RE:` や `FWD:` など) が削除されたアイテムの件名を取得します。閲覧モードのみ。

normalizedSubject プロパティは、アイテムの件名に電子メール プログラムによって標準のプレフィックス (`RE:` や `FW:` など) が追加されたものを取得します。これらのプレフィックスが付いたままの状態でアイテムの件名を取得するには、[`subject`](Office.context.mailbox.item.md#subject) プロパティを使用します。

##### <a name="type:"></a>型:

*   String

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](../tutorial-api-requirement-sets.md)| 1.0|
|[最小限のアクセス許可レベル](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用可能な Outlook のモード| 読み取り|

##### <a name="example"></a>例

```JavaScript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="optionalattendees-:array.<[emailaddressdetails](simple-types.md#emailaddressdetails)>|[recipients](recipients.md)"></a>optionalAttendees :Array.<[EmailAddressDetails](simple-types.md#emailaddressdetails)>|[Recipients](Recipients.md)

任意出席者の電子メール アドレスのリストを取得または設定します。

##### <a name="read-mode"></a>閲覧モード

`optionalAttendees` プロパティは、会議への各任意出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。

##### <a name="compose-mode"></a>新規作成モード

`optionalAttendees` プロパティは会議への任意出席者を取得および設定するためのメソッドを提供する `Recipients` オブジェクトを返します。

##### <a name="type:"></a>型:

*   Array.<[EmailAddressDetails](simple-types.md#emailaddressdetails)> | [Recipients](Recipients.md)

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](../tutorial-api-requirement-sets.md)| 1.0|
|[最小限のアクセス許可レベル](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用可能な Outlook のモード| 作成または読み取り|

##### <a name="example"></a>例

```JavaScript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-:[emailaddressdetails](simple-types.md#emailaddressdetails)"></a>organizer :[EmailAddressDetails](simple-types.md#emailaddressdetails)

指定の会議の会議開催者の電子メール アドレスを取得します。閲覧モードのみ。

##### <a name="type:"></a>型:

*   [EmailAddressDetails](simple-types.md#emailaddressdetails)

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](../tutorial-api-requirement-sets.md)| 1.0|
|[最小限のアクセス許可レベル](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用可能な Outlook のモード| 読み取り|

##### <a name="example"></a>例

```JavaScript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-:array.<[emailaddressdetails](simple-types.md#emailaddressdetails)>|[recipients](recipients.md)"></a>requiredAttendees :Array.<[EmailAddressDetails](simple-types.md#emailaddressdetails)>|[Recipients](Recipients.md)

必須出席者の電子メール アドレスのリストを取得または設定します。

##### <a name="read-mode"></a>閲覧モード

`requiredAttendees` プロパティは、会議への各必須出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。

##### <a name="compose-mode"></a>新規作成モード

`requiredAttendees` プロパティは会議への必須出席者を取得および設定するためのメソッドを提供する `Recipients` オブジェクトを返します。

##### <a name="type:"></a>型:

*   Array.<[EmailAddressDetails](simple-types.md#emailaddressdetails)> | [Recipients](Recipients.md)

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](../tutorial-api-requirement-sets.md)| 1.0|
|[最小限のアクセス許可レベル](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用可能な Outlook のモード| 作成または読み取り|

##### <a name="example"></a>例

```JavaScript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="resources-:[emailaddressdetails](simple-types.md#emailaddressdetails)"></a>resources :[EmailAddressDetails](simple-types.md#emailaddressdetails)

予定に必要なリソースを取得します。閲覧モードのみ。

##### <a name="type:"></a>型:

*   [EmailAddressDetails](simple-types.md#emailaddressdetails)

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](../tutorial-api-requirement-sets.md)| 1.0|
|[最小限のアクセス許可レベル](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用可能な Outlook のモード| 読み取り|
#### <a name="sender-:[emailaddressdetails](simple-types.md#emailaddressdetails)"></a>sender :[EmailAddressDetails](simple-types.md#emailaddressdetails)

電子メール メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。

メッセージが代理人から送信された場合を除き、[`from`](Office.context.mailbox.item.md#from) プロパティと `sender` プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。

##### <a name="type:"></a>型:

*   [EmailAddressDetails](simple-types.md#emailaddressdetails)

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](../tutorial-api-requirement-sets.md)| 1.0|
|[最小限のアクセス許可レベル](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用可能な Outlook のモード| 読み取り|

##### <a name="example"></a>例

```JavaScript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-:date|[time](time.md)"></a>start :Date|[Time](Time.md)

予定を開始する日時を取得または設定します。

`start` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](Office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) メソッドを使用して、値をクライアントのローカル日時に変換することができます。

##### <a name="read-mode"></a>閲覧モード

`start` プロパティは `Date` オブジェクトを返します。

##### <a name="compose-mode"></a>新規作成モード

`start` プロパティは `Time` オブジェクトを返します。

[`Time.setAsync`](Time.md#setasync) メソッドを使用して開始時刻を設定する場合、[`convertToUtcClientTime`](Office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。

##### <a name="type:"></a>型:

*   Date | [Time](Time.md)

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](../tutorial-api-requirement-sets.md)| 1.0|
|[最小限のアクセス許可レベル](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用可能な Outlook のモード| 作成または読み取り|

##### <a name="example"></a>例

次の例では、`Time` オブジェクトの [`setAsync`](Time.md#setasync) メソッドを使用して、新規作成モードで予定の開始時刻を設定します。

```JavaScript
var startTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used
  // in the callback
     asyncContext: {verb:"Set"}
}
Office.context.mailbox.item.start.setAsync(startTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function
    console.debug("Start Time " + result.asyncContext.verb);
  }
});
```

####  <a name="subject-:string|[subject](subject.md)"></a>subject :String|[Subject](Subject.md)

アイテムの件名フィールドに示される説明を取得または設定します。

`subject` プロパティは、電子メール サーバーによって送信されたアイテムの件名全体を取得または設定します。

##### <a name="read-mode"></a>閲覧モード

`subject` プロパティは文字列を返します。[`normalizedSubject`](Office.context.mailbox.item.md#normalizedsubject-string) プロパティを使用して、`RE:` や `FW:` のような先頭部分のすべてのプレフィックスを除去した件名を取得します。

```
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a>新規作成モード

`subject` プロパティは件名を取得および設定するためのメソッドを提供する `Subject` オブジェクトを返します。

```JavaScript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type:"></a>型:

*   String | [Subject](Subject.md)

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](../tutorial-api-requirement-sets.md)| 1.0|
|[最小限のアクセス許可レベル](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用可能な Outlook のモード| 作成または読み取り|
####  to :Array.<[EmailAddressDetails](simple-types.md#emailaddressdetails)>|[Recipients](Recipients.md)

電子メール メッセージの受信者を取得または設定します。

##### <a name="read-mode"></a>閲覧モード

`to` プロパティは、メッセージの **To** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。

##### <a name="compose-mode"></a>新規作成モード

`to` プロパティは、メッセージの **To** 行にある受信者を操作するメソッドを提供する `Recipients` オブジェクトを返します。

##### <a name="type:"></a>型:

*   Array.<[EmailAddressDetails](simple-types.md#emailaddressdetails)> | [Recipients](Recipients.md)

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](../tutorial-api-requirement-sets.md)| 1.0|
|[最小限のアクセス許可レベル](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用可能な Outlook のモード| 作成または読み取り|

##### <a name="example"></a>例

```JavaScript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a>メソッド

####  <a name="addfileattachmentasync(uri,-attachmentname,-[options],-[callback])"></a>addFileAttachmentAsync(uri, attachmentName, [options], [callback])

ファイルを添付ファイルとしてメッセージまたは予定に追加します。

`addFileAttachmentAsync` メソッドは、指定した URI にあるファイルをアップロードし、新規作成フォーム内のアイテムに添付します。

その後、[`removeAttachmentAsync`](Office.context.mailbox.item.md#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。

##### <a name="parameters:"></a>パラメーター:

|名前| 型| 属性| 説明|
|---|---|---|---|
|`uri`| String||メッセージまたは予定に添付するファイルの場所を示す URI。最大の長さは 2048 文字です。|
|`attachmentName`| String||添付ファイルをアップロードするときに表示される添付ファイルの名前です。最大の長さは 255 文字です。|
|`options`| Object| &lt;optional&gt;|次のプロパティのうち 1 つ以上を含むオブジェクト リテラルです。<br/><br/>**プロパティ**<br/><table class="nested-table"><thead><tr><th>名前</th><th>型</th><th>属性</th><th>説明</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;optional&gt;</td><td>開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</td></tr></tbody></table>|
|`callback`| function| &lt;optional&gt;|メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](simple-types.md#asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。 <br/>成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。<br/>添付ファイルのアップロードが失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。<br/><table class="nested-table"><thead><tr><th>エラー コード</th><th>説明</th></tr></thead><tbody><tr><td><code>AttachmentSizeExceeded</code></td><td>添付ファイルのサイズが上限を超えています。</td></tr><tr><td><code>FileTypeNotSupported</code></td><td>許可されていない拡張子の添付ファイルです。</td></tr><tr><td><code>NumberOfAttachmentsExceeded</code></td><td>メッセージまたは予定の添付ファイルが多すぎます。</td></tr></tbody></table>|

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](../tutorial-api-requirement-sets.md)| 1.1|
|[最小限のアクセス許可レベル](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteItem|
|適用可能な Outlook のモード| 作成|

##### <a name="example"></a>例

```JavaScript
function callback(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    showMessage("Attachment added");
  }
}

function addAttachment() {
  // The values in asyncContext can be accessed in the callback
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  var attachmentURL = "https://contoso.com/rtm/icon.png";
  Office.context.mailbox.item.addFileAttachmentAsync(attachmentURL, attachmentURL, options, callback);
}
```

####  <a name="additemattachmentasync(itemid,-attachmentname,-[options],-[callback])"></a>addItemAttachmentAsync(itemId, attachmentName, [options], [callback])

メッセージなどの Exchange アイテムを添付ファイルとして、メッセージまたは予定に追加します。

`addItemAttachmentAsync` メソッドは、指定した Exchange 識別子を持つアイテムを新規作成フォーム内のアイテムに添付します。コールバック メソッドを指定する場合、`asyncResult` という 1 つのパラメーターがあるメソッドが呼び出されます。このパラメーターには、添付ファイルの識別子、またはアイテムの添付中に発生したエラーを示すコードが含まれます。必要に応じて、`options` パラメーターを使用して、状態情報をコールバック メソッドに渡すことができます。

その後、[`removeAttachmentAsync`](Office.context.mailbox.item.md#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。

Office アドインを Outlook Web App で実行している場合、`addItemAttachmentAsync` メソッドはアイテムを、編集中のアイテム以外のアイテムに添付できますが、これはサポートされておらず、推奨されていません。

##### <a name="parameters:"></a>パラメーター:

|名前| 型| 属性| 説明|
|---|---|---|---|
|`itemId`| String||添付するアイテムの Exchange 識別子。最大の長さは 100 文字です。|
|`attachmentName`| String||添付するアイテムの件名。最大の長さは 255 文字です。|
|`options`| Object| &lt;optional&gt;|次のプロパティのうち 1 つ以上を含むオブジェクト リテラルです。<br/><br/>**プロパティ**<br/><table class="nested-table"><thead><tr><th>名前</th><th>型</th><th>属性</th><th>説明</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;optional&gt;</td><td>開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</td></tr></tbody></table>|
|`callback`| function| &lt;optional&gt;|メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](simple-types.md#asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。 <br/>成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。<br/>添付ファイルの追加が失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。<br/><table class="nested-table"><thead><tr><th>エラー コード</th><th>説明</th></tr></thead><tbody><tr><td><code>NumberOfAttachmentsExceeded</code></td><td>メッセージまたは予定の添付ファイルが多すぎます。</td></tr></tbody></table>|

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](../tutorial-api-requirement-sets.md)| 1.1|
|[最小限のアクセス許可レベル](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteItem|
|適用可能な Outlook のモード| 作成|

##### <a name="example"></a>例

次の例では、既存の Outlook アイテムが名前 `My Attachment` の添付ファイルとして追加されます。

```JavaScript
function callback(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    showMessage("Attachment added");
  }
}

function addAttachment() {
  // EWS ID of item to attach
  // (Shortened for readability)
  var itemId = "AAMkADI1...AAA=";

  // The values in asyncContext can be accessed in the callback
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  Office.context.mailbox.item.addItemAttachmentAsync(itemId, "My Attachment", options, callback);
}
```

#### <a name="displayreplyallform(formdata)"></a>displayReplyAllForm(formData)

選択したメッセージの送信者とすべての受信者、または選択した予定の開催者とすべての参加者を示した回答フォームが表示されます。

Outlook Web App では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。

文字列パラメーターのいずれかが制限値を超えると、`displayReplyAllForm` は例外をスローします。

> **注:**`displayReplyAllForm` に対する呼び出しに添付ファイルを含める機能は、要件セット 1.1 ではサポートされていません。添付ファイルのサポートは、要件セット 1.2 以降で `displayReplyAllForm` に追加されました。

##### <a name="parameters:"></a>パラメーター:

|名前| 型| 説明|
|---|---|---|
|`formData`| String &#124; Object|回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。<br/>**または**<br/>本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。<br/><br/>**プロパティ**<br/><table class="nested-table"><thead><tr><th>名前</th><th>型</th><th>属性</th><th>説明</th></tr></thead><tbody><tr><td><code>htmlBody</code></td><td>String</td><td>&lt;省略可能&gt;</td><td>回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</td></tr><tr><td><code>callback</code></td><td>function</td><td>&lt;省略可能&gt;</td><td>メソッドが完了すると、<code>callback</code> パラメーターに渡された関数が、<a href="simple-types.md#asyncresult"><code>AsyncResult</code></a> オブジェクトである 1 つのパラメーター <code>asyncResult</code> で呼び出されます。詳細については、「<a href="tutorial-asynchronous.html">asynchronous メソッドの使用</a>」を参照してください。</td></tr></tbody></table>|

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](../tutorial-api-requirement-sets.md)| 1.0|
|[最小限のアクセス許可レベル](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用可能な Outlook のモード| 読み取り|

##### <a name="examples"></a>例

次のコードは `displayReplyAllForm` 関数に文字列を渡します。

```JavaScript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

空の本文を返信します。

```JavaScript
Office.context.mailbox.item.displayReplyAllForm({});
```

本文だけを返信します。

```JavaScript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

本文とコールバックを返信します。

```JavaScript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="displayreplyform(formdata)"></a>displayReplyForm(formData)

選択したメッセージの送信者のみ、または選択した予定の開催者のみを示した回答フォームが表示されます。

Outlook Web App では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。

文字列パラメーターのいずれかが制限値を超えると、`displayReplyForm` は例外をスローします。

> **注:**`displayReplyForm` に対する呼び出しに添付ファイルを含める機能は、要件セット 1.1 ではサポートされていません。添付ファイルのサポートは、要件セット 1.2 以降で `displayReplyForm` に追加されました。

##### <a name="parameters:"></a>パラメーター:

|名前| 型| 説明|
|---|---|---|
|`formData`| String &#124; Object|回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。<br/>**または**<br/>本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。<br/><br/>**プロパティ**<br/><table class="nested-table"><thead><tr><th>名前</th><th>型</th><th>属性</th><th>説明</th></tr></thead><tbody><tr><td><code>htmlBody</code></td><td>String</td><td>&lt;省略可能&gt;</td><td>回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</td></tr><tr><td><code>callback</code></td><td>function</td><td>&lt;省略可能&gt;</td><td>メソッドが完了すると、<code>callback</code> パラメーターに渡された関数が、<a href="simple-types.md#asyncresult"><code>AsyncResult</code></a> オブジェクトである 1 つのパラメーター <code>asyncResult</code> で呼び出されます。詳細については、「<a href="tutorial-asynchronous.html">asynchronous メソッドの使用</a>」を参照してください。</td></tr></tbody></table>|

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](../tutorial-api-requirement-sets.md)| 1.0|
|[最小限のアクセス許可レベル](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用可能な Outlook のモード| 読み取り|

##### <a name="examples"></a>例

次のコードは `displayReplyForm` 関数に文字列を渡します。

```JavaScript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

空の本文を返信します。

```JavaScript
Office.context.mailbox.item.displayReplyForm({});
```

本文だけを返信します。

```JavaScript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

本文とコールバックを返信します。

```JavaScript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="getentities()-→-{[entities](simple-types.md#entities)}"></a>getEntities() → {[Entities](simple-types.md#entities)}

選択したアイテムにあるエンティティを取得します。

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](../tutorial-api-requirement-sets.md)| 1.0|
|[最小限のアクセス許可レベル](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用可能な Outlook のモード| 読み取り|

##### <a name="returns:"></a>戻り値:

型:
[Entities](simple-types.md#entities)

##### <a name="example"></a>例

次の例は、現在のアイテム上の連絡先のエンティティにアクセスします。

```JavaScript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytype(entitytype)-→-(nullable)-{array.<(string|[contact](simple-types.md#contact)|[meetingsuggestion](simple-types.md#meetingsuggestion)|[phonenumber](simple-types.md#phonenumber)|[tasksuggestion](simple-types.md#tasksuggestion))>}"></a>getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](simple-types.md#contact)|[MeetingSuggestion](simple-types.md#meetingsuggestion)|[PhoneNumber](simple-types.md#phonenumber)|[TaskSuggestion](simple-types.md#tasksuggestion))>}

選択したアイテム内で検出された指定のエンティティ型のすべてのエンティティを含む配列を取得します。

##### <a name="parameters:"></a>パラメーター:

|名前| 型| 説明|
|---|---|---|
|`entityType`| [Office.MailboxEnums.EntityType](Office.MailboxEnums.md#.entitytype-string)|EntityType 列挙値の 1 つ。|

##### <a name="requirements"></a>Requirements

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](../tutorial-api-requirement-sets.md)| 1.0|
|[最小限のアクセス許可レベル](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| 制限あり|
|適用可能な Outlook のモード| 読み取り|

##### <a name="returns:"></a>戻り値:

`entityType` に渡された値が `EntityType` 列挙型の有効なメンバーでない場合、メソッドは null を返します。指定した型のエンティティがアイテムに存在しない場合、メソッドは空の配列を返します。それ以外の場合は、返される配列内のオブジェクトの型は、`entityType` パラメーター内の要求されたエンティティの型によって異なります。

このメソッドを使用する最小限のアクセス許可レベルは **Restricted** ですが、一部のエンティティ型には、次の表で指定されているように、アクセスに **ReadItem** が必要です。

| `entityType` の値 | 返される配列内のオブジェクトの型 | 必要なアクセス許可のレベル |
| --- | --- | --- |
| `Address` | String | **制限あり** |
| `Contact` | 連絡先 | **ReadItem** |
| `EmailAddress` | String | **ReadItem** |
| `MeetingSuggestion` | MeetingSuggestion | **ReadItem** |
| `PhoneNumber` | PhoneNumber | **制限あり** |
| `TaskSuggestion` | TaskSuggestion | **ReadItem** |
| `URL` | String | **制限あり** |

型:Array.<(String|[Contact](simple-types.md#contact)|[MeetingSuggestion](simple-types.md#meetingsuggestion)|[PhoneNumber](simple-types.md#phonenumber)|[TaskSuggestion](simple-types.md#tasksuggestion))></dd>


##### <a name="example"></a>例

次の例は、現在のアイテムの件名または本文にある郵便アドレスを表す文字列の配列にアクセスする方法を示します。

```JavaScript
// The initialize function is required for all apps.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    var item = Office.context.mailbox.item;
    // Get an array of strings that represent postal addresses in the current item.
    var addresses = item.getEntitiesByType(Office.MailboxEnums.EntityType.Address);
    // Continue processing the array of addresses.
  });
}
```

#### <a name="getfilteredentitiesbyname(name)-→-(nullable)-{array.<(string|[contact](simple-types.md#contact)|[meetingsuggestion](simple-types.md#meetingsuggestion)|[phonenumber](simple-types.md#phonenumber)|[tasksuggestion](simple-types.md#tasksuggestion))>}"></a>getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](simple-types.md#contact)|[MeetingSuggestion](simple-types.md#meetingsuggestion)|[PhoneNumber](simple-types.md#phonenumber)|[TaskSuggestion](simple-types.md#tasksuggestion))>}

マニフェスト XML ファイルで定義された名前付きフィルターを通過する、選択したアイテム内の既知のエンティティを返します。

`getFilteredEntitiesByName` メソッドは、マニフェスト XML ファイル内の、指定された `FilterName` 要素値を持つ [ItemHasKnownEntity](https://msdn.microsoft.com/en-us/library/office/fp161166.aspx) ルール要素で定義された正規表現に一致するエンティティを返します。

##### <a name="parameters:"></a>パラメーター:

|名前| 型| 説明|
|---|---|---|
|`name`| String|一致するフィルターを定義する `ItemHasKnownEntity` ルール要素の名前。|

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](../tutorial-api-requirement-sets.md)| 1.0|
|[最小限のアクセス許可レベル](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用可能な Outlook のモード| 読み取り|

##### <a name="returns:"></a>戻り値:

`FilterName` 要素の値が `name` パラメーターと一致するマニフェスト内に `ItemHasKnownEntity` 要素がない場合、メソッドは `null` を返します。`name` パラメーターがマニフェスト内の `ItemHasKnownEntity` 要素と一致せず、現在のアイテム内に一致するエンティティがない場合は、メソッドは空の配列を返します。


型:Array.<(String|[Contact](simple-types.md#contact)|[MeetingSuggestion](simple-types.md#meetingsuggestion)|[PhoneNumber](simple-types.md#phonenumber)|[TaskSuggestion](simple-types.md#tasksuggestion))>


#### <a name="getregexmatches()-→-{object}"></a>getRegExMatches() → {Object}

選択したアイテム内の、マニフェスト XML ファイルで定義された正規表現に一致する文字列の値を返します。

`getRegExMatches` メソッドは、マニフェスト XML ファイル内の、各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定されたアイテムのプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。

たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。

```JavaScript
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

`getRegExMatches` から返されるオブジェクトに `fruits` および `veggies` という 2 つのプロパティがあります。

```JavaScript
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。この場合、代わりに [`Body.getAsync`](Body.md#getAsync) メソッドを使用して本文全体を取得します。

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](../tutorial-api-requirement-sets.md)| 1.0|
|[最小限のアクセス許可レベル](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用可能な Outlook のモード| 読み取り|

##### <a name="returns:"></a>戻り値:

マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクト。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。

<dl class="param-type">

<dt>型</dt>

<dd>Object</dd>

</dl>

##### <a name="example"></a>例

次の例は、マニフェストで指定された正規表現 <rule> の要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</rule>

```JavaScript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbyname(name)-→-(nullable)-{array.<string>}"></a>getRegExMatchesByName(name) → (nullable) {Array.<String>}

選択したアイテム内の、マニフェスト XML ファイルで定義された、指定された正規表現に一致する文字列の値を返します。

`getRegExMatchesByName` メソッドは、`ItemHasRegularExpressionMatch` ルール要素で定義された正規表現に一致する文字列を返します。このルール要素は、指定された `RegExName` 要素値を持つマニフェスト XML ファイル内にあります。

アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。

##### <a name="parameters:"></a>パラメーター:

|名前| 型| 説明|
|---|---|---|
|`name`| String|一致するフィルターを定義する `ItemHasRegularExpressionMatch` ルール要素の名前。|

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](../tutorial-api-requirement-sets.md)| 1.0|
|[最小限のアクセス許可レベル](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用可能な Outlook のモード| 読み取り|

##### <a name="returns:"></a>戻り値:

マニフェスト XML ファイルで定義された正規表現に一致する文字列が格納された配列。

<dl class="param-type">

<dt>型</dt>

<dd>Array。<String></dd>

</dl>

##### <a name="example"></a>例

```JavaScript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasync(coerciontype,-[options],-callback)-→-{string}"></a>getSelectedDataAsync(coercionType, [options], callback) → {String}

メッセージの件名または本文から非同期的に選択したデータを返します。

選択したデータがなく、カーソルが本文または件名にある場合、選択したデータに対して null が返されます。本文または件名以外のフィールドが選択されている場合、`InvalidSelection` エラーが返されます。

##### <a name="parameters:"></a>パラメーター:

|名前| 型| 属性| 説明|
|---|---|---|---|
|`coercionType`| [Office.CoercionType](Office.md#coerciontype-string)||データの形式を要求します。テキストの場合、メソッドは文字列としてプレーン テキストを返し、存在する HTML タグはすべて削除されます。HTMLの場合、メソッドは、プレーンテキストまたは HTML のいずれの場合も選択されたテキストを返します。|
|`options`| Object| &lt;optional&gt;|次のプロパティのうち 1 つ以上を含むオブジェクト リテラルです。<br/><br/>**プロパティ**<br/><table class="nested-table"><thead><tr><th>名前</th><th>型</th><th>属性</th><th>説明</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;optional&gt;</td><td>開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</td></tr></tbody></table>|
|`callback`| function||メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](simple-types.md#asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。

コールバック メソッドから選択したデータにアクセスするには、`asyncResult.value.data` を呼び出します。選択のソース プロパティにアクセスするには、`asyncResult.value.sourceProperty` を呼び出します。これは `body` または `subject` になります。|

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](../tutorial-api-requirement-sets.md)| 1.0|
|[最小限のアクセス許可レベル](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteItem|
|適用可能な Outlook のモード| 作成|

##### <a name="returns:"></a>戻り値:

選択されたデータ (`coercionType` で決定された形式の文字列)。

<dl class="param-type">

<dt>型</dt>

<dd>String</dd>

</dl>

##### <a name="example"></a>例

```JavaScript
// getting selected data
Office.initialize = function () {
    Office.context.mailbox.item.getSelectedDataAsync(Office.CoercionType.Text, {}, getCallback);
}

function getCallback(asyncResult) {
    var text = asyncResult.value.data;
    var prop = asyncResult.value.sourceProperty;

    Office.context.mailbox.item.setSelectedDataAsync('Setting ' + prop + ': ' + text, {}, setCallback);
}

function setCallback(asyncResult) {
    // check for errors
}
```

####  <a name="loadcustompropertiesasync(callback,-[usercontext])"></a>loadCustomPropertiesAsync(callback, [userContext])

選択されたアイテムのこのアドインのカスタム プロパティを非同期に読み込みます。

カスタム プロパティは、アプリケーションごと、アイテムごとのキーと値のペアとして格納されます。このメソッドは、コールバックで `CustomProperties` オブジェクトを返します。このオブジェクトは、現在のアイテムおよび現在のアドインに固有のカスタム プロパティにアクセスするためのメソッドを提供します。カスタム プロパティは、アイテム上では暗号化されません。そのため、セキュリティ保護記憶域として使用するべきではありません。

##### <a name="parameters:"></a>パラメーター:

|名前| 型| 属性| 説明|
|---|---|---|---|
|`callback`| function||メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](simple-types.md#asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。

カスタム プロパティは `asyncResult.value` プロパティの [`CustomProperties`](CustomProperties.md) オブジェクトとして指定されます。このオブジェクトは、アイテムのカスタム プロパティを取得、設定、および削除し、カスタム プロパティに対する変更をサーバーに設定し直すために使用できます。| 
|`userContext`| オブジェクト| &lt;省略可能&gt;|開発者は、コールバック関数でアクセスする任意のオブジェクトを指定できます。このオブジェクトには、コールバック関数の `asyncResult.asyncContext` プロパティによってアクセスすることができます。

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](../tutorial-api-requirement-sets.md)| 1.0|
|[最小限のアクセス許可レベル](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用可能な Outlook のモード| 作成または読み取り|

##### <a name="example"></a>例

次のコード例では、`loadCustomPropertiesAsync` メソッドを使用して、現在のアイテムに固有のカスタム プロパティを非同期的に読み込む方法を示します。また、`CustomProperties.saveAsync` メソッドを使用して、これらのプロパティをサーバーに保存する方法も紹介します。カスタム プロパティをロードした後、このコード サンプルでは `CustomProperties.get` メソッドを使用してカスタム プロパティ `myProp` を読み取り、`CustomProperties.set` メソッドでカスタム プロパティ `otherProp` を書き込み、最後に `saveAsync` メソッドを呼び出して、カスタム プロパティを保存します。

```JavaScript
// The initialize function is required for all add-ins.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
  // After the DOM is loaded, add-in-specific code can run.
  var item = Office.context.mailbox.item;
  item.loadCustomPropertiesAsync(customPropsCallback);
  });
}

function customPropsCallback(asyncResult) {
  var customProps = asyncResult.value;
  var myProp = customProps.get("myProp");

  customProps.set("otherProp", "value");
  customProps.saveAsync(saveCallback);
}

function saveCallback(asyncResult) {
}
```

####  <a name="removeattachmentasync(attachmentid,-[options],-[callback])"></a>removeAttachmentAsync(attachmentId, [options], [callback])

メッセージまたは予定から添付ファイルを削除します。

`removeAttachmentAsync` メソッドは、指定した識別子の添付ファイルをアイテムから削除します。ベスト プラクティスとして、同じメール アプリが同じセッションで添付ファイルを追加した場合にのみ、その添付ファイルの識別子を使用して添付ファイルを削除することをお勧めします。Outlook Web App とデバイス用 OWA では、添付ファイルの識別子は同じセッション内でのみ有効です。ユーザーがアプリを閉じるか、ユーザーがインライン フォームで新規作成を開始してインライン フォームが表示され、別ウィンドウで操作を継続すると、セッションは終了します。

##### <a name="parameters:"></a>パラメーター:

|名前| 型| 属性| 説明|
|---|---|---|---|
|`attachmentId`| String||削除する添付ファイルの識別子。文字列の最大の長さは 100 文字です。|
|`options`| Object| &lt;optional&gt;|次のプロパティのうち 1 つ以上を含むオブジェクト リテラルです。<br/><br/>**プロパティ**<br/><table class="nested-table"><thead><tr><th>名前</th><th>型</th><th>属性</th><th>説明</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;optional&gt;</td><td>開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</td></tr></tbody></table>|
|`callback`| function| &lt;optional&gt;|メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](simple-types.md#asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。 <br/>添付ファイルの削除に失敗すると、`asyncResult.error` プロパティにはエラー コードとエラーの理由が含まれます。<br/><table class="nested-table"><thead><tr><th>エラー コード</th><th>説明</th></tr></thead><tbody><tr><td><code>InvalidAttachmentId</code></td><td>添付ファイル識別子が存在しません。</td></tr></tbody></table>|

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](../tutorial-api-requirement-sets.md)| 1.1|
|[最小限のアクセス許可レベル](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteItem|
|適用可能な Outlook のモード| 作成|

##### <a name="example"></a>例

次のコードは、'0' の識別子を持つ添付ファイルを削除します。

```JavaScript
Office.context.mailbox.item.removeAttachmentAsync(
  '0',
  { asyncContext : null },
  function (asyncResult)
  {
    console.log(asyncResult.status);
  }
);
```
