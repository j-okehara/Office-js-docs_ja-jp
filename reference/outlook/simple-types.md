

# 単純型

####  AsyncResult

要求が失敗した場合の状態やエラー情報など、非同期要求の結果をカプセル化するオブジェクト。

##### プロパティ:

|名前| 型| 説明|
|---|---|---|
|`asyncContext`| Object|呼び出されたメソッドの省略可能な `asyncContext` パラメーターに渡されたオブジェクトを、渡されたときと同じ状態で取得します。|
|`error`| エラー|エラーの説明を提供する Error オブジェクトを取得します (エラーが発生した場合)。|
|`status`| [Office.AsyncResultStatus](Office.md#.AsyncResultStatus-string)|非同期操作の状態を取得します。|
|`value`| Object|この非同期操作のペイロードまたはコンテンツを取得します (ある場合)。|

##### 要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](./tutorial-api-requirement-sets.md)| 1.0|
|適用可能な Outlook のモード| 作成または読み取り|
#### AttachmentDetails

サーバーからのアイテムの添付ファイルを表します。閲覧モードのみ。

`AttachmentDetail` オブジェクトの配列が、`attachments` または `Appointment` オブジェクトの `Message` のプロパティとして返されます。

##### プロパティ:

|名前| 型| 説明|
|---|---|---|
|`attachmentType`| [Office.MailboxEnums.AttachmentType](Office.MailboxEnums.md#attachmenttype-string)|添付ファイルの種類を示す値を取得します。|
|`contentType`| String|添付ファイルの MIME コンテンツ タイプを取得します。|
|`id`| String|添付ファイルの Exchange 添付ファイル ID を取得します。|
|`isInline`| ブール型 (Boolean)|添付ファイルをアイテムの本文に表示するかどうかを示す値を取得します。|
|`name`| String|添付ファイルの名前を取得します。|
|`size`| Number|添付ファイルのサイズをバイト単位で取得します。|

##### 要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](./tutorial-api-requirement-sets.md)| 1.0|
|[最小のアクセス許可レベル](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用可能な Outlook のモード| 読み取り|
#### Contact

サーバーに格納された連絡先を表します。閲覧モードのみ。

アクティブなアイテムの `contacts` メソッドまたは `Entities` メソッドによって返される [`getEntities`](simple-types.md#entities) オブジェクトの `getEntitiesByType` プロパティに、電子メール メッセージまたは予定に関連付けられた連絡先のリストが返されます。

##### プロパティ:

|名前| 型| 属性| 説明|
|---|---|---|---|
|`addresses`| Array.&lt;String&gt;| &lt;nullable&gt;|連絡先に関連付けられているメールアドレスと住所を含む文字列の配列。|
|`businessName`| String| &lt;nullable&gt;|連絡先に関連付けられた取引先の名前が含まれている文字列。|
|`emailAddresses`| Array.&lt;String&gt;| &lt;nullable&gt;|連絡先に関連付けられている SMTP メールアドレスを含む文字列の配列。|
|`personName`| String| &lt;nullable&gt;|連絡先に関連付けられた人物の名前が含まれている文字列。|
|`phoneNumbers`| Array.&lt;[PhoneNumber](simple-types.md#phonenumber)&gt;| &lt;nullable&gt;|連絡先に関連付けられた各電話番号の `PhoneNumber` オブジェクトが含まれている配列。|
|`urls`| Array.&lt;String&gt;| &lt;nullable&gt;|連絡先に関連付けられているインターネットの URL を含む文字列の配列。|

##### 要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](./tutorial-api-requirement-sets.md)| 1.0|
|[最小のアクセス許可レベル](../../docs/outlook/understanding-outlook-add-in-permissions.md)| Restricted|
|適用可能な Outlook のモード| 読み取り|
####  EmailAddressDetails

電子メール メッセージまたは予定の送信者または指定受信者の電子メール プロパティを提供します。

##### 型:

*   Object

##### プロパティ:

|名前| 型| 説明|
|---|---|---|
|`appointmentResponse`| [Office.MailboxEnums.ResponseType](Office.MailboxEnums.md#responsetype-string)|予定に対して出席者が戻した応答を取得します。このプロパティは、[`optionalAttendees`](Office.context.mailbox.item.md#optionalattendees-arrayemailaddressdetails) プロパティまたは [`requiredAttendees`](Office.context.mailbox.item.md#requiredattendees-arrayemailaddressdetailsrecipients) プロパティで表わされる予定の出席者にのみ適用されます。このプロパティは、他のシナリオでは `undefined` を返します。|
|`displayName`| String|電子メール アドレスに関連付けられた表示名を取得します。|
|`emailAddress`| String|SMTP 電子メール アドレスを取得します。|
|`recipientType`| [Office.MailboxEnums.RecipientType](Office.MailboxEnums.md#recipienttype-string)|受信者の電子メール アドレスの種類を取得します。|

##### 要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](./tutorial-api-requirement-sets.md)| 1.0|
|[最小のアクセス許可レベル](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用可能な Outlook のモード| 作成または読み取り|
#### EmailUser

Exchange Server 上の電子メール アカウントを表します。

##### プロパティ:

|名前| 型| 説明|
|---|---|---|
|`displayName`| String|電子メール アドレスに関連付けられた表示名を取得します。|
|`emailAddress`| String|SMTP 電子メール アドレスを取得します。|

##### 要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](./tutorial-api-requirement-sets.md)| 1.0|
|[最小のアクセス許可レベル](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用可能な Outlook のモード| 読み取り|

#### エンティティ

電子メール メッセージまたは予定に含まれているエンティティのコレクションを表します。閲覧モードのみ。

`Entities` オブジェクトは、サーバーによって見つけられた 1 つ以上のエンティティがアイテム (電子メール メッセージまたは予定) に含まれている場合に、`getEntities` メソッドと `getEntitiesByType` メソッドによって返されるエンティティ配列のコンテナーです。これらのエンティティをコード内で使用することにより、アイテム内のアドレスへのマップなどの追加のコンテキスト情報をビューアーに提供したり、アイテム内の電話番号に対してダイヤラーを開いたりできます。

プロパティで指定された型のエンティティがアイテム内にない場合、そのエンティティに関連付けられているプロパティは `null` になります。たとえば、メッセージに番地と電話番号が含まれている場合、`addresses` プロパティと `phoneNumbers` プロパティには情報が含まれ、それ以外のプロパティは `null` になります。

住所として認識されるには、文字列に米国の住所 (少なくとも番地、通り名、都市名、州名、郵便番号の要素を含む) が含まれている必要があります。

電話番号として認識されるためには、北アメリカの電話番号の形式を文字列に含める必要があります。

エンティティの認識には、大量のデータの機械学習に基づいた自然言語認識を利用しています。エンティティの認識は決定論的ではなく、結果がアイテムの特定のコンテキストに左右されることがあります。

`getEntitiesByType` メソッドによってプロパティ配列が返された場合、指定のエンティティのプロパティだけにデータが含まれ、それ以外のプロパティはすべて `null` になります。

##### プロパティ:

|名前| 型| 属性| 説明|
|---|---|---|---|
|`addresses`| Array.&lt;String&gt;| &lt;nullable&gt;|電子メール メッセージまたは予定に含まれている物理的な住所 (番地または郵送先住所) を取得します。|
|`contacts`| Array.&lt;[Contact](simple-types.md#contact)&gt;| &lt;nullable&gt;|電子メール アドレスまたは予定に含まれている連絡先を取得します。|
|`emailAddresses`| Array.&lt;String&gt;| &lt;nullable&gt;|電子メール メッセージまたは予定に含まれている電子メール アドレスを取得します。|
|`meetingSuggestions`| Array.&lt;[MeetingSuggestion](simple-types.md#meetingsuggestion)&gt;| &lt;nullable&gt;|電子メール メッセージ含まれている会議の提案を取得します。|
|`phoneNumbers`| Array.&lt;[PhoneNumber](simple-types.md#phonenumber)&gt;| &lt;nullable&gt;|電子メール メッセージや予定に含まれている電話番号を取得します。|
|`taskSuggestions`| Array.&lt;[TaskSuggestion](simple-types.md#tasksuggestion)&gt;| &lt;nullable&gt;|電子メール メッセージまたは予定に含まれている、タスクの提案を取得します。|
|`urls`| Array.&lt;String&gt;| &lt;nullable&gt;|電子メール メッセージまたは予定に含まれているインターネット URL を取得します。|

##### 要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](./tutorial-api-requirement-sets.md)| 1.0|
|[最小のアクセス許可レベル](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用可能な Outlook のモード| 読み取り|
#### LocalClientTime

ローカルのクライアントのタイム ゾーンの日付と時刻を表します。閲覧モードのみ。

##### プロパティ:

|名前| 型| 説明|
|---|---|---|
|`month`| Number|月を表す整数値。0 は 1 月を表し、11 は 12 月を表します。|
|`date`| Number|日付を表す整数値。|
|`year`| Number|年を表す整数値。|
|`hours`| Number|24 時間制の時間を表す整数値。|
|`minutes`| Number|分を表す整数値。|
|`seconds`| Number|秒を表す整数値。|
|`milliseconds`| Number|ミリ秒を表す整数値。|
|`timezoneOffset`| Number|ローカル タイム ゾーンと UTC との間の分数の差を表す整数値。|

##### 要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](./tutorial-api-requirement-sets.md)| 1.0|
|[最小のアクセス許可レベル](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用可能な Outlook のモード| 読み取り|
#### MeetingSuggestion

アイテムに含まれている提案された会議を表します。閲覧モードのみ。

アクティブなアイテムに対して [`meetingSuggestions`](simple-types.md#entities) メソッドまたは [`Entities`](Office.context.mailbox.item.md#getentities--entities) メソッドが呼び出されたときに返される、[`getEntities`](Office.context.mailbox.item.md#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) オブジェクトの `getEntitiesByType` プロパティに、電子メール メッセージに含まれている提案された会議のリストが返されます。

`start` および `end` の値は、会議の候補の開始日時と終了日時が含まれている Date オブジェクトの文字列表現です。値は、現在のユーザーに対して指定された既定のタイム ゾーンです。

##### プロパティ:

|名前| 型| 説明|
|---|---|---|
|`attendees`| Array.&lt;[EmailUser](simple-types.md#emailuser)&gt;|提案された会議の出席者を取得します。|
|`end`| String|提案された会議が終了する日付と時刻を取得します。|
|`location`| String|提案された会議の場所を取得します。|
|`meetingString`| String|会議の提案として識別された文字列を取得します。|
|`start`| String|提案された会議が開始する日時を取得します。|
|`subject`| String|提案された会議の件名を取得します。|

##### 要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](./tutorial-api-requirement-sets.md)| 1.0|
|[最小のアクセス許可レベル](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用可能な Outlook のモード| 読み取り|
####  NotificationMessageDetails

`NotificationMessageDetails` オブジェクトの配列は、[`NotificationMessages.getAllAsync`](NotificationMessages.md#getallasyncoptions-callback) メソッドによって返されます。

##### 型:

*   Object

##### プロパティ:

|名前| 型| 説明|
|---|---|---|
|`key`| String|通知メッセージの識別子。|
|`type`| [Office.MailboxEnums.ItemNotificationMessageType](Office.MailboxEnums.md#.ItemNotificationMessageType)|通知メッセージの型。|
|`icon`| String|メッセージに使用するアイコンのリソース識別子。`type` が `InformationalMessage` の場合にのみ適用されます。|
|`message`| String|これは、メッセージのテキストです。最大長は 150 文字です。|
|`persistent`| ブール型 (Boolean)|`true` の場合、メッセージはこのアドインによって削除されるか、ユーザーが非表示にするまで残されます。`false` の場合、ユーザーが別のアイテムに移動すると削除されます。`type` が `InformationalMessage` の場合にのみ適用されます。|

##### 要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](./tutorial-api-requirement-sets.md)| 1.3|
|[最小のアクセス許可レベル](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用可能な Outlook のモード| 作成または読み取り|
#### PhoneNumber

アイテム内の識別される電話番号を表します。閲覧モードのみ。

電子メール メッセージ内の電話番号が含まれる `PhoneNumber` オブジェクトの配列は、選択したアイテムの [`phoneNumbers`](simple-types.md#entities) メソッドを呼び出したときに返される、[`Entities`](Office.context.mailbox.item.md#getEntities) オブジェクトの `getEntities` プロパティに返されます。

##### 型:

*   Object

##### プロパティ:

|名前| 型| 説明|
|---|---|---|
|`originalPhoneString`| String|アイテム内の電話番号として識別されたテキストを取得します。|
|`phoneString`| String|電話番号が含まれている文字列を取得します。この文字列は、電話番号の数字のみを含みます。元のアイテムにかっこやハイフンなどの文字が含まれている場合でも、この文字列にはそれらの文字は含まれません。|
|`type`| String|電話番号の種類 (`Home`、`Work`、`Mobile`、`Unspecified`) を識別する文字列を取得します。|

##### 要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](./tutorial-api-requirement-sets.md)| 1.0|
|[最小のアクセス許可レベル](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用可能な Outlook のモード| 読み取り|
#### TaskSuggestion

アイテム内の識別される推奨タスクを表します。閲覧モードのみ。

アクティブなアイテムに対して [`taskSuggestions`](simple-types.md#entities) メソッドまたは [`Entities`](Office.context.mailbox.item.md#getEntities) メソッドが呼び出されたときに返される、[`Entities`][`getEntities`](Office.context.mailbox.item.md#getEntitiesByType) オブジェクトの `getEntitiesByType` プロパティに、電子メール メッセージに含まれている提案されたタスクのリストが返されます。

##### プロパティ:

|名前| 型| 説明|
|---|---|---|
|`assignees`| Array.&lt;[EmailUser](simple-types.md#emailuser)&gt;|推奨タスクに割り当てる必要のあるユーザーを取得します。|
|`taskString`| String|タスクの提案として識別されたアイテムのテキストを取得します。|

##### 要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](./tutorial-api-requirement-sets.md)| 1.0|
|[最小のアクセス許可レベル](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用可能な Outlook のモード| 読み取り|
