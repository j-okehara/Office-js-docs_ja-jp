 

# MailboxEnums

## [Office](Office.md).MailboxEnums

##### 要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](./tutorial-api-requirement-sets.md)| 1.0|
|適用可能な Outlook のモード| 作成または読み取り|

### メンバー

#### AttachmentType :String

添付ファイルの種類を指定します。

AttachmentType

##### 型:

*   String

##### プロパティ:

|名前| 型| 値 | 説明|
|---|---|---|---|
|`File`| String|`file`|この添付ファイルはファイルです。|
|`Item`| String|`item`|この添付ファイルは Exchange のアイテムです。|

##### 要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](./tutorial-api-requirement-sets.md)| 1.0|
|適用可能な Outlook のモード| 作成または読み取り|
#### EntityType :String

エンティティの種類を指定します。

EntityType

##### 型:

*   String

##### プロパティ:

|名前| 型| 値 | 説明|
|---|---|---|---|
|`Address`| String|`address`|エンティティが郵送先住所であることを指定します。|
|`Contact`| String|`contact`|エンティティが連絡先であることを指定します。|
|`EmailAddress`| String|`emailAddress`|エンティティが SMTP 電子メール アドレスであることを指定します。|
|`MeetingSuggestion`| String|`meetingSuggestion`|エンティティが提案された会議であることを指定します。|
|`PhoneNumber`| String|`phoneNumber`|エンティティが米国の電話番号であることを指定します。|
|`TaskSuggestion`| String|`taskSuggestion`|エンティティがタスクのヒントであることを指定します。|
|`URL`| String|`url`|エンティティがインターネット URL であることを指定します。|

##### 要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](./tutorial-api-requirement-sets.md)| 1.0|
|適用可能な Outlook のモード| 作成または読み取り|
#### ItemNotificationMessageType :String

予定またはメッセージの通知メッセージの種類を指定します。

ItemNotificationMessageType

##### 型:

*   String

##### プロパティ:

|名前| 型| 値 | 説明|
|---|---|---|---|
|`ProgressIndicator`| String|`progressIndicator`|notificationMessage は進行状況インジケーターです。|
|`InformationalMessage`| String|`informationalMessage`|notificationMessage は情報メッセージです。|
|`ErrorMessage`| String|`errorMessage`|notificationMessage はエラー メッセージです。|

##### 要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](./tutorial-api-requirement-sets.md)| 1.3|
|適用可能な Outlook のモード| 作成または読み取り|
#### ItemType :String

アイテムの種類を指定します。

ItemType

##### 型:

*   String

##### プロパティ:

|名前| 型| 値 | 説明|
|---|---|---|---|
|`Message`| String|`message`|電子メール、会議出席依頼、会議の返信、または会議の取り消し。|
|`Appointment`| String|`appointment`|予定アイテム。|

##### 要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](./tutorial-api-requirement-sets.md)| 1.0|
|適用可能な Outlook のモード| 作成または読み取り|
#### RecipientType :String

予定の受信者の種類を指定します。

RecipientType

##### 型:

*   String

##### プロパティ:

|名前| 型| 値 | 説明|
|---|---|---|---|
|`Other`| String|`other`|受信者は、他の種類の受信者ではありません。|
|`DistributionList`| String|`distributionList`|受信者は、電子メール アドレスの一覧を含む配布リストです。|
|`User`| String|`user`|受信者は、Exchange サーバー上の SMTP 電子メール アドレスです。|
|`ExternalUser`| String|`externalUser`|受信者は、Exchange サーバー上にはない SMTP 電子メール アドレスです。|

##### 要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](./tutorial-api-requirement-sets.md)| 1.1|
|適用可能な Outlook のモード| 作成または読み取り|
#### ResponseType :String

会議出席依頼への応答の種類を指定します。

ResponseType

##### 型:

*   String

##### プロパティ:

|名前| 型| 値 | 説明|
|---|---|---|---|
|`None`| String|`none`|出席者からの応答がありません。|
|`Organizer`| String|`organizer`|出席者は会議開催者です。|
|`Tentative`| String|`tentative`|出席者が会議出席依頼を仮承諾しました。|
|`Accepted`| String|`accepted`|出席者が会議出席依頼を承諾しました。|
|`Declined`| String|`declined`|出席者が会議出席依頼を拒否しました。|

##### 要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](./tutorial-api-requirement-sets.md)| 1.0|
|適用可能な Outlook のモード| 作成または読み取り|

#### RestVersion :String

REST 形式のアイテム ID に対応する REST API のバージョンを指定します。 

RestVersion

##### 型:

*   String

##### プロパティ:

|名前| 型| 値 | 説明|
|---|---|---|---|
|`v1_0`| String|`v1.0`|バージョン 1.0。|
|`v2_0`| String|`v2.0`|バージョン 2.0。|
|`Beta`| String|`beta`|ベータ版。|

##### 要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](./tutorial-api-requirement-sets.md)| 1.3|
|適用可能な Outlook のモード| 作成または読み取り|
