 

# MailboxEnums

## [Office](Office.md).MailboxEnums

##### 要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](../tutorial-api-requirement-sets.md)| 1.0|
|適用可能な Outlook のモード| 作成|

### メンバー

#### AttachmentType :String

添付ファイルの種類を指定します。新規作成モードのみ。

AttachmentType

##### 型:

*   String

##### プロパティ:

|名前| 型| 説明|
|---|---|---|
|`File`| String|この添付ファイルはファイルです。|
|`Item`| String|この添付ファイルは Exchange のアイテムです。|

##### 要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](../tutorial-api-requirement-sets.md)| 1.0|
|適用可能な Outlook のモード| 作成|
#### EntityType :String

エンティティの種類を指定します。新規作成モードのみ。

EntityType

##### 型:

*   String

##### プロパティ:

|名前| 型| 説明|
|---|---|---|
|`Address`| String|エンティティが郵送先住所であることを指定します。|
|`Contact`| String|エンティティが連絡先であることを指定します。|
|`EmailAddress`| String|エンティティが SMTP 電子メール アドレスであることを指定します。|
|`MeetingSuggestion`| String|エンティティが提案された会議であることを指定します。|
|`PhoneNumber`| String|エンティティが米国の電話番号であることを指定します。|
|`TaskSuggestion`| String|エンティティがタスクのヒントであることを指定します。|
|`URL`| String|エンティティがインターネット URL であることを指定します。|

##### 要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](../tutorial-api-requirement-sets.md)| 1.0|
|適用可能な Outlook のモード| 作成|
#### ItemType :String

アイテムの種類を指定します。新規作成モードのみ。

ItemType

##### 型:

*   String

##### プロパティ:

|名前| 型| 説明|
|---|---|---|
|`Message`| String|電子メール、会議出席依頼、会議の返信、または会議の取り消し。|
|`Appoinment`| String|予定アイテム。|

##### 要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](../tutorial-api-requirement-sets.md)| 1.0|
|適用可能な Outlook のモード| 作成|
#### RecipientType :String

予定の受信者の種類を指定します。作成モードのみ。

RecipientType

##### 型:

*   String

##### プロパティ:

|名前| 型| 説明|
|---|---|---|
|`Other`| String|受信者は、他の種類の受信者ではありません。|
|`DistributionList`| String|受信者は、電子メール アドレスの一覧を含む配布リストです。|
|`User`| String|受信者は、Exchange サーバー上の SMTP 電子メール アドレスです。|
|`ExternalUser`| String|受信者は、Exchange サーバー上にはない SMTP 電子メール アドレスです。|

##### 要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](../tutorial-api-requirement-sets.md)| 1.1|
|適用可能な Outlook のモード| 作成|
#### ResponseType :String

会議出席依頼への応答の種類を指定します。新規作成モードのみ。

ResponseType

##### 型:

*   String

##### プロパティ:

|名前| 型| 説明|
|---|---|---|
|`None`| String|出席者からの応答がありません。|
|`Organizer`| String|出席者は会議開催者です。|
|`Tentative`| String|出席者が会議出席依頼を仮承諾しました。|
|`Accepted`| String|出席者が会議出席依頼を承諾しました。|
|`Declined`| String|出席者が会議出席依頼を拒否しました。|

##### 要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](../tutorial-api-requirement-sets.md)| 1.0|
|適用可能な Outlook のモード| 作成|
