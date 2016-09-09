

# userProfile

## [Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile

##### 要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](./tutorial-api-requirement-sets.md)| 1.0|
|[最小のアクセス許可レベル](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用可能な Outlook のモード| 作成または読み取り|

### メンバー

####  displayName :String

ユーザーの表示名を取得します。

##### 型:

*   String

##### 要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](./tutorial-api-requirement-sets.md)| 1.0|
|[最小のアクセス許可レベル](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用可能な Outlook のモード| 作成または読み取り|

##### 例

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  emailAddress :String

ユーザーの SMTP 電子メール アドレスを取得します。

##### 型:

*   String

##### 要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](./tutorial-api-requirement-sets.md)| 1.0|
|[最小のアクセス許可レベル](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用可能な Outlook のモード| 作成または読み取り|

##### 例

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  timeZone :String

ユーザーの既定のタイム ゾーンを取得します。

##### 型:

*   String

##### 要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](./tutorial-api-requirement-sets.md)| 1.0|
|[最小のアクセス許可レベル](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用可能な Outlook のモード| 作成または読み取り|

##### 例

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```