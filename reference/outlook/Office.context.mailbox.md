

# mailbox

## [Office](Office.md)[.context](Office.context.md). mailbox

Microsoft Outlook と Microsoft Outlook on the web の Outlook アドイン オブジェクト モデルへのアクセスを提供します。

##### 要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](./tutorial-api-requirement-sets.md)| 1.0|
|[最小のアクセス許可レベル](../../docs/outlook/understanding-outlook-add-in-permissions.md)| Restricted|
|適用可能な Outlook のモード| 作成または読み取り|

### 名前空間

[diagnostics](Office.context.mailbox.diagnostics.md):Outlook アドインに診断情報を提供します。

[item](Office.context.mailbox.item.md):Outlook アドインのメッセージや予定にアクセスするためのメソッドとプロパティを提供します。

[userProfile](Office.context.mailbox.userProfile.md):Outlook アドインのユーザーに関する情報を提供します。</dd>

### メンバー

#### ewsUrl :String

この電子メール アカウントの Exchange Web サービス (EWS) エンドポイントの URL を取得します。 閲覧モードのみ。

`ewsUrl` 値は、リモート サービスで、ユーザーのメールボックスに EWS 呼び出しを行うために使うことができます。 たとえば、[選択したアイテムから添付ファイルを取得する](https://msdn.microsoft.com/EN-US/library/office/dn148008.aspx)ためにリモート サービスを作成できます。

アプリが閲覧モードで `ewsUrl` メンバーを呼び出すには、アプリのマニフェスト内に **ReadItem** アクセス許可が指定されている必要があります。

新規作成モードでは、[`saveAsync`](Office.context.mailbox.item#saveAsync) メソッドを呼び出してから、`ewsUrl` メンバーを使用する必要があります。 アプリには、`saveAsync` メソッドを呼び出す **ReadWriteItem** アクセス許可が必要です。

##### 型:

*   String

##### 要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](./tutorial-api-requirement-sets.md)| 1.0|
|[最小のアクセス許可レベル](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用可能な Outlook のモード| 作成または読み取り|

### メソッド

####  convertToEwsId(itemId, restVersion) → {String}

REST 形式のアイテム ID を EWS 形式に変換します。

REST API ([Outlook Mail API](https://msdn.microsoft.com/office/office365/APi/mail-rest-operations) や [Microsoft Graph](http://graph.microsoft.io/) など) で取得されたアイテム ID は、Exchange Web サービス (EWS) に使用される形式とは異なる形式を使用します。 `convertToEwsId` メソッドは、REST 形式の ID を EWS 用の適切な形式に変換します。

##### パラメーター:

|名前| 型| 説明|
|---|---|---|
|`itemId`| String|Outlook REST API 形式のアイテム ID|
|`restVersion`| [Office.MailboxEnums.RestVersion](Office.MailboxEnums.md#restversion)|アイテム ID の取得に使用された Outlook REST API のバージョンを示す値。|

##### 要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](./tutorial-api-requirement-sets.md)| 1.3|
|[最小のアクセス許可レベル](../../docs/outlook/understanding-outlook-add-in-permissions.md)| Restricted|
|適用可能な Outlook のモード| 作成または読み取り|

##### 戻り値:

型:String

##### 例

```
// Get an item's ID from a REST API
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the
// Outlook Mail API
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

####  convertToLocalClientTime(timeValue) → {[LocalClientTime](simple-types.md#localclienttime)}

クライアントのローカル時間で時間情報が含まれている辞書を取得します。

Outlook 用メール アプリや Outlook Web App で使う日付と時刻では、異なるタイム ゾーンを使うことができます。Outlook では、クライアント コンピューターのタイム ゾーンを使います。Outlook Web App では、Exchange 管理センター (EAC) で設定されたタイム ゾーンを使います。ユーザー インターフェイスに表示される値が、常にユーザーが期待するタイム ゾーンと一致するように日付と時刻の値を処理する必要があります。

Outlook でメール アプリが実行されている場合、`convertToLocalClientTime` メソッドは、クライアント コンピューターのタイム ゾーンに設定された値のディクショナリ オブジェクトを返します。 Office Web Apps でメール アプリが実行されている場合、`convertToLocalClientTime` メソッドは、EAC に指定されたタイム ゾーンに設定された値のディクショナリ オブジェクトを返します。

##### パラメーター:

|名前| 型| 説明|
|---|---|---|
|`timeValue`| Date|日付オブジェクト|

##### 要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](./tutorial-api-requirement-sets.md)| 1.0|
|[最小のアクセス許可レベル](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用可能な Outlook のモード| 作成または読み取り|

##### 戻り値:

型:[LocalClientTime](simple-types.md#localclienttime)

####  convertToRestId(itemId, restVersion) → {String}

EWS 形式のアイテム ID を REST 形式に変換します。

EWS または `itemId` プロパティで取得されるアイテム ID は、REST API ([Outlook Mail API](https://msdn.microsoft.com/office/office365/APi/mail-rest-operations) や [Microsoft Graph](http://graph.microsoft.io/) など) に使用される形式とは異なる形式を使用します。 `convertToRestId` メソッドは、EWS 形式の ID を REST 用の適切な形式に変換します。

##### パラメーター:

|名前| 型| 説明|
|---|---|---|
|`itemId`| String|Exchange Web サービス (EWS) 形式のアイテム ID|
|`restVersion`| [Office.MailboxEnums.RestVersion](Office.MailboxEnums.md#restversion)|変換後の ID を使用する Outlook REST API のバージョンを示す値。|

##### 要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](./tutorial-api-requirement-sets.md)| 1.3|
|[最小のアクセス許可レベル](../../docs/outlook/understanding-outlook-add-in-permissions.md)| Restricted|
|適用可能な Outlook のモード| 作成または読み取り|

##### 戻り値:

型:String

##### 例

```
// Get the currently selected item's ID
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the
// Outlook Mail API
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

####  convertToUtcClientTime(input) → {Date}

時間情報が含まれているディクショナリから日付オブジェクトを取得します。

`convertToUtcClientTime` メソッドは、ローカルの日付と時刻を含むディクショナリを、ローカルの日付と時刻の正しい値を持つ日付オブジェクトに変換します。

##### パラメーター:

|名前| 型| 説明|
|---|---|---|
|`input`| [LocalClientTime](simple-types.md#localclienttime)|変換するローカル時刻の値。|

##### 要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](./tutorial-api-requirement-sets.md)| 1.0|
|[最小のアクセス許可レベル](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用可能な Outlook のモード| 作成または読み取り|

##### 戻り値:

時間が UTC で表現された日付オブジェクト。

<dl class="param-type">

<dt>型</dt>

<dd>Date</dd>

</dl>

####  displayAppointmentForm(itemId)

既存の予定を表示します。

`displayAppointmentForm` メソッドは、デスクトップ上の新しいウィンドウやモバイル デバイス上のダイアログ ボックスに既存の予定を開きます。

Outlook for Mac では、この方法を使って、定期的な系列の一部ではない単一の予定や定期的な系列のマスター予定を表示できます。ただし、系列のインスタンスは表示できません。これは、Outlook for Mac においては定期的な系列のインスタンスのプロパティ (アイテム ID を含む) にアクセスできないためです。

Outlook Web App では、このメソッドは指定されたフォームの本文が 32 KB 以下の文字数の場合にフォームを開きます。

指定のアイテム識別子が既存の予定を表していない場合は、クライアント コンピューターまたはデバイスで空のウィンドウが開き、エラー メッセージは返されません。

##### パラメーター:

|名前| 型| 説明|
|---|---|---|
|`itemId`| String|既存の予定の Exchange Web サービス (EWS) 識別子。|

##### 要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](./tutorial-api-requirement-sets.md)| 1.0|
|[最小のアクセス許可レベル](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用可能な Outlook のモード| 作成または読み取り|

##### 例

```
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

####  displayMessageForm(itemId)

既存のメッセージを表示します。

`displayMessageForm` メソッドは、デスクトップ上の新しいウィンドウやモバイル デバイス上のダイアログ ボックスに既存のメッセージを開きます。

Outlook Web App では、このメソッドは指定されたフォームの本文が 32 KB 以下の文字数の場合にフォームを開きます。

指定のアイテム識別子が既存のメッセージを表していない場合は、クライアント コンピューターにはメッセージは表示されず、エラー メッセージも返されません。

予定を表す `itemId` を含む `displayMessageForm` を使用しないでください。 既存の予定を表示するには、`displayAppointmentForm` メソッドを使用します。新しい予定を作成するフォームを表示するには、`displayNewAppointmentForm` メソッドを使用します。

##### パラメーター:

|名前| 型| 説明|
|---|---|---|
|`itemId`| String|既存のメッセージの Exchange Web サービス (EWS) 識別子。|

##### 要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](./tutorial-api-requirement-sets.md)| 1.0|
|[最小のアクセス許可レベル](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用可能な Outlook のモード| 作成または読み取り|

##### 例

```
Office.context.mailbox.displayMessageForm(messageId);
```

#### displayNewAppointmentForm(parameters)

新しい予定を作成するためのフォームを表示します。

`displayNewAppointmentForm` メソッドを使用すると、ユーザーが新しい予定または会議を作成できるフォームが開きます。 パラメーターを指定すると、予定のフォーム フィールドにパラメーターの内容が自動的に設定されます。

このメソッドは、Outlook Web App と OWA for Devices において、出席者フィールドが含まれるフォームを必ず表示します。入力引数として出席者を指定しないと、このメソッドにより **[保存]** ボタンのあるフォームが表示されます。出席者を指定した場合には、フォームにその出席者と **[送信]** ボタンが表示されます。

Outlook リッチ クライアントと Outlook RT で、`requiredAttendees`、`optionalAttendees`、または `resources` パラメーターに出席者またはリソースを指定し、このメソッドを実行すると、[**送信**] ボタンがある会議フォームが表示されます。 受信者を指定せずにこのメソッドを実行すると、[**保存して閉じる**] ボタンがある予定フォームが表示されます。

パラメーターのいずれかが指定のサイズ制限を超える場合、または不明なパラメーター名が指定されている場合は、例外がスローされます。

##### パラメーター:

|名前| 型| 説明|
|---|---|---|
|`parameters`| Object|新しい予定を記述するパラメーターのディクショナリ。<br/><br/>**プロパティ**<br/><table class="nested-table"><thead><tr><th>名前</th><th>型</th><th>説明</th></tr></thead><tbody><tr><td><code>requiredAttendees</code></td><td>Array.&lt;String&gt; &#124; Array.&lt;<a href="simple-types.md#emailaddressdetails">EmailAddressDetails</a>&gt;</td><td>予定に必要な各出席者について、メール アドレスを含む文字列の配列、または <code>EmailAddressDetails</code> オブジェクトを含む配列。 配列の上限は 100 エントリです。</td></tr><tr><td><code>optionalAttendees</code></td><td>Array.&lt;String&gt; &#124; Array.&lt;<a href="simple-types.md#emailaddressdetails">EmailAddressDetails</a>&gt;</td><td>予定の任意出席者の各電子メール アドレスが含まれている文字列から成る配列または EmailAddressDetails オブジェクトの配列。この配列は最大 100 個のエントリに制限されます。</td></tr><tr><td><code>start</code></td><td>Date</td><td>予定の開始日時を指定する Date オブジェクト。</td></tr><tr><td><code>end</code></td><td>Date</td><td>予定の終了日時を指定する Date オブジェクト。</td></tr><tr><td><code>location</code></td><td>String</td><td>予定の場所を含む文字列。 文字列は最大 255 文字に制限されます。</td></tr><tr><td><code>resources</code></td><td>Array.&lt;String&gt;</td><td>予定に必要なリソースを含む文字列の配列。 配列の上限は 100 エントリです。</td></tr><tr><td><code>subject</code></td><td>String</td><td>予定の件名を含む文字列です。 文字列は最大 255 文字に制限されます。</td></tr><tr><td><code>body</code></td><td>String</td><td>予定メッセージの本文。 本文の内容は、最大サイズが 32 KB に制限されます。</td></tr></tbody></table>|

##### 要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](./tutorial-api-requirement-sets.md)| 1.0|
|[最小のアクセス許可レベル](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用可能な Outlook のモード| 読み取り|

##### 例

```
var start = new Date();
var end = new Date();
end.setHours(start.getHours() + 1);

Office.context.mailbox.displayNewAppointmentForm(
  {
    requiredAttendees: ['bob@contoso.com'],
    optionalAttendees: ['sam@contoso.com'],
    start: start,
    end: end,
    location: 'Home',
    resources: ['projector@contoso.com'],
    subject: 'meeting',
    body: 'Hello World!'
  });
```

#### getCallbackTokenAsync(callback, [userContext])

Exchange Server から添付ファイルやアイテムを取得するために使うトークンを含む文字列を取得します。

`getCallbackTokenAsync` メソッドは、ユーザーのメールボックスをホストする Exchange Server から不透明なトークンを取得する非同期の呼び出しを行います。 コールバック トークンの有効期間は 5 分です。

トークンと予定の識別子またはアイテムの識別子をサードパーティ システムに渡すことができます。 サードパーティ システムは、トークンをベアラー承認トークンとして使用し、Exchange Web サービス (EWS) [GetAttachment](https://msdn.microsoft.com/en-us/library/office/aa494316.aspx) または [GetItem](https://msdn.microsoft.com/en-us/library/office/aa565934.aspx) 操作を呼び出して、添付ファイルまたはアイテムを返します。 たとえば、[選択したアイテムから添付ファイルを取得する](https://msdn.microsoft.com/EN-US/library/office/dn148008.aspx)ためにリモート サービスを作成できます。

アプリが閲覧モードで `getCallbackTokenAsync` メソッドを呼び出すには、アプリのマニフェスト内に **ReadItem** アクセス許可が指定されている必要があります。

新規作成モードでは、[`saveAsync`](Office.context.mailbox.item#saveAsync) メソッドを呼び出してアイテムの識別子を `getCallbackTokenAsync` メソッドに渡す必要があります。 アプリには、`saveAsync` メソッドを呼び出す **ReadWriteItem** アクセス許可が必要です。

##### パラメーター:

|名前| 型| 属性| 説明|
|---|---|---|---|
|`callback`| function||メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](simple-types.md#asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。 トークンは、`asyncResult.value` プロパティで文字列として提供されます。|
|`userContext`| Object| &lt;省略可能&gt;|非同期メソッドに渡される状態データです。|

##### 要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](./tutorial-api-requirement-sets.md)| 1.3|
|[最小のアクセス許可レベル](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用可能な Outlook のモード| 新規作成と閲覧|

##### 例

```js
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  getUserIdentityTokenAsync(callback, [userContext])

ユーザーと Office アドインを識別するトークンを取得します。


  `getUserIdentityTokenAsync` メソッドは、[アドインとユーザーをサード パーティのシステムで識別して認証](https://msdn.microsoft.com/EN-US/library/office/fp179828.aspx)することのできるトークンを返します。

##### パラメーター:

|名前| 型| 属性| 説明|
|---|---|---|---|
|`callback`| function||メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](simple-types.md#asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。

トークンは、`asyncResult.value` プロパティの文字列として提供されます。||`userContext`|Object|&lt;省略可能&gt;|非同期メソッドに渡される状態データです。|

##### 要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](./tutorial-api-requirement-sets.md)| 1.0|
|[最小のアクセス許可レベル](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用可能な Outlook のモード| 作成または読み取り|

##### 例

```js
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  makeEwsRequestAsync(data, callback, [userContext])

ユーザーのメールボックスをホストしている Exchange サーバー上の Exchange Web サービス (EWS) のサービスに対して非同期の要求を行ないます。

`makeEwsRequestAsync` メソッドは、アドインの代わりに Exchange に EWS 要求を送信します。

`makeEwsRequestAsync` メソッドでは、フォルダー関連アイテムを要求できません。

XML 要求では UTF-8 エンコードを指定する必要があります。

```
<?xml version="1.0" encoding="utf-8"?>
```

`makeEwsRequestAsync` メソッドを使用するには、アドインに **ReadWriteMailbox** アクセス許可が必要です。 **ReadWriteMailbox** アクセス許可と、`makeEwsRequestAsync` メソッドで呼び出せる EWS 操作の使い方については、「[ユーザーのメールボックスへのメール アドイン アクセスのアクセス許可を指定する](../../docs/outlook/understanding-outlook-add-in-permissions.md)」を参照してください。

**注**:サーバー管理者は、クライアント アクセス サーバーの EWS ディレクトリで `OAuthAuthentication` を true に設定して、`makeEwsRequestAsync` メソッドで EWS 要求を行うことができるようにする必要があります。

#### バージョンの相違点

バージョン 15.0.4535.1004 より前のバージョンの Outlook で実行しているメール アプリで `makeEwsRequestAsync` メソッドを使う場合は、エンコード値を `ISO-8859-1` に設定する必要があります。

```
<?xml version="1.0" encoding="iso-8859-1"?>
```

Outlook on the web でメール アプリを実行している場合は、エンコード値を設定する必要はありません。mailbox.diagnostics.hostName プロパティを使って、メール アプリを Outlook で実行しているのか、Outlook on the web で実行しているのかを確認できます。mailbox.diagnostics.hostVersion プロパティを使って、どのバージョンの Outlook を使って実行しているかを確認できます。

##### パラメーター:

|名前| 型| 属性| 説明|
|---|---|---|---|
|`data`| String||EWS 要求です。|
|`callback`| function||メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](simple-types.md#asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。

EWS 呼び出しの XML 結果は、`asyncResult.value` プロパティ内の文字列として提供されています。 結果のサイズが 1 MB を超える場合は、エラー メッセージが返されます。| |`userContext`|Object|&lt;省略可能&gt;|非同期メソッドに渡される状態データです。|

##### 要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](./tutorial-api-requirement-sets.md)| 1.0|
|[最小限のアクセス許可レベル](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteMailbox|
|適用可能な Outlook のモード| 作成または読み取り|

##### 例

次の例は、`GetItem` 操作を使ってアイテムの件名を取得するため、`makeEwsRequestAsync` を呼び出します。

```js
function getSubjectRequest(id) {
   // Return a GetItem operation request for the subject of the specified item.
   var request =
    '<?xml version="1.0" encoding="utf-8"?>' +
    '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
    '               xmlns:xsd="http://www.w3.org/2001/XMLSchema"' +
    '               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
    '               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
    '  <soap:Header>' +
    '    <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
    '  </soap:Header>' +
    '  <soap:Body>' +
    '    <GetItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">' +
    '      <ItemShape>' +
    '        <t:BaseShape>IdOnly</t:BaseShape>' +
    '        <t:AdditionalProperties>' +
    '            <t:FieldURI FieldURI="item:Subject"/>' +
    '        </t:AdditionalProperties>' +
    '      </ItemShape>' +
    '      <ItemIds><t:ItemId Id="' + id + '"/></ItemIds>' +
    '    </GetItem>' +
    '  </soap:Body>' +
    '</soap:Envelope>';

   return request;
}

function sendRequest() {
   // Create a local variable that contains the mailbox.
   Office.context.mailbox.makeEwsRequestAsync(
    getSubjectRequest(mailbox.item.itemId), callback);
}

function callback(asyncResult)  {
   var result = asyncResult.value;
   var context = asyncResult.asyncContext;

   // Process the returned response here.
}
```
