

# <a name="notificationmessages"></a>NotificationMessages

## <a name="notificationmessages"></a>NotificationMessages

`NotificationMessages` オブジェクトは、アイテムの [`notificationMessages`](Office.context.mailbox.item.md#notificationmessages-notificationmessages) プロパティとして返されます。

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](./tutorial-api-requirement-sets.md)| 1.3|
|[最小限のアクセス許可レベル](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用可能な Outlook のモード| 作成または読み取り|

### <a name="methods"></a>メソッド

####  <a name="addasync(key,-jsonmessage,-[options],-[callback])"></a>addAsync(key, JSONmessage, [options], [callback])

アイテムに通知を追加します。

メッセージあたりの最大通知数は 5 です。その数より多く設定すると、`NumberOfNotificationMessagesExceeded` エラーが返されます。

##### <a name="parameters:"></a>パラメーター:

|名前| 型| 属性| 説明|
|---|---|---|---|
|`key`| String||この通知メッセージを参照するために使用される、開発者が指定したキー。開発者は、このキーを使用して後ほどこのメッセージを変更できます。32 文字以内にしてください。|
|`JSONmessage`| オブジェクト||アイテムに追加される通知メッセージを格納する JSON オブジェクト。次のプロパティで構成されます。<br/><br/>**プロパティ**<br/><table class="nested-table"><thead><tr><th>名前</th><th>型</th><th>説明</th></tr></thead><tbody><tr><td><code>type</code></td><td><a href="Office.MailboxEnums.md#.ItemNotificationMessageType">Office.MailboxEnums.ItemNotificationMessageType</a></td><td>メッセージの型を指定します。型が <code>ProgressIndicator</code> または <code>ErrorMessage</code> である場合、アイコンが自動的に提供され、メッセージは永続的ではありません。したがって、icon プロパティと persistent プロパティは、これらの型のメッセージでは無効になります。それらを指定すると、<code>ArgumentException</code> が生じます。型が <code>ProgressIndicator</code> である場合、開発者は操作の完了時に進行状況インジケーターを置き換える必要があります。</td></tr><tr><td><code>icon</code></td><td>String</td><td><code>Resource</code>セクションのマニフェストで定義されているアイコンへの参照。情報バー領域に表示されます。これは型が <code>InformationalMessage</code> である場合にのみ適用可能です。サポートされていない型にこのパラメーターを指定すると例外が生じます。</td></tr><tr><td><code>message</code></td><td>String</td><td>通知メッセージのテキスト。最大の長さは 150 文字です。開発者が、長めの文字列を渡した場合、<code>ArgumentOutOfRange</code> 例外がスローされます。</td></tr><tr><td><code>persistent</code></td><td>ブール型 (Boolean)</td><td>型が <code>InformationalMessage</code> の場合にのみ適用可能。<code>true</code> の場合、メッセージはこのアドインによって削除されるか、ユーザーが非表示にするまで残されます。<code>false</code> の場合、ユーザーが別のアイテムに移動すると削除されます。エラーの通知の場合、メッセージはユーザーが 1 回表示するまで残されます。このパラメーターをサポートされない型に指定すると、例外がスローされます。</td></tr></tbody></table>|
|`options`| Object| &lt;optional&gt;|次のプロパティのうち 1 つ以上を含むオブジェクト リテラルです。<br/><br/>**プロパティ**<br/><table class="nested-table"><thead><tr><th>名前</th><th>型</th><th>属性</th><th>説明</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;optional&gt;</td><td>開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</td></tr></tbody></table>|
|`callback`| function| &lt;optional&gt;|メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](simple-types.md#asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。 |

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](./tutorial-api-requirement-sets.md)| 1.3|
|[最小限のアクセス許可レベル](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用可能な Outlook のモード| 作成または読み取り|

##### <a name="example"></a>例

```
// Create three notifications, each with a different key
Office.context.mailbox.item.notificationMessages.addAsync("progress", {
  type: "progressIndicator",
  message : "An add-in is processing this message."
});
Office.context.mailbox.item.notificationMessages.addAsync("information", {
  type: "informationalMessage",
  message : "The add-in processed this message.",
  icon : "iconid",
  persistent: false
});
Office.context.mailbox.item.notificationMessages.addAsync("error", {
  type: "errorMessage",
  message : "The add-in failed to process this message."
});
```

####  <a name="getallasync([options],-[callback])"></a>getAllAsync([options], [callback])

アイテムのすべてのキーとメッセージを返します。

##### <a name="parameters:"></a>パラメーター:

|名前| 型| 属性| 説明|
|---|---|---|---|
|`options`| Object| &lt;optional&gt;|次のプロパティのうち 1 つ以上を含むオブジェクト リテラルです。<br/><br/>**プロパティ**<br/><table class="nested-table"><thead><tr><th>名前</th><th>型</th><th>属性</th><th>説明</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;optional&gt;</td><td>開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</td></tr></tbody></table>|
|`callback`| function| &lt;optional&gt;|メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](simple-types.md#asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。

正常な完了時に、`asyncResult.value` プロパティには [`NotificationMessageDetails`](simple-types.md#notificationmessagedetails) オブジェクトの配列が含まれます。|

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](./tutorial-api-requirement-sets.md)| 1.3|
|[最小限のアクセス許可レベル](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用可能な Outlook のモード| 作成または読み取り|

##### <a name="example"></a>例

```
// Get all notifications
Office.context.mailbox.item.notificationMessages.getAllAsync(function (asyncResult) {
  if (asyncResult.status != "failed") {
    Office.context.mailbox.item.notificationMessages.replaceAsync( "notifications", {
      type: "informationalMessage",
      message : "Found " + asyncResult.value.length + " notifications.",
      icon : "iconid",
      persistent: false
    });
  }
});
```

####  <a name="removeasync(key,-[options],-[callback])"></a>removeAsync(key, [options], [callback])

アイテムの通知メッセージを削除します。

##### <a name="parameters:"></a>パラメーター:

|名前| 型| 属性| 説明|
|---|---|---|---|
|`key`| String||通知メッセージを削除するためのキー。|
|`options`| Object| &lt;optional&gt;|次のプロパティのうち 1 つ以上を含むオブジェクト リテラルです。<br/><br/>**プロパティ**<br/><table class="nested-table"><thead><tr><th>名前</th><th>型</th><th>属性</th><th>説明</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;optional&gt;</td><td>開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</td></tr></tbody></table>|
|`callback`| function| &lt;optional&gt;|メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](simple-types.md#asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。

キーが見つからない場合、`KeyNotFound` プロパティに `asyncResult.error` エラーが返されます。|

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](./tutorial-api-requirement-sets.md)| 1.3|
|[最小限のアクセス許可レベル](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用可能な Outlook のモード| 作成または読み取り|

##### <a name="example"></a>例

```
// Remove a notification
Office.context.mailbox.item.notificationMessages.removeAsync("progress");
```

####  <a name="replaceasync(key,-jsonmessage,-[options],-[callback])"></a>replaceAsync(key, JSONmessage, [options], [callback])

指定のキーが含まれる通知メッセージを別のメッセージに置換します。

指定したキーを持つ通知メッセージが存在しない場合は、`replaceAsync` が通知を追加します。

##### <a name="parameters:"></a>パラメーター:

|名前| 型| 属性| 説明|
|---|---|---|---|
|`key`| String||置換する通知メッセージのキー。32 文字以内にする必要があります。|
|`JSONmessage`| オブジェクト||既存のメッセージを置換する新しい通知メッセージを格納する JSON オブジェクト。次のプロパティで構成されます。<br/><br/>**プロパティ**<br/><table class="nested-table"><thead><tr><th>名前</th><th>型</th><th>説明</th></tr></thead><tbody><tr><td><code>type</code></td><td><a href="Office.MailboxEnums.md#.ItemNotificationMessageType">Office.MailboxEnums.ItemNotificationMessageType</a></td><td>メッセージの型を指定します。型が <code>ProgressIndicator</code> または <code>ErrorMessage</code> である場合、アイコンが自動的に提供され、メッセージは永続的ではありません。したがって、icon プロパティと persistent プロパティは、これらの型のメッセージでは無効になります。それらを指定すると、<code>ArgumentException</code> が生じます。型が <code>ProgressIndicator</code> である場合、開発者は操作の完了時に進行状況インジケーターを置き換える必要があります。</td></tr><tr><td><code>icon</code></td><td>String</td><td><code>Resource</code>セクションのマニフェストで定義されているアイコンへの参照。情報バー領域に表示されます。これは型が <code>InformationalMessage</code> である場合にのみ適用可能です。サポートされていない型にこのパラメーターを指定すると例外が生じます。</td></tr><tr><td><code>message</code></td><td>String</td><td>通知メッセージのテキスト。最大の長さは 150 文字です。開発者が、長めの文字列を渡した場合、<code>ArgumentOutOfRange</code> 例外がスローされます。</td></tr><tr><td><code>persistent</code></td><td>ブール型 (Boolean)</td><td>型が <code>InformationalMessage</code> の場合にのみ適用可能。<code>true</code> の場合、メッセージはこのアドインによって削除されるか、ユーザーが非表示にするまで残されます。<code>false</code> の場合、ユーザーが別のアイテムに移動すると削除されます。エラーの通知の場合、メッセージはユーザーが 1 回表示するまで残されます。このパラメーターをサポートされない型に指定すると、例外がスローされます。</td></tr></tbody></table>|
|`options`| Object| &lt;optional&gt;|次のプロパティのうち 1 つ以上を含むオブジェクト リテラルです。<br/><br/>**プロパティ**<br/><table class="nested-table"><thead><tr><th>名前</th><th>型</th><th>属性</th><th>説明</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;optional&gt;</td><td>開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</td></tr></tbody></table>|
|`callback`| function| &lt;optional&gt;|メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](simple-types.md#asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。 |

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](./tutorial-api-requirement-sets.md)| 1.3|
|[最小限のアクセス許可レベル](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用可能な Outlook のモード| 作成または読み取り|

##### <a name="example"></a>例

```
// Replace a notification with an informational notification
Office.context.mailbox.item.notificationMessages.replaceAsync("progress", {
  type: "informationalMessage",
  message : "The message was processed successfully.",
  icon : "iconid",
  persistent: false
});
```
