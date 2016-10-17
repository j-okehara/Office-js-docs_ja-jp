

# <a name="recipients"></a>受信者

##### <a name="requirements"></a>Requirements

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](./tutorial-api-requirement-sets.md)| 1.1|
|[最小限のアクセス許可レベル](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用可能な Outlook のモード| 作成|

### <a name="methods"></a>メソッド

####  <a name="addasync(recipients,-[options],-[callback])"></a>addAsync(recipients, [options], [callback])

予定やメッセージの既存の受信者に、受信者のリストを追加します。

`recipients` パラメーターには、次のいずれかの配列を指定できます。

*   SMTP 電子メールアドレスを含む文字列
*   `EmailUser` オブジェクト
*   `EmailAddressDetails` オブジェクト

##### <a name="parameters:"></a>パラメーター:

|名前| 型| 属性| 説明|
|---|---|---|---|
|`recipients`| Array.&lt;(String&#124;[EmailUser](simple-types.md#emailuser)&#124;[EmailAddressDetails](simple-types.md#emailaddressdetails))&gt;||受信者リストに追加する受信者。|
|`options`| オブジェクト| &lt;optional&gt;|次のプロパティのうち 1 つ以上を含むオブジェクト リテラルです。<br/><br/>**プロパティ**<br/><table class="nested-table"><thead><tr><th>名前</th><th>型</th><th>属性</th><th>説明</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;optional&gt;</td><td>開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</td></tr></tbody></table>|
|`callback`| function| &lt;optional&gt;|メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](simple-types.md#asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。 <br/>受信者の追加に失敗すると、`asyncResult.error` プロパティにエラー コードが格納されます。<br/><table class="nested-table"><thead><tr><th>エラー コード</th><th>説明</th></tr></thead><tbody><tr><td>`NumberOfRecipientsExceeded</td><td>受信者の数が 100 エントリを超えました。</td></tr></tbody></table>|

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](./tutorial-api-requirement-sets.md)| 1.1|
|[最小限のアクセス許可レベル](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteItem|
|適用可能な Outlook のモード| 作成|

##### <a name="example"></a>例

次の例は、`EmailUser` オブジェクトの配列を作成し、それらをメッセージの宛先の受信者に追加します。

```
var newRecipients = [
  {
    "displayName": "Allie Bellew",
    "emailAddress": "allieb@contoso.com"
  },
  {
    "displayName": "Alex Darrow",
    "emailAddress": "alexd@contoso.com"
  }
];

Office.context.mailbox.item.to.addAsync(newRecipients, function(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    showMessage("Recipients added");
  }
});
```

####  <a name="getasync([options],-callback)"></a>getAsync([options], callback)

予定やメッセージの受信者リストを取得します。

##### <a name="parameters:"></a>パラメーター:

|名前| 型| 属性| 説明|
|---|---|---|---|
|`options`| Object| &lt;optional&gt;|次のプロパティのうち 1 つ以上を含むオブジェクト リテラルです。<br/><br/>**プロパティ**<br/><table class="nested-table"><thead><tr><th>名前</th><th>型</th><th>属性</th><th>説明</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;optional&gt;</td><td>開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</td></tr></tbody></table>|
|`callback`| function||メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](simple-types.md#asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。

呼び出しが完了すると、`asyncResult.value` プロパティには [`EmailAddressDetails`](simple-types.md#emailaddressdetails) オブジェクトの配列が含まれます。|

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](./tutorial-api-requirement-sets.md)| 1.1|
|[最小限のアクセス許可レベル](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用可能な Outlook のモード| 作成|

##### <a name="example"></a>例

次の例は、会議への任意出席者を取得します。

```js
Office.context.mailbox.item.optionalAttendees.getAsync(function(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    var msg = "";
    result.value.forEach(function(recip, index) {
      msg = msg + recip.displayName + " (" + recip.emailAddress + ");";
    });
    showMessage(msg);
  }
});
```

####  <a name="setasync(recipients,-[options],-callback)"></a>setAsync(recipients, [options], callback)

予定やメッセージの受信者リストを設定します。

`setAsync` メソッドは、現在の受信者のリストを上書きします。

`recipients` パラメーターには、次のいずれかの配列を指定できます。

*   SMTP 電子メールアドレスを含む文字列
*   `EmailUser` オブジェクト
*   `EmailAddressDetails` オブジェクト

##### <a name="parameters:"></a>パラメーター:

|名前| 型| 属性| 説明|
|---|---|---|---|
|`recipients`| Array.&lt;(String&#124;[EmailUser](simple-types.md#emailuser)&#124;[EmailAddressDetails](simple-types.md#emailaddressdetails))&gt;||受信者リストに追加する受信者。|
|`options`| オブジェクト| &lt;optional&gt;|次のプロパティのうち 1 つ以上を含むオブジェクト リテラルです。<br/><br/>**プロパティ**<br/><table class="nested-table"><thead><tr><th>名前</th><th>型</th><th>属性</th><th>説明</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;optional&gt;</td><td>開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</td></tr></tbody></table>|
|`callback`| function||メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](simple-types.md#asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。 <br/>受信者の設定に失敗した場合、`asyncResult.error` プロパティには、データの追加時に発生したエラーを示すコードが含まれます。<br/><table class="nested-table"><thead><tr><th>エラー コード</th><th>説明</th></tr></thead><tbody><tr><td>`NumberOfRecipientsExceeded</td><td>受信者の数が 100 エントリを超えました。</td></tr></tbody></table>|

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](./tutorial-api-requirement-sets.md)| 1.1|
|[最小限のアクセス許可レベル](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteItem|
|適用可能な Outlook のモード| 作成|

##### <a name="example"></a>例

次の例は、`EmailUser` オブジェクトの配列を作成し、メッセージの CC 宛先をその配列に置き換えます。

```
var newRecipients = [
  {
    "displayName": "Allie Bellew",
    "emailAddress": "allieb@contoso.com"
  },
  {
    "displayName": "Alex Darrow",
    "emailAddress": "alexd@contoso.com"
  }
];

Office.context.mailbox.item.cc.setAsync(newRecipients, function(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    showMessage("Recipients overwritten");
  }
});
```
