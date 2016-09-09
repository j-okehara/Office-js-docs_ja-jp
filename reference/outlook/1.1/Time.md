

# 時刻

`Time` オブジェクトは、作成モードで予定の [`start`](Office.context.mailbox.item.md#start-datetime) プロパティまたは [`end`](Office.context.mailbox.item.md#end-datetime) プロパティとして返されます。

##### 要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](../tutorial-api-requirement-sets.md)| 1.1|
|[最小のアクセス許可レベル](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用可能な Outlook のモード| 作成|

### メソッド

####  getAsync([options], callback)

予定の開始または終了の時刻を取得します。

##### パラメーター:

|名前| 型| 属性| 説明|
|---|---|---|---|
|`options`| Object| &lt;optional&gt;|次のプロパティのうち 1 つ以上を含むオブジェクト リテラルです。<br/><br/>**プロパティ**<br/><table class="nested-table"><thead><tr><th>名前</th><th>型</th><th>属性</th><th>説明</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;optional&gt;</td><td>開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</td></tr></tbody></table>|
|`callback`| function||メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](simple-types.md#asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。

日付と時刻は、`asyncResult.value` プロパティの Date オブジェクトとして指定されます。値は、世界協定時刻 (UTC) です。[`convertToLocalClientTime`](Office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) メソッドを使用して、UTC 時刻をローカル時刻に変換できます。|

##### 要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](../tutorial-api-requirement-sets.md)| 1.1|
|[最小のアクセス許可レベル](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用可能な Outlook のモード| 作成|
####  setAsync(dateTime, [options], [callback])

予定の開始または終了の時刻を設定します。

`setAsync` メソッドが [`start`](Office.context.mailbox.item.md#start-datetime) プロパティで呼び出された場合、[`end`](Office.context.mailbox.item.md#end-datetime) プロパティは以前に設定された予定の期間を保持するために調整されます。`setAsync` メソッドが `end` プロパティで呼び出された場合、予定の期間は新しい終了時刻に拡張されます。

この時刻は UTC でなければなりません。[`convertToUtcClientTime`](Office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、正確な UTC 時刻を取得できます。

##### パラメーター:

|名前| 型| 属性| 説明|
|---|---|---|---|
|`dateTime`| 日付||世界協定時刻 (UTC) の Date オブジェクト。|
|`options`| Object| &lt;optional&gt;|次のプロパティのうち 1 つ以上を含むオブジェクト リテラルです。<br/><br/>**プロパティ**<br/><table class="nested-table"><thead><tr><th>名前</th><th>型</th><th>属性</th><th>説明</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;optional&gt;</td><td>開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</td></tr></tbody></table>|
|`callback`| function| &lt;optional&gt;|メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](simple-types.md#asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。 <br/>日付と時刻の設定が失敗した場合、`asyncResult.error` プロパティにはエラー コードが格納されます。<br/><table class="nested-table"><thead><tr><th>エラー コード</th><th>説明</th></tr></thead><tbody><tr><td><code>InvalidEndTime</code></td><td>予定終了時間が、予定開始時刻より前に設定されています。</td></tr></tbody></table>|

##### 要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](../tutorial-api-requirement-sets.md)| 1.1|
|[最小限のアクセス許可レベル](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteItem|
|適用可能な Outlook のモード| 作成|

##### 例

次の例では、予定の開始時刻を設定します。

```js
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
