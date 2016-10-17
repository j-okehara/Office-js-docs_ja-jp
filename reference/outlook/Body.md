

# <a name="body"></a>本文

`body` オブジェクトは、メッセージまたは予定の内容を追加および更新するためのメソッドを提供します。選択したアイテムの `body` プロパティで返されます。

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](./tutorial-api-requirement-sets.md)| 1.1|
|[最小限のアクセス許可レベル](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用可能な Outlook のモード| 作成または読み取り|

### <a name="methods"></a>メソッド

####  <a name="getasync(coerciontype,-[options],-[callback])"></a>getAsync(coercionType, [options], [callback])

現在の本文を指定された形式で返します。

このメソッドは、現在の本文全体を `coercionType` に指定された形式で返します 。

##### <a name="parameters:"></a>パラメーター:

|名前| 型| 属性| 説明|
|---|---|---|---|
|`coercionType`| [Office.CoercionType](Office.md#coerciontype-string)||返される本文の形式です。|
|`options`| Object| &lt;optional&gt;|次のプロパティのうち 1 つ以上を含むオブジェクト リテラルです。<br/><br/>**プロパティ**<br/><table class="nested-table"><thead><tr><th>名前</th><th>型</th><th>属性</th><th>説明</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;optional&gt;</td><td>開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</td></tr></tbody></table>|
|`callback`| function| &lt;optional&gt;|メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](simple-types.md#asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。

本文は、`asyncResult.value` プロパティで要求された形式で提供されます。|

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](./tutorial-api-requirement-sets.md)| 1.3|
|[最小限のアクセス許可レベル](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用可能な Outlook のモード| 作成または読み取り|

##### <a name="examples"></a>例

この例では、メッセージの本文をプレーンテキストで取得します。

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext:"This is passed to the callback" },
  function callback(result) {
    // Do something with the result
  });
```

次の例は、コールバック関数に渡される `result` パラメーターの例です。

```js
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

####  <a name="gettypeasync([options],-[callback])"></a>getTypeAsync([options], [callback])

コンテンツの形式が HTML とテキストのどちらであるかを示す値を取得します。

##### <a name="parameters:"></a>パラメーター:

|名前| 型| 属性| 説明|
|---|---|---|---|
|`options`| Object| &lt;optional&gt;|次のプロパティのうち 1 つ以上を含むオブジェクト リテラルです。<br/><br/>**プロパティ**<br/><table class="nested-table"><thead><tr><th>名前</th><th>型</th><th>属性</th><th>説明</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;optional&gt;</td><td>開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</td></tr></tbody></table>|
|`callback`| function| &lt;optional&gt;|メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](simple-types.md#asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。

コンテンツ タイプは、`asyncResult.value` プロパティの [CoercionType](Office.md#coerciontype-string) の 1 つとして返されます。|

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](./tutorial-api-requirement-sets.md)| 1.1|
|[最小限のアクセス許可レベル](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用可能な Outlook のモード| 作成|
####  <a name="prependasync(data,-[options],-[callback])"></a>prependAsync(data, [options], [callback])

アイテム本文の先頭に指定の内容を追加します。

`prependAsync` メソッドは、指定された文字列をアイテム本体の先頭に挿入します。`prependAsync` メソッドの呼び出しは、本文の先頭に挿入ポイントを指定して [`setSelectedDataAsync`](#setselecteddataasyncdata-options-callback) メソッドを呼び出すのと同じです。

リンクに HTML マークアップを含める場合、アンカー (`<a>`) の `id` 属性を `LPNoLP` に設定することで、オンライン リンク プレビューを無効にできます。たとえば次のようにします。

```js
Office.context.mailbox.item.body.prependAsync(
  '<a id="LPNoLP" href="http://www.contoso.com">Click here!</a>',
  {coercionType: Office.CoercionType.Html},
  callback);
```

##### <a name="parameters:"></a>パラメーター:

|名前| 型| 属性| 説明|
|---|---|---|---|
|`data`| String||本文の先頭に挿入する文字列。文字列の最大長は 1,000,000 文字です。|
|`options`| Object| &lt;optional&gt;|次のプロパティのうち 1 つ以上を含むオブジェクト リテラルです。<br/><br/>**プロパティ**<br/><table class="nested-table"><thead><tr><th>名前</th><th>型</th><th>属性</th><th>説明</th></tr></thead><tbody><tr><td><code>coercionType</code></td><td><a href="Office.md#coerciontype-string">Office.CoercionType</a></td><td>&lt;optional&gt;</td><td>本文の必要な形式です。<code>data</code> パラメーター内の文字列は、この形式に変換されます。</td></tr><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;optional&gt;</td><td>開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</td></tr></tbody></table>|
|`callback`| function| &lt;optional&gt;|メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](simple-types.md#asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。 <br/>検出されたすべてのエラーは `asyncResult.error` プロパティに表示されます。<br/><table class="nested-table"><thead><tr><th>エラー コード</th><th>説明</th></tr></thead><tbody><tr><td><code>DataExceedsMaximumSize</code></td><td><code>data</code> パラメーターは 1,000,000 文字よりも長くなっています。</td></tr></tbody></table>|

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](./tutorial-api-requirement-sets.md)| 1.1|
|[最小限のアクセス許可レベル](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteItem|
|適用可能な Outlook のモード| 作成|
####  <a name="setasync(data,-[options],-[callback])"></a>setAsync(data, [options], [callback])

本文全体を指定されたテキストに置換します。

`setAsync` メソッドは、項目の既存の本文を指定の文字列に置換します。または、エディターでテキストを選択する場合には、選択したテキストを置換します。

リンクに HTML マークアップを含める場合、アンカー (`<a>`) の `id` 属性を `LPNoLP` に設定することで、オンライン リンク プレビューを無効にできます。たとえば次のようにします。

```js
Office.context.mailbox.item.body.setAsync(
  '<a id="LPNoLP" href="http://www.contoso.com">Click here!</a>',
  {coercionType: Office.CoercionType.Html},
  callback);
```

##### <a name="parameters:"></a>パラメーター:

|名前| 型| 属性| 説明|
|---|---|---|---|
|`data`| String||既存の本文を置換する文字列。文字列の長さは 1,000,000 文字までに制限されています。|
|`options`| Object| &lt;optional&gt;|次のプロパティのうち 1 つ以上を含むオブジェクト リテラルです。<br/><br/>**プロパティ**<br/><table class="nested-table"><thead><tr><th>名前</th><th>型</th><th>属性</th><th>説明</th></tr></thead><tbody><tr><td><code>coercionType</code></td><td><a href="Office.md#coerciontype-string">Office.CoercionType</a></td><td>&lt;optional&gt;</td><td>本文の必要な形式です。<code>data</code> パラメーター内の文字列は、この形式に変換されます。</td></tr><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;optional&gt;</td><td>開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</td></tr></tbody></table>|
|`callback`| function| &lt;optional&gt;|メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](simple-types.md#asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。 <br/>検出されたすべてのエラーは `asyncResult.error` プロパティに表示されます。<br/><table class="nested-table"><thead><tr><th>エラー コード</th><th>説明</th></tr></thead><tbody><tr><td><code>DataExceedsMaximumSize</code></td><td><code>data</code> パラメーターは 1,000,000 文字より長くなっています。</td></tr><tr><td><code>InvalidFormatError</code></td><td><code>options.coercionType</code> パラメーターは <code>Office.CoercionType.Html</code> に設定されており、メッセージ本文はプレーンテキストです。</td></tr></tbody></table>|

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](./tutorial-api-requirement-sets.md)| 1.3|
|[最小限のアクセス許可レベル](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteItem|
|適用可能な Outlook のモード| 作成|

##### <a name="examples"></a>例

次の例は、本体を HTML コンテンツで置き換えます。

```js
Office.context.mailbox.item.body.setAsync(
  "<b>(replaces all body, including threads you are replying to that may be on the bottom)</b>",
  { coercionType:"html", asyncContext:"This is passed to the callback" },
  function callback(result) {
    // Process the result
  });
```

次の例は、コールバック関数に渡される `result` パラメーターの例です。

```js
{
  "value":null,
  "status":"succeeded",
  "asyncContext":"This is passed to the callback"
}
```

####  <a name="setselecteddataasync(data,-[options],-[callback])"></a>setSelectedDataAsync(data, [options], [callback])

本文の選択部分を、指定のテキストに置き換えます。

`setSelectedDataAsync` メソッドは、アイテムの本文のカーソル位置に指定された文字列を挿入します。また、エディターでテキストが選択されている場合は、選択されたテキストを置換します。アイテムの本文中にカーソルが存在しないか、UI でアイテムの本文がフォーカスを喪失している場合、文字列は本文の先頭に挿入されます。

リンクに HTML マークアップを含める場合、アンカー (`<a>`) の `id` 属性を `LPNoLP` に設定することで、オンライン リンク プレビューを無効にできます。たとえば次のようにします。

```js
Office.context.mailbox.item.body.setSelectedDataAsync(
  '<a id="LPNoLP" href="http://www.contoso.com">Click here!</a>',
  {coercionType: Office.CoercionType.Html},
  callback);
```

##### <a name="parameters:"></a>パラメーター:

|名前| 型| 属性| 説明|
|---|---|---|---|
|`data`| String||本文に挿入する文字列。文字列の最大長は 1,000,000 文字です。|
|`options`| Object| &lt;optional&gt;|次のプロパティのうち 1 つ以上を含むオブジェクト リテラルです。<br/><br/>**プロパティ**<br/><table class="nested-table"><thead><tr><th>名前</th><th>型</th><th>属性</th><th>説明</th></tr></thead><tbody><tr><td><code>coercionType</code></td><td><a href="Office.md#coerciontype-string">Office.CoercionType</a></td><td>&lt;optional&gt;</td><td>本文の必要な形式です。<code>data</code> パラメーター内の文字列は、この形式に変換されます。</td></tr><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;optional&gt;</td><td>開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</td></tr></tbody></table>|
|`callback`| function| &lt;optional&gt;|メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](simple-types.md#asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。 <br/>検出されたすべてのエラーは `asyncResult.error` プロパティに表示されます。<br/><table class="nested-table"><thead><tr><th>エラー コード</th><th>説明</th></tr></thead><tbody><tr><td><code>DataExceedsMaximumSize</code></td><td><code>data</code> パラメーターは 1,000,000 文字より長くなっています。</td></tr><tr><td><code>InvalidFormatError</code></td><td>本文の種類は HTML に設定されており、データ パラメーターにはプレーンテキストが含まれます。</td></tr></tbody></table>|

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](./tutorial-api-requirement-sets.md)| 1.1|
|[最小限のアクセス許可レベル](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteItem|
|適用可能な Outlook のモード| 作成|
