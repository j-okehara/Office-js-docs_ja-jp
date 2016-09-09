

# 本文

`body` オブジェクトは、メッセージまたは予定の内容を追加および更新するためのメソッドを提供します。選択したアイテムの `body` プロパティで返されます。

##### 要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](../tutorial-api-requirement-sets.md)| 1.1|
|[最小のアクセス許可レベル](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用可能な Outlook のモード| 作成または読み取り|

### メソッド

####  getTypeAsync([options], [callback])

コンテンツの形式が HTML とテキストのどちらであるかを示す値を取得します。

##### パラメーター:

|名前| 型| 属性| 説明|
|---|---|---|---|
|`options`| Object| &lt;optional&gt;|次のプロパティのうち 1 つ以上を含むオブジェクト リテラルです。<br/><br/>**プロパティ**<br/><table class="nested-table"><thead><tr><th>名前</th><th>型</th><th>属性</th><th>説明</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;optional&gt;</td><td>開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</td></tr></tbody></table>|
|`callback`| function| &lt;optional&gt;|メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](simple-types.md#asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。

コンテンツ タイプは、`asyncResult.value` プロパティの [CoercionType](Office.md#coerciontype-string) の 1 つとして返されます。|

##### 要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](../tutorial-api-requirement-sets.md)| 1.1|
|[最小のアクセス許可レベル](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用可能な Outlook のモード| 作成|
####  prependAsync(data, [options], [callback])

アイテム本文の先頭に指定の内容を追加します。

`prependAsync` メソッドは、指定された文字列をアイテム本体の先頭に挿入します。`prependAsync` メソッドの呼び出しは、本文の先頭に挿入ポイントを指定して [`setSelectedDataAsync`](Body.md#setselecteddataasyncdata-options-callback) メソッドを呼び出すのと同じです。

リンクに HTML マークアップを含める場合、アンカー (`<a>`) の `id` 属性を `LPNoLP` に設定することで、オンライン リンク プレビューを無効にできます。 たとえば次のようにします。

```
Office.context.mailbox.item.body.prependAsync(
  '<a id="LPNoLP" href="http://www.contoso.com">Click here!</a>',
  {coercionType: Office.CoercionType.Html},
  callback);
```

##### パラメーター:

|名前| 型| 属性| 説明|
|---|---|---|---|
|`data`| String||本文の先頭に挿入する文字列。文字列の最大長は 1,000,000 文字です。|
|`options`| Object| &lt;optional&gt;|次のプロパティのうち 1 つ以上を含むオブジェクト リテラルです。<br/><br/>**プロパティ**<br/><table class="nested-table"><thead><tr><th>名前</th><th>型</th><th>属性</th><th>説明</th></tr></thead><tbody><tr><td><code>coercionType</code></td><td><a href="Office.md#coerciontype-string">Office.CoercionType</a></td><td>&lt;optional&gt;</td><td>本文の必要な形式です。<code>data</code> パラメーター内の文字列は、この形式に変換されます。</td></tr><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;optional&gt;</td><td>開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</td></tr></tbody></table>|
|`callback`| function| &lt;optional&gt;|メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](simple-types.md#asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。 <br/>検出されたすべてのエラーは `asyncResult.error` プロパティに表示されます。<br/><table class="nested-table"><thead><tr><th>エラー コード</th><th>説明</th></tr></thead><tbody><tr><td><code>DataExceedsMaximumSize</code></td><td><code>data</code> パラメーターは 1,000,000 文字よりも長くなっています。</td></tr></tbody></table>|

##### 要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](../tutorial-api-requirement-sets.md)| 1.1|
|[最小限のアクセス許可レベル](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteItem|
|適用可能な Outlook のモード| 作成|
####  setSelectedDataAsync(data, [options], [callback])

本文の選択部分を、指定のテキストに置き換えます。

`setSelectedDataAsync` メソッドは、アイテムの本文のカーソル位置に指定された文字列を挿入します。また、エディターでテキストが選択されている場合は、選択されたテキストを置換します。アイテムの本文中にカーソルが存在しないか、UI でアイテムの本文がフォーカスを喪失している場合、文字列は本文の先頭に挿入されます。

リンクに HTML マークアップを含める場合、アンカー (`<a>`) の `id` 属性を `LPNoLP` に設定することで、オンライン リンク プレビューを無効にできます。 たとえば次のようにします。

```
Office.context.mailbox.item.body.setSelectedDataAsync(
  '<a id="LPNoLP" href="http://www.contoso.com">Click here!</a>',
  {coercionType: Office.CoercionType.Html},
  callback);
```

##### パラメーター:

|名前| 型| 属性| 説明|
|---|---|---|---|
|`data`| String||本文に挿入する文字列。文字列の最大長は 1,000,000 文字です。|
|`options`| Object| &lt;optional&gt;|次のプロパティのうち 1 つ以上を含むオブジェクト リテラルです。<br/><br/>**プロパティ**<br/><table class="nested-table"><thead><tr><th>名前</th><th>型</th><th>属性</th><th>説明</th></tr></thead><tbody><tr><td><code>coercionType</code></td><td><a href="Office.md#coerciontype-string">Office.CoercionType</a></td><td>&lt;optional&gt;</td><td>本文の必要な形式です。<code>data</code> パラメーター内の文字列は、この形式に変換されます。</td></tr><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;optional&gt;</td><td>開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</td></tr></tbody></table>|
|`callback`| function| &lt;optional&gt;|メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](simple-types.md#asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。 <br/>検出されたすべてのエラーは `asyncResult.error` プロパティに表示されます。<br/><table class="nested-table"><thead><tr><th>エラー コード</th><th>説明</th></tr></thead><tbody><tr><td><code>DataExceedsMaximumSize</code></td><td><code>data</code> パラメーターは 1,000,000 文字より長くなっています。</td></tr><tr><td><code>InvalidFormatError</code></td><td>本文の種類は HTML に設定されており、データ パラメーターにはプレーンテキストが含まれます。</td></tr></tbody></table>|

##### 要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](../tutorial-api-requirement-sets.md)| 1.1|
|[最小限のアクセス許可レベル](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteItem|
|適用可能な Outlook のモード| 作成|
