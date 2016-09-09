

# Location

Outlook アドインで会議の場所を取得し設定する方法を提供します。

##### 要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](./tutorial-api-requirement-sets.md)| 1.1|
|[最小のアクセス許可レベル](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用可能な Outlook のモード| 作成|

### メソッド

####  getAsync([options], callback)

予定の場所を取得します。

`getAsync` メソッドは、Exchange サーバーへの非同期呼び出しを開始し、予定の場所を取得します。

##### パラメーター:

|名前| 型| 属性| 説明|
|---|---|---|---|
|`options`| Object| &lt;optional&gt;|次のプロパティのうち 1 つ以上を含むオブジェクト リテラルです。<br/><br/>**プロパティ**<br/><table class="nested-table"><thead><tr><th>名前</th><th>型</th><th>属性</th><th>説明</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;optional&gt;</td><td>開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</td></tr></tbody></table>|
|`callback`| function||メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](simple-types.md#asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。

予定の場所は、`asyncResult.value` プロパティに文字列として指定されます。|

##### 要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](./tutorial-api-requirement-sets.md)| 1.1|
|[最小のアクセス許可レベル](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用可能な Outlook のモード| 作成|
####  setAsync(location, [options], [callback])

予定の場所を設定します。

`setAsync` メソッドは、予定の場所を設定する Exchange サーバーへの非同期呼び出しを開始します。予定の場所を設定すると、現在の場所が上書きされます。

##### パラメーター:

|名前| 型| 属性| 説明|
|---|---|---|---|
|`location`| String||予定の場所。文字列の最大長は 255 文字です。|
|`options`| Object| &lt;optional&gt;|次のプロパティのうち 1 つ以上を含むオブジェクト リテラルです。<br/><br/>**プロパティ**<br/><table class="nested-table"><thead><tr><th>名前</th><th>型</th><th>属性</th><th>説明</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;optional&gt;</td><td>開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</td></tr></tbody></table>|
|`callback`| function| &lt;optional&gt;|メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](simple-types.md#asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。 <br/>場所の設定に失敗すると、`asyncResult.error` プロパティにエラー コードが格納されます。<br/><table class="nested-table"><thead><tr><th>エラー コード</th><th>説明</th></tr></thead><tbody><tr><td><code>DataExceedsMaximumSize</code></td><td><code>location</code> パラメーターは 255 文字以上です。</td></tr></tbody></table>|

##### 要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](./tutorial-api-requirement-sets.md)| 1.1|
|[最小のアクセス許可レベル](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用可能な Outlook のモード| 作成|
