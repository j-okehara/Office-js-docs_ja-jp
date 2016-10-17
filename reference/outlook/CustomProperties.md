

# <a name="customproperties"></a>CustomProperties

`CustomProperties` オブジェクトが表すカスタム プロパティは、特定のアイテムに固有であり、Outlook 用のメール アドインに固有です。たとえば、メール アドインは、アドインをアクティブ化する現在のメール メッセージに固有のいくつかのデータを保存する必要があります。ユーザーが、将来同じメッセージを再び取り上げ、もう一度メール アドインをアクティブ化する場合、アドインは、カスタム プロパティとして保存されていたデータを取得することができます。

Outlook for Mac はカスタム プロパティをキャッシュに入れないため、ユーザーのネットワークが使用できなくなると、メール アドインは、そうしたカスタム プロパティにアクセスできなくなります。

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](./tutorial-api-requirement-sets.md)| 1.0|
|[最小限のアクセス許可レベル](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用可能な Outlook のモード| 作成または読み取り|

### <a name="example"></a>例

次の例では、`loadCustomPropertiesAsync` メソッドを使用して、現在のアイテムに固有のカスタム プロパティを非同期的にロードする方法を示します。また、[`saveAsync`](#saveasynccallback-asynccontext) メソッドを使用して、これらのプロパティをサーバーに保存する方法も紹介します。カスタム プロパティをロードした後、この例では [`get`](CustomProperties.md#getname--string) メソッドを使用してカスタム プロパティ `myProp` を読み取り、[`set`](CustomProperties.md#setname-value) メソッドでカスタム プロパティ `otherProp` を書き込み、最後に `saveAsync` メソッドを呼び出して、カスタム プロパティを保存します。

```
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
    // After the DOM is loaded, add-in-specific code can run.
    var mailbox = Office.context.mailbox;
    mailbox.item.loadCustomPropertiesAsync(customPropsCallback);
  });
}
function customPropsCallback(asyncResult) {
  var customProps = asyncResult.value;
  var myProp = customProps.get("myProp");

  customProps.set("otherProp", "value");
  customProps.saveAsync(saveCallback);
}

function saveCallback(asyncResult) {
}
```

### <a name="methods"></a>メソッド

####  <a name="get(name)-→-{string}"></a>get(name) → {String}

指定したカスタム プロパティの値を返します。

##### <a name="parameters:"></a>パラメーター:

|名前| 型| 説明|
|---|---|---|
|`name`| String|取得するカスタム プロパティの名前。|

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](./tutorial-api-requirement-sets.md)| 1.0|
|[最小限のアクセス許可レベル](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用可能な Outlook のモード| 作成または読み取り|

##### <a name="returns:"></a>戻り値:

指定したカスタム プロパティの値。

<dl class="param-type">

<dt>型</dt>

<dd>String</dd>

</dl>

####  <a name="remove(name)"></a>remove(name)

カスタム プロパティ コレクションから指定のプロパティを削除します。

プロパティを完全に削除するには、`CustomProperties` オブジェクトの [`saveAsync`](CustomProperties.md#saveasynccallback-asynccontext) メソッドを呼び出す必要があります。

##### <a name="parameters:"></a>パラメーター:

|名前| 型| 説明|
|---|---|---|
|`name`| String|削除するプロパティの名前。|

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](./tutorial-api-requirement-sets.md)| 1.0|
|[最小限のアクセス許可レベル](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用可能な Outlook のモード| 作成または読み取り|
####  <a name="saveasync([callback],-[asynccontext])"></a>saveAsync([callback], [asyncContext])

アイテム固有のカスタム プロパティをサーバーに保存します。

`saveAsync` メソッドを呼び出して、[`set`](CustomProperties.md#setname-value) メソッド、または `CustomProperties` オブジェクトの [`remove`](CustomProperties.md#removename) メソッドで行ったすべての変更を保持する必要があります。保存操作は非同期です。

コールバック機能を確認し、`saveAsync` からエラーを処理することをお勧めします。特に、ユーザーが表示フォームの接続状態時に、読み取り用のアドインがアクティブ化され、その後ユーザーが切断されます。切断状態でアドインが `saveAsync` を呼び出す場合、`saveAsync` はエラーを返します。コールバック メソッドは、このエラーを適切に処理する必要があります。

##### <a name="parameters:"></a>パラメーター:

|名前| 型| 属性| 説明|
|---|---|---|---|
|`callback`| function| &lt;optional&gt;|メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](simple-types.md#asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。 |
|`asyncContext`| Object| &lt;optional&gt;|コールバック メソッドに渡される状態データです。|

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](./tutorial-api-requirement-sets.md)| 1.0|
|[最小限のアクセス許可レベル](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用可能な Outlook のモード| 作成または読み取り|

##### <a name="example"></a>例

次の JavaScript コード サンプルは、`loadCustomPropertiesAsync` メソッドを非同期的に使用して、現在のアイテムに固有のカスタム プロパティをロードし、[`saveAsync`](CustomProperties.md#saveasynccallback-asynccontext) メソッドでこれらのプロパティをサーバーに保存する方法を示します。カスタム プロパティをロードした後、このコード サンプルでは [`get`](#getname--string) メソッドを使用してカスタム プロパティ `myProp` を読み取り、[`set`](CustomProperties.md#setname-value) メソッドでカスタム プロパティ `otherProp` を書き込み、最後に `saveAsync` メソッドを呼び出して、カスタム プロパティを保存します。

```
// The initialize function is required for all add-ins.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
    // After the DOM is loaded, add-in-specific code can run.
    var item = Office.context.mailbox.item;
    item.loadCustomPropertiesAsync(customPropsCallback);
  });
}
function customPropsCallback(asyncResult) {
  var customProps = asyncResult.value;
  var myProp = customProps.get("myProp");

  customProps.set("otherProp", "value");
  customProps.saveAsync(saveCallback);
}

function saveCallback(asyncResult) {
  if (asyncResult.status == Office.AsyncResultStatus.Failed){
    write(asyncResult.error.message);
  }
  else {
    // Async call to save custom properties completed.
    // Proceed to do the appropriate for your add-in.
  }
}

// Writes to a div with id='message' on the page.
function write(message){
  document.getElementById('message').innerText += message;
}
```

####  <a name="set(name,-value)"></a>set(name, value)

指定のプロパティを指定の値に設定します。

`set` メソッドは、指定のプロパティを指定の値に設定します。[`saveAsync`](CustomProperties.md#saveasynccallback-asynccontext) メソッドを使用して、プロパティをサーバーに保存する必要があります。

指定のプロパティが存在しない場合、`set` メソッドによって新しいプロパティが作成されます。存在する場合は、既存の値が新しい値に置き換えられます。`value` パラメーターには任意の型を使用できます。ただし、サーバーには常に文字列として渡されます。

##### <a name="parameters:"></a>パラメーター:

|名前| 型| 説明|
|---|---|---|
|`name`| String|設定するプロパティの名前。|
|`value`| Object|設定するプロパティの値。|

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](./tutorial-api-requirement-sets.md)| 1.0|
|[最小限のアクセス許可レベル](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用可能な Outlook のモード| 作成または読み取り|
