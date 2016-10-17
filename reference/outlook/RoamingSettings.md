

# <a name="roamingsettings"></a>RoamingSettings

`RoamingSettings` オブジェクトのメソッドを使用して作成された設定は、アドインごと、およびユーザーごとに保存されます。つまり、設定を作成したアドインでのみ利用できます。また、設定を保存したユーザーのメール ボックスからのみ利用できます。

> Outlook アドイン API では、それらの設定を作成したアドインのみが設定にアクセスできますが、これらの設定がセキュアなストレージであると見なすことはできません。これらの設定は、Exchange Web サービスや拡張 MAPI からアクセスできます。それらに、ユーザー資格情報やセキュリティ トークンなどの機密情報を格納しないでください。

設定の名前は String ですが、値は String、Number、Boolean、null、Object、Array のいずれかになります。

`RoamingSettings` オブジェクトは `Office.context` 名前空間の [`roamingSettings`](Office.context.md#roamingsettings-roamingsettings) プロパティ経由でアクセス可能です。

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](./tutorial-api-requirement-sets.md)| 1.0|
|[最小限のアクセス許可レベル](../../docs/outlook/understanding-outlook-add-in-permissions.md)| 制限あり|
|適用可能な Outlook のモード| 作成または読み取り|

### <a name="example"></a>例

```
// Get the current value of the 'myKey' setting
var value = Office.context.roamingSettings.get('myKey');
// Update the value of the 'myKey' setting
Office.context.roamingSettings.set('myKey', 'Hello World!');
// Persist the change
Office.context.roamingSettings.saveAsync();
```

### <a name="methods"></a>メソッド

####  <a name="get(name)-→-(nullable)-{string|number|boolean|object|array}"></a>get(name) → (nullable) {String|Number|Boolean|Object|Array}

指定された設定を取得します。

##### <a name="parameters:"></a>パラメーター:

|名前| 型| 説明|
|---|---|---|
|`name`| String|取得する設定の名前 (大文字と小文字を区別)。|

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](./tutorial-api-requirement-sets.md)| 1.0|
|[最小限のアクセス許可レベル](../../docs/outlook/understanding-outlook-add-in-permissions.md)| 制限あり|
|適用可能な Outlook のモード| 作成または読み取り|

##### <a name="returns:"></a>戻り値:

<dl class="param-type">

<dt>型</dt>

<dd>String | Number | Boolean | Object | Array</dd>

</dl>

####  <a name="remove(name)"></a>remove(name)

指定された設定を削除します。

##### <a name="parameters:"></a>パラメーター:

|名前| 型| 説明|
|---|---|---|
|`name`| String|削除する設定の名前 (大文字と小文字を区別)。|

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](./tutorial-api-requirement-sets.md)| 1.0|
|[最小限のアクセス許可レベル](../../docs/outlook/understanding-outlook-add-in-permissions.md)| 制限あり|
|適用可能な Outlook のモード| 作成または読み取り|
####  <a name="saveasync([callback])"></a>saveAsync([callback])

設定を保存します。

アドインによって以前保存された設定は、アプリの初期化時に読み込まれます。したがって、セッション実行中、[`set`](RoamingSettings.md#setname-value) および [`get`](RoamingSettings.md#getname--nullable-stringnumberbooleanobjectarray) メソッドを使用し、設定プロパティ バッグのメモリ内のコピーと共に使用できます。これらの設定をアドインの次回使用時にも使用できるように保存するときは、`saveAsync` メソッドを使用します。

##### <a name="parameters:"></a>パラメーター:

|名前| 型| 属性| 説明|
|---|---|---|---|
|`callback`| function| &lt;optional&gt;|メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](simple-types.md#asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。 |

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](./tutorial-api-requirement-sets.md)| 1.0|
|[最小限のアクセス許可レベル](../../docs/outlook/understanding-outlook-add-in-permissions.md)| 制限あり|
|適用可能な Outlook のモード| 作成または読み取り|
####  <a name="set(name,-value)"></a>set(name, value)

指定された設定を設定または作成します。

set メソッドは、指定された名前の新しい設定を作成するか (その設定が存在しない場合)、指定された名前の既存の設定を設定します。値は、そのデータ型のシリアル化された JSON 表現としてドキュメントに格納されます。

各アドインの設定に最大 2 MB を使用でき、個々の設定は、32 KB に制限されています。

`set` 関数を使用し設定に加えられた変更は、[`saveAsync`](RoamingSettings.md#saveasynccallback) 関数が呼び出されるまでサーバーに保存されません。

##### <a name="parameters:"></a>パラメーター:

|名前| 型| 説明|
|---|---|---|
|`name`| String|設定または作成する設定の名前 (大文字と小文字を区別します)。|
|`value`| String、Number、Boolean、Object、Array|格納する値。|

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](./tutorial-api-requirement-sets.md)| 1.0|
|[最小限のアクセス許可レベル](../../docs/outlook/understanding-outlook-add-in-permissions.md)| 制限あり|
|適用可能な Outlook のモード| 作成または読み取り|
