# <a name="requestcontext-object-(javascript-api-for-word)"></a>RequestContext オブジェクト (JavaScript API for Word)

RequestContext オブジェクトは、2 つのアプリケーションが別のプロセスで実行されているときの、Word アドインから Word への要求を容易にします。

_適用対象:Word 2016、Word for iPad、Word for Mac、Word Online_

## <a name="properties"></a>プロパティ
なし

## <a name="methods"></a>メソッド

| メソッド         | 戻り値の型    |説明|
|:---------------|:--------|:----------|
|[load(object: object, option: object)](#loadobject-object-option-object)  |void     |JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオプションを設定します。|
|[sync()](#sync)  |Promise オブジェクト |要求キューを Word に送信し、さらに多くの操作を連続的に繋ぐために使用できる約束オブジェクトを返します。|

## <a name="method-details"></a>メソッドの詳細

### <a name="load(object:-object,-option:-object)"></a>load(object: object, option: object)
JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオプションを設定します。

#### <a name="syntax"></a>構文
```js
requestContextObject.load(object, loadOption);
```

#### <a name="parameters"></a>パラメーター
| パラメーター       | 型    |説明|
|:----------------|:--------|:----------|
|object|object|省略可能。読み込むオブジェクトの名前を指定します。|
|option|[loadOption](loadoption.md)|省略可能ですが、このオプションの使用をお勧めします。select、expand、skip、top などの読み込みオプションを指定します。 |

#### <a name="returns"></a>戻り値
void

##### <a name="examples"></a>例

次の例は、要求コンテキストを使用して、text プロパティを段落コレクションに読み込む方法を示しています。

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;

    // Queue a commmand to load the text property for all of the paragraphs.
    context.load(paragraphs, 'text');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {

        // Queue a a set of commands to get the HTML of the first paragraph.
        var html = paragraphs.items[0].getHtml();

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Paragraph HTML: ' + html.value);
        });
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});

```

#### <a name="additional-information"></a>その他の情報

追跡対象のオブジェクトを追加した後は、load() を呼び出す必要があります。

### <a name="sync()"></a>sync()
要求キューを Word に送信し、さらに多くの操作を連続的に繋ぐために使用できる約束オブジェクトを返します。

#### <a name="syntax"></a>構文
```js
requestContextObject.sync();
```

#### <a name="parameters"></a>パラメーター
なし

#### <a name="returns"></a>戻り値
Promise オブジェクト。

#### <a name="examples"></a>例

次の例は、2 回使用されている sync メソッドを示しています。1) コンテンツ コントロールのコレクションに、それぞれのコンテンツ コントロールの text プロパティを読み込み、2) コレクション内の最初のコンテンツ コントロールの内容をクリアします。

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the content controls collection.
    var contentControls = context.document.contentControls;

    // Queue a command to load the content controls collection.
    contentControls.load('text');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {

        if (contentControls.items.length === 0) {
            console.log("There isn't a content control in this document.");
        } else {

            // Queue a command to clear the contents of the first content control.
            contentControls.items[0].clear();
            // Synchronize the document state by executing the queued commands,
            // and return a promise to indicate task completion.
            return context.sync().then(function () {
                console.log('Content control cleared of contents.');
            });
        }

    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});

```

## <a name="support-details"></a>サポートの詳細
実行時のチェックで[要件セット](../office-add-in-requirement-sets.md)を使用して、アプリケーションが Word のホスト バージョンによってサポートされていることを確かめます。Office ホスト アプリケーションとサーバーの要件の詳細については、「[Office アドインを実行するための要件](../../docs/overview/requirements-for-running-office-add-ins.md)」を参照してください。
