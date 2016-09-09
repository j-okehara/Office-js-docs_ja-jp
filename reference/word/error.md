# OfficeExtension.Error オブジェクト (JavaScript API for Word)

Word JavaScript API の使用時に発生するエラーを表します。

_適用対象:Word 2016、Word for iPad、Word for Mac_

## プロパティ
| プロパティ     | 型   |説明
|:---------------|:--------|:----------|
|code|string|エラーの種類を示す値を取得します。 次の値をとることができます。"AccessDenied"、"GeneralException"、"ActivityLimitReached"、"InvalidArgument"、"ItemNotFound"、"NotImplemented"。 <!-- Values come from OfficeExtension.Error and Word.ErrorCodes. -->|
|debugInfo|string|エラーが発生したときに何が起こったかを示す値を取得します。この値は、開発中またはデバッグ中のみに使用することが想定されています。  |
|message |string| エラー コードに対応する、人間が判読できるローカライズされた文字列を取得します。|
|name |string| 常に "OfficeExtension.Error" である値を取得します。 |
|traceMessages |string[]| Context.trace(); を使用して設定するインストルメンテーション メッセージに対応する値の配列を取得します。 |

_プロパティのアクセスの[例](#例)を参照してください。_

## メソッド

| メソッド           | 戻り値の型    |説明|
|:---------------|:--------|:----------|
|[toString()](#tostring)|string|次の形式でエラー コードとメッセージの値を返します: "{0}: {1}", コード, メッセージ。|

## メソッドの詳細

### toString()
次の形式でエラー コードとメッセージの値を返します: "{0}: {1}", コード, メッセージ。

#### 構文
```js
error.toString()
```

#### パラメーター
なし。

#### 戻り値
string

#### 例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to insert text in to the beginning of the body.
    // This will cause an OfficeExtension.Error.
    body.insertText(0);

    // Synchronize the document state by executing the queued-up commands,
    // and return a promise to indicate task completion.
    return context.sync();
})
.catch(function (error) {
    if (error instanceof OfficeExtension.Error) {
        console.log('Error code and message: ' + error.toString());
    }
});

```

## プロパティのアクセスの例

### トレース メッセージのインストルメンテーション

次の例は、エラーが発生した場所を判別するコマンドのバッチをインストルメントする方法を示しています。最初のバッチは最初の 2 つの段落を正常に挿入し、エラーは発生していません。2 番目のバッチは 3 番目と 4 番目の段落を正常に挿入しましたが、5 番目の段落を挿入する呼び出しに失敗しました。5 番目のトレース メッセージを追加するコマンドを含め、バッチ内で失敗したコマンドより後の他のすべてのコマンドは実行されていません。この例では、4 番目の段落が挿入された後、5 番目のトレース メッセージを追加する前にエラーが発生しました。

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to insert the paragraph at the end of the document body.
    // Start a batch of commands.
    body.insertParagraph('1st paragraph', Word.InsertLocation.end);
    // Queue a command for instrumenting this part of the batch.
    context.trace('1st paragraph successful');

    body.insertParagraph('2nd paragraph', Word.InsertLocation.end);
    context.trace('2nd paragraph successful');

    // Synchronize the document state by executing the queued-up commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        // Queue a commmand to insert the paragraph at the end of the document body.
        // Start a new batch of commands.
        body.insertParagraph('3rd paragraph', Word.InsertLocation.end);
        context.trace('3rd paragraph successful');

        body.insertParagraph('4th paragraph', Word.InsertLocation.end);
        context.trace('4th paragraph successful');

        // This command will cause an error. The trace messages in the queue up to
        // this point will be available via Error.traceMessages.
        body.insertParagraph(0, '5th paragraph', Word.InsertLocation.end);
        // Queue a command for instrumenting this part of the batch.
        // This trace message will not be set on Error.traceMessages.
        context.trace('5th paragraph successful');
    }).then(context.sync);
})
.catch(function (error) {
    if (error instanceof OfficeExtension.Error) {
        console.log('Trace messages: ' + error.traceMessages);
    }
});

// Output: "Trace messages: 3rd paragraph successful,4th paragraph successful"

```
