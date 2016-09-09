# SectionCollection オブジェクト (JavaScript API for Word)

ドキュメントの[section](section.md) オブジェクトのコレクションが含まれています。

_適用対象:Word 2016、Word for iPad、Word for Mac_

## プロパティ
| プロパティ     | 型   |説明
|:---------------|:--------|:----------|
|Items|[Section[]](section.md)|セクション オブジェクトのコレクション。読み取り専用です。|

## リレーションシップ
なし


## メソッド

| メソッド           | 戻り値の型    |説明|
|:---------------|:--------|:----------|
|[load(param: object)](#loadparam-object)|void|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|

## メソッドの詳細

### load(param: object)
JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。

#### 構文
```js
object.load(param);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|param|object|省略可能。パラメーターとリレーションシップ名を、区切られた文字列または 1 つの配列として受け入れます。あるいは、[loadOption](loadoption.md) オブジェクトを提供します。|

#### 戻り値
void

#### 例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy sectionsCollection object.
    var mySections = context.document.sections;

    // Queue a commmand to load the sections.
    context.load(mySections, 'body/style');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {

        // Create a proxy object the primary header of the first section.
        // Note that the header is a body object.
        var myHeader = mySections.items[0].getHeader("primary");

        // Queue a command to insert text at the end of the header.
        myHeader.insertText("This is a header.", Word.InsertLocation.end);

        // Queue a command to wrap the header in a content control.
        myHeader.insertContentControl();

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log("Added a header to the first section.");
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

## サポートの詳細
実行時のチェックで[要件セット](../office-add-in-requirement-sets.md)を使用して、アプリケーションが Word のホスト バージョンによってサポートされていることを確かめます。Office ホスト アプリケーションとサーバーの要件の詳細については、「[Office アドインを実行するための要件](../../docs/overview/requirements-for-running-office-add-ins.md)」を参照してください。