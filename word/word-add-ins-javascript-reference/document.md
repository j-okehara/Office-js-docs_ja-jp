# Document オブジェクト (JavaScript API for Word)

Document オブジェクトは、最上位レベルのオブジェクトです。ドキュメント オブジェクトには、1 つ以上のセクション、コンテンツ コントロール、ドキュメントの内容を含む本文が含まれています。

_適用対象:Word 2016、Word for iPad、Word for Mac_

## プロパティ
| プロパティ   | 型|説明
|:---------------|:--------|:----------|
|Saved|bool|ドキュメント内の変更が保存されているかどうかを示します。値 true は、ドキュメントが保存されてから変更されていないことを示します。読み取り専用です。|

## 関係
| リレーションシップ | 型|説明|
|:---------------|:--------|:----------|
|body|[Body](body.md)|ドキュメントの本文を取得します。本文は、ヘッダー、フッター、脚注、テキストボックスなどを除いたテキストです。読み取り専用です。|
|contentControls|[ContentControlCollection](contentcontrolcollection.md)|現在のドキュメントにあるコンテンツ コントロール オブジェクトのコレクションを取得します。これには、ドキュメントの本文、ヘッダー、フッター、テキストボックスなどにあるコンテンツ コントロールが含まれます.読み取り専用です。|
|sections|[SectionCollection](sectioncollection.md)|ドキュメントにあるセクション オブジェクトのコレクションを取得します。読み取り専用です。|

## メソッド

| メソッド   | 戻り値の型|説明|
|:---------------|:--------|:----------|
|[getSelection()](#getselection)|[Range](range.md)|ドキュメントの現在の選択範囲を取得します。複数選択はサポートされていません。|
|[load(param: object)](#loadparam-object)|void|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|
|[save()](#save)|void|ドキュメントを保存します。ここでは、ドキュメントが保存されたことがない場合は、Word の既定のファイルの名前付け規則を使用します。|

## メソッドの詳細

### getSelection()
ドキュメントの現在の選択範囲を取得します。複数選択はサポートされていません。

#### 構文
```js
documentObject.getSelection();
```

#### パラメーター
なし

#### 戻り値
[Range](range.md)

#### 例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    var textSample = 'This is an example of the insert text method. This is a method ' + 
        'which allows users to insert text into a selection. It can insert text into a ' +
        'relative location or it can overwrite the current selection. Since the ' +
        'getSelection method returns a range object, look up the range object documentation ' +
        'for everything you can do with a selection.';
    
    // Create a range proxy object for the current selection.
    var range = context.document.getSelection();
    
    // Queue a commmand to insert text at the end of the selection.
    range.insertText(textSample, Word.InsertLocation.end);
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Inserted the text at the end of the selection.');
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### load(param: object)
JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。

#### 構文
```js
object.load(param);
```

#### パラメータ
| パラメーター   | 型|説明|
|:---------------|:--------|:----------|
|param|object|省略可能。パラメーターとリレーションシップ名を、区切られた文字列または 1 つの配列として受け入れます。あるいは、[loadOption](loadoption.md) オブジェクトを提供します。|

#### 戻り値
(非推奨)

#### 例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the document.
    var thisDocument = context.document;
    
    // Queue a command to load content control properties.
    context.load(thisDocument, 'contentControls/id, contentControls/text, contentControls/tag');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        if (thisDocument.contentControls.items.length !== 0) {
            for (var i = 0; i < thisDocument.contentControls.items.length; i++) {
                console.log(thisDocument.contentControls.items[i].id);
                console.log(thisDocument.contentControls.items[i].text);
                console.log(thisDocument.contentControls.items[i].tag);
            }
        } else {
            console.log('No content controls in this document.');
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

### save()
ドキュメントを保存します。ここでは、ドキュメントが保存されたことがない場合は、Word の既定のファイルの名前付け規則を使用します。

#### 構文
```js
documentObject.save();
```

#### パラメーター
なし

#### 戻り値
(非推奨)

#### 例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the document.
    var thisDocument = context.document;

    // Queue a commmand to load the document save state (on the saved property).
    context.load(thisDocument, 'saved');    
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        
        if (thisDocument.saved === false) {
            // Queue a command to save this document.
            thisDocument.save();
            
            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync().then(function () {
                console.log('Saved the document');
            });
        } else {
            console.log('The document has not changed since the last save.');
        }
    });  
})
.catch(function (error) {
    console.log("Error: " + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

## サポートの詳細

実行時のチェックで[要件セット](https://msdn.microsoft.com/EN-US/library/office/mt590206.aspx)を使用して、アプリケーションが Word のホスト バージョンによってサポートされていることを確かめます。Office ホスト アプリケーションとサーバーの要件の詳細については、「[Office アドインを実行するための要件](https://msdn.microsoft.com/EN-US/library/office/dn833104.aspx)」を参照してください。 
