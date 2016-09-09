# フォント オブジェクト (JavaScript API for Word)

フォントを表します。

_適用対象:Word 2016、Word for iPad、Word for Mac_

## プロパティ
| プロパティ     | 型   |説明
|:---------------|:--------|:----------|
|bold|bool|フォントが太字かどうかを示す値を取得または設定します。フォントの書式設定が太字の場合は true、それ以外の場合は false です。|
|color|string|指定されたフォントの色を取得または設定します。値を "#RRGGBB" 形式または色の名前で指定できます。|
|doubleStrikeThrough|bool|フォントが二重取り消し線付きかどうかを示す値を取得または設定します。フォントの書式が二重取り消し線付きのテキストである場合は true、それ以外の場合は false です。|
|highlightColor|string|指定したフォントの強調表示色を取得または設定します。値を "#RRGGBB" 形式または色の名前で指定できます。|
|italic|bool|フォントが斜体かどうかを示す値を取得または設定します。フォントが斜体の場合は true、それ以外の場合は false です。|
|name|string|フォント名を表す値を取得または設定します。|
|strikeThrough|bool|フォントが取り消し線付きかどうかを示す値を取得または設定します。フォントの書式が取り消し線付きのテキストである場合は true、それ以外の場合は false です。|
|subscript|bool|フォントが下付き文字かどうかを示す値を取得または設定します。フォントの書式が下付き文字である場合は true、それ以外の場合は false です。|
|superscript|bool|フォントが上付き文字かどうかを示す値を取得または設定します。フォントの書式が上付き文字である場合は true、それ以外の場合は false です。|

_プロパティのアクセスの[例](#例)を参照してください。_

## リレーションシップ
| リレーションシップ | 型   |説明|
|:---------------|:--------|:----------|
|size|**浮動小数点数**|フォント サイズをポイント単位で表す値を取得または設定します。|
|underline|**string**|フォントの下線の種類を示す値を取得または設定します。有効な値は次のとおりです。"None"、"Single"、"Word"、"Double"、"Dotted"、"Hidden"、"Thick"、"Dashline"、"Dotline"、"DotDashLine"、"TwoDotDashLine"、"Wave"|

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

    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;

    // Queue a commmand to load the font property for all of the paragraphs.
    context.load(paragraphs, 'font');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {

        // Create a proxy object for the font object on the first paragraph in the collection.
        var font = paragraphs.items[0].font;

        // Queue a set of property value changes on the font proxy object.
        font.size = 32;
        font.bold = true;
        font.color = '#0000ff';
        font.highlightColor = '#ffff00';

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('The font has changed.');
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

## プロパティのアクセスの例

### フォント名の変更
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a range proxy object for the current selection.
    var selection = context.document.getSelection();

    // Queue a commmand to change the current selection's font name.
    selection.font.name = 'Arial';

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('The font name has changed.');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### フォントの色の変更
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a range proxy object for the current selection.
    var selection = context.document.getSelection();

    // Queue a commmand to change the font color of the current selection.
    selection.font.color = 'blue';

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('The font color of the selection has been changed.');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### フォント サイズの変更
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a range proxy object for the current selection.
    var selection = context.document.getSelection();

    // Queue a commmand to change the current selection's font size.
    selection.font.size = 20;

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('The font size has changed.');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### 選択したテキストの強調表示
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a range proxy object for the current selection.
    var selection = context.document.getSelection();

    // Queue a commmand to highlight the current selection.
    selection.font.highlightColor = '#FFFF00'; // Yellow

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('The selection has been highlighted.');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### 太字のテキスト
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a range proxy object for the current selection.
    var selection = context.document.getSelection();

    // Queue a commmand to make the current selection bold.
    selection.font.bold = true;

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('The selection is now bold.');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});

```

### 下線付きのテキスト
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a range proxy object for the current selection.
    var selection = context.document.getSelection();

    // Queue a commmand to underline the current selection.
    selection.font.underline = Word.UnderlineType.single;

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('The selection now has an underline style.');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### 取り消し線付きのテキスト
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a range proxy object for the current selection.
    var selection = context.document.getSelection();

    // Queue a commmand to strikethrough the font of the current selection.
    selection.font.strikeThrough = true;

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('The selection now has a strikethrough.');
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
