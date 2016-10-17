# <a name="inlinepicture-object-(javascript-api-for-word)"></a>InlinePicture オブジェクト (JavaScript API for Word)

インライン画像を表します。

_適用対象:Word 2016、Word for iPad、Word for Mac、Word Online_

## <a name="properties"></a>プロパティ
| プロパティ     | 型   |説明
|:---------------|:--------|:----------|
|altTextDescription|string|インライン画像に関連付けられた代替テキストを表す文字列を取得または設定します|
|altTextTitle|string|インライン画像のタイトルを含む文字列を取得または設定します。|
|hyperlink|文字列|インライン画像に関連付けられているハイパーリンクを取得または設定します。|
|lockAspectRatio|bool|インライン画像のサイズを変更する際にその元の縦横比を保持するかどうかを示す値を取得または設定します。|

## <a name="relationships"></a>リレーションシップ
| リレーションシップ | 型   |説明|
|:---------------|:--------|:----------|
|height|**float**|インライン画像の高さを表す数値を取得または設定します。これは、ポイント単位で測定されます。 |
|parentContentControl|[ContentControl](contentcontrol.md)|インライン画像を含むコンテンツ コントロールを取得します。親コンテンツ コントロールがない場合は null を返します。読み取り専用です。|
|paragraph|[paragraph](paragraph.md)|インライン画像を含む段落を取得します。読み取り専用です。
|width|**float**|インライン画像の幅を表す数値を取得または設定します。これは、ポイント単位で測定されます。|

## <a name="methods"></a>メソッド

| メソッド           | 戻り値の型    |説明|
|:---------------|:--------|:----------|
|[delete()](#delete)|void|ドキュメントから画像を削除します。|
|[getBase64ImageSrc()](#getbase64imagesrc)|object|値がインライン画像の base64 エンコード文字列表記であるオブジェクトを取得します。|
|[insertBreak(breakType: BreakType, insertLocation: InsertLocation)](#insertbreakbreaktype-breaktype-insertlocation-insertlocation)|void|指定した位置に、区切りを挿入します。有効な insertLocation の値は、'Before' または 'After' です。|
|[insertContentControl()](#insertcontentcontrol)|[ContentControl](contentcontrol.md)|リッチ テキスト コンテンツ コントロールでインライン画像をラップします。|
|[insertFileFromBase64(base64File: string, insertLocation:InsertLocation)](#insertfilefrombase64base64file-string-insertlocation-insertlocation)|[Range](range.md)|文書を本文の指定された位置に挿入します。有効な insertLocation の値は、'Before' または 'After' です。|
|[insertHtml(html: string, insertLocation:InsertLocation)](#inserthtmlhtml-string-insertlocation-insertlocation)|[Range](range.md)|指定した位置に HTML を挿入します。有効な insertLocation の値は、'Before' または 'After' です。|
|[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation:InsertLocation)](#insertInlinePictureFromBase64base64EncodedImage-string-insertlocation-insertlocation)|[InlinePicture](inlinepicture.md)|画像を本文の指定された位置に挿入します。有効な insertLocation の値は、'Replace'、'Before' または 'After' です。 |
|[insertOoxml(ooxml: string, insertLocation:InsertLocation)](#insertooxmlooxml-string-insertlocation-insertlocation)|[Range](range.md)|指定した位置に OOXML を挿入します。有効な insertLocation の値は、'Before' または 'After' です。|
|[insertParagraph(paragraphText: string, insertLocation:InsertLocation)](#insertparagraphparagraphtext-string-insertlocation-insertlocation)|[Paragraph](paragraph.md)|指定した位置に、段落を挿入します。有効な insertLocation の値は、'Before' または 'After' です。|
|[insertText(text: string, insertLocation:InsertLocation)](#inserttexttext-string-insertlocation-insertlocation)|[Range](range.md)|テキストを本文の指定された位置に挿入します。有効な insertLocation の値は、'Before' または 'After' です。|
|[select(selectionMode: SelectionMode)](#selectselectionmode-selectionmode)|void|画像を選択し、その本文に Word の UI を移動します。selectionMode 値は、'Select'、'Start'、'End' のいずれかになります。|
|[load(param: object)](#loadparam-object)|(非推奨)|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|

## <a name="method-details"></a>メソッドの詳細

### <a name="delete()"></a>delete()
ドキュメントから画像を削除します。

#### <a name="syntax"></a>構文
```js
inlinePictureObject.delete();
```

#### <a name="parameters"></a>パラメーター
なし

#### <a name="returns"></a>戻り値
void

### <a name="getbase64imagesrc()"></a>getBase64ImageSrc()
値がインライン画像の base64 エンコード文字列表記であるオブジェクトを取得します。

#### <a name="syntax"></a>構文
```js
var base64 = inlinePictureObject.getBase64ImageSrc();
return context.sync().then(function () {    
    console.log("base64 string is " + base64.value);
});

```

#### <a name="parameters"></a>パラメーター
なし

#### <a name="returns"></a>戻り値
オブジェクト 



### <a name="insertbreak(breaktype:-breaktype,-insertlocation:-insertlocation)"></a>insertBreak(breakType: BreakType, insertLocation: InsertLocation)

#### <a name="syntax"></a>構文
```js
inlinePictureObject.insertBreak(breakType, insertLocation);
```
#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|breakType|BreakType|必須。本文に追加する区切りの種類。|
|insertLocation|InsertLocation|必須。有効な値は、'Before' または 'After' です。|

#### <a name="returns"></a>戻り値
void

### <a name="insertcontentcontrol()"></a>insertContentControl()
リッチ テキストのコンテンツ コントロールでインライン画像をラップします。

#### <a name="syntax"></a>構文
```js
inlinePictureObject.insertContentControl();
```

#### <a name="parameters"></a>パラメーター
なし

#### <a name="returns"></a>戻り値
[ContentControl](contentcontrol.md)

### <a name="insertfilefrombase64(base64file:-string,-insertlocation:-insertlocation)"></a>insertFileFromBase64(base64File: string, insertLocation:InsertLocation)
文書を本文の指定された位置に挿入します。insertLocation の値は、'Before' か 'After' になります。

#### <a name="syntax"></a>構文
```js
inlinePictureObject.insertFileFromBase64(base64File, insertLocation);
```
#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|base64File|string|必須。docx ファイルの base64 でエンコードされたコンテンツ。|
|insertLocation|InsertLocation|必須。有効な値は、'Before' または 'After' です。|

#### <a name="returns"></a>戻り値
[Range](range.md)

### <a name="inserthtml(html:-string,-insertlocation:-insertlocation)"></a>insertHtml(html: string, insertLocation:InsertLocation)
指定した位置にHTML を挿入します。insertLocation の値は、'Before' か 'After' になります。

#### <a name="syntax"></a>構文
```js
inlinePictureObject.insertHtml(html, insertLocation);
```

#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|Html|string|必須。文書に挿入する HTML。|
|insertLocation|InsertLocation|必須。有効な値は、'Before' または 'After' です。|

#### <a name="returns"></a>戻り値
[Range](range.md)


### <a name="insertinlinepicturefrombase64(base64encodedimage:-string,-insertlocation:-insertlocation)"></a>insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation:InsertLocation)
画像を本文の指定された位置に挿入します。insertLocation の値は、'Before' か 'After' になります。

#### <a name="syntax"></a>構文
inlinePictureObject.insertInlinePictureFromBase64(image, insertLocation);

#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|base64EncodedImage|string|必須。本文に挿入される base64 でエンコードされた画像。|
|insertLocation|InsertLocation|必須。有効な値は、'Before' または 'After' です。|

#### <a name="returns"></a>戻り値
[InlinePicture](inlinepicture.md)


### <a name="insertooxml(ooxml:-string,-insertlocation:-insertlocation)"></a>insertOoxml(ooxml: string, insertLocation:InsertLocation)
指定した位置に、OOXML を挿入します。insertLocation の値は、'Before' か 'After' になります。

#### <a name="syntax"></a>構文
```js
inlinePictureObject.insertOoxml(ooxml, insertLocation);
```

#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|ooxml|string|必須。挿入する OOXML を指定します。|
|insertLocation|InsertLocation|必須。有効な値は、'Before' または 'After' です。|

#### <a name="returns"></a>戻り値
[Range](range.md)

### <a name="insertparagraph(paragraphtext:-string,-insertlocation:-insertlocation)"></a>insertParagraph(paragraphText: string, insertLocation:InsertLocation)
指定した位置に、段落を挿入します。有効な insertLocation の値は、'Before' または 'After' です。

#### <a name="syntax"></a>構文
```js
inlinePictureObject.insertParagraph(paragraphText, insertLocation);
```

#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|paragraphText|string|必須。挿入する段落テキスト。|
|insertLocation|InsertLocation|必須。有効な値は、'Before' または 'After' です。|

#### <a name="returns"></a>戻り値
[Paragraph](paragraph.md)

### <a name="inserttext(text:-string,-insertlocation:-insertlocation)"></a>insertText(text: string, insertLocation:InsertLocation)
テキストを本文の指定された位置に挿入します。insertLocation の値は、'Before' か 'After' になります。

#### <a name="syntax"></a>構文
```js
inlinePictureObject.insertText(text, insertLocation);
```

#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|text|string|必須。挿入するテキスト。|
|insertLocation|InsertLocation|必須。有効な値は、'Before' または 'After' です。|

#### <a name="returns"></a>戻り値
[Range](range.md)

### <a name="select(selectionmode:-selectionmode)"></a>select(selectionMode: SelectionMode)
画像を選択し、その本文に Word の UI を移動します。selectionMode 値は、'Select'、'Start'、'End' のいずれかになります。

#### <a name="syntax"></a>構文
```js
inlinePictureObject.select(selectionMode);
```

#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|selectionMode|SelectionMode|省略可能。選択モードは、'Select'、'Start'、'End' のいずれかになります。'Select' が既定値です。|

#### <a name="returns"></a>戻り値
void

### <a name="load(param:-object)"></a>load(param: object)
JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。

#### <a name="syntax"></a>構文
```js
object.load(param);
```

#### <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|param|object|省略可能。パラメーターとリレーションシップ名を、区切られた文字列または 1 つの配列として受け入れます。あるいは、[loadOption](loadoption.md) オブジェクトを提供します。|

#### <a name="returns"></a>戻り値
void

## <a name="support-details"></a>サポートの詳細
実行時のチェックで[要件セット](../office-add-in-requirement-sets.md)を使用して、アプリケーションが Word のホスト バージョンによってサポートされていることを確かめます。Office ホスト アプリケーションとサーバーの要件の詳細については、「[Office アドインを実行するための要件](../../docs/overview/requirements-for-running-office-add-ins.md)」を参照してください。
