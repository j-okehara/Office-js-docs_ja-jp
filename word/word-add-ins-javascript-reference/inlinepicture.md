# InlinePicture オブジェクト (JavaScript API for Word)

インライン画像を表します。

_適用対象:Word 2016、Word for iPad、Word for Mac_

## プロパティ
| プロパティ   | 型|説明
|:---------------|:--------|:----------|
|altTextDescription|string|インライン画像に関連付けられた代替テキストを表す文字列を取得または設定します|
|altTextTitle|string|インライン画像のタイトルを含む文字列を取得または設定します。|
|hyperlink|string|インライン画像に関連付けられているハイパーリンクを取得または設定します。|
|lockAspectRatio|bool|インライン画像のサイズを変更する際にその元の縦横比を保持するかどうかを示す値を取得または設定します。|


_プロパティのアクセスの[例](#property-access-examples)を参照してください。_

## 関係
| リレーションシップ | 型|説明|
|:---------------|:--------|:----------|
|height|**float**|インライン画像の高さを表す数値を取得するか設定します。これはポイント単位で測定されます。 |
|parentContentControl|[ContentControl](contentcontrol.md)|インライン画像を含むコンテンツ コントロールを取得します。親コンテンツ コントロールがない場合は null を返します。読み取り専用です。|
|Paragraph|[paragraph](paragraph.md)|インライン画像を含む段落を取得します。読み取り専用です。
|width|**float**|インライン画像の幅を表す数値を取得するか設定します。これはポイント単位で測定されます。|

## メソッド

| メソッド   | 戻り値の型|説明|
|:---------------|:--------|:----------|
|[delete()](#delete)|void|ドキュメントから画像を削除します。|
|[getBase64ImageSrc()](#getbase64imagesrc)|string|値がインライン画像の base64 エンコード文字列表記であるオブジェクトを取得します。|
|[insertBreak(breakType: BreakType, insertLocation: InsertLocation)](#insertbreakbreaktype-breaktype-insertlocation-insertlocation)|void|指定した位置に改行を挿入します。insertLocation の値は、'Before' か 'After' になります。|
|[insertContentControl()](#insertcontentcontrol)|[ContentControl](contentcontrol.md)|リッチ テキストのコンテンツ コントロールでインライン画像をラップします。|
|[insertFileFromBase64(base64File: string, insertLocation:InsertLocation)](#insertfilefrombase64base64file-string-insertlocation-insertlocation)|[Range](range.md)|文書を本文の指定された位置に挿入します。insertLocation の値は、'Before' か 'After' になります。|
|[insertHtml(html: string, insertLocation: InsertLocation)](#inserthtmlhtml-string-insertlocation-insertlocation)|[Range](range.md)|指定した位置にHTML を挿入します。insertLocation の値は、'Before' か 'After' になります。|
|[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: InsertLocation)](#insertInlinePictureFromBase64base64EncodedImage-string-insertlocation-insertlocation)|[InlinePicture](inlinepicture.md)|画像を本文の指定された位置に挿入します。insertLocation の値は、'Replace'、'Before'、'After' のいずれかになります。 |
|[insertOoxml(ooxml: string, insertLocation: InsertLocation)](#insertooxmlooxml-string-insertlocation-insertlocation)|[Range](range.md)|指定した位置に、OOXML を挿入します。insertLocation の値は、'Before' か 'After' になります。|
|[insertParagraph(paragraphText: string, insertLocation: InsertLocation)](#insertparagraphparagraphtext-string-insertlocation-insertlocation)|[Paragraph](paragraph.md)|指定した位置に、段落を挿入します。有効な insertLocation の値は、'Before' または 'After' です。|
|[insertText(text: string, insertLocation:InsertLocation)](#inserttexttext-string-insertlocation-insertlocation)|[Range](range.md)|テキストを本文の指定された位置に挿入します。insertLocation の値は、'Before' か 'After' になります。|
|[select(selectionMode: SelectionMode)](#selectselectionmode-selectionmode)|void|画像を選択し、その本文に Word の UI を移動します。selectionMode 値は、'Select'、'Start'、'End' のいずれかになります。|
|[load(param: object)](#loadparam-object)|void|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|

## メソッドの詳細

### delete()
ドキュメントから画像を削除します。

#### 構文
```js
inlinePictureObject.delete();
```

#### パラメーター
なし

#### 戻り値
void

### getBase64ImageSrc()
値がインライン画像の base64 エンコード文字列表記であるオブジェクトを取得します。

#### 構文
```js
inlinePictureObject.getBase64ImageSrc();
```

#### パラメーター
なし

#### 戻り値
string

### insertBreak(breakType: BreakType, insertLocation: InsertLocation)

#### 構文
```js
inlinePictureObject.insertBreak(breakType, insertLocation);
```
#### パラメータ
| パラメーター   | 型|説明|
|:---------------|:--------|:----------|
|breakType|BreakType|必須。本文に追加する区切りの種類。|
|insertLocation|InsertLocation|必須。有効な値は、'Before' または 'After' です。|

#### 戻り値
void

### insertContentControl()
リッチ テキストのコンテンツ コントロールでインライン画像をラップします。

#### 構文
```js
inlinePictureObject.insertContentControl();
```

#### パラメーター
なし

#### 戻り値
[ContentControl](contentcontrol.md)

### insertFileFromBase64(base64File: string, insertLocation: InsertLocation)
文書を本文の指定された位置に挿入します。insertLocation の値は、'Before' か 'After' になります。

#### 構文
```js
inlinePictureObject.insertFileFromBase64(base64File, insertLocation);
```
#### パラメータ
| パラメーター   | 型|説明|
|:---------------|:--------|:----------|
|base64File|string|必須。docx ファイルの base64 でエンコードされたコンテンツ。|
|insertLocation|InsertLocation|必須。有効な値は、'Before' または 'After' です。|

#### 戻り値
[Range](range.md)

### insertHtml(html: string, insertLocation:InsertLocation)
指定した位置にHTML を挿入します。insertLocation の値は、'Before' か 'After' になります。

#### 構文
```js
inlinePictureObject.insertHtml(html, insertLocation);
```

#### パラメータ
| パラメーター   | 型|説明|
|:---------------|:--------|:----------|
|Html|string|必須。文書に挿入する HTML。|
|insertLocation|InsertLocation|必須。有効な値は、'Before' または 'After' です。|

#### 戻り値
[Range](range.md)


### insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: InsertLocation)
画像を本文の指定された位置に挿入します。insertLocation の値は、'Before' か 'After' になります。

#### 構文
inlinePictureObject.insertInlinePictureFromBase64(image, insertLocation);

#### パラメーター
| パラメーター   | 型|説明|
|:---------------|:--------|:----------|
|base64EncodedImage|string|必須。本文に挿入される base64 でエンコードされた画像。|
|insertLocation|InsertLocation|必須。有効な値は、'Before' または 'After' です。|

#### 戻り値
[InlinePicture](inlinepicture.md)


### insertOoxml(ooxml: string, insertLocation: InsertLocation)
指定した位置に、OOXML を挿入します。insertLocation の値は、'Before' か 'After' になります。

#### 構文
```js
inlinePictureObject.insertOoxml(ooxml, insertLocation);
```

#### パラメータ
| パラメーター   | 型|説明|
|:---------------|:--------|:----------|
|ooxml|string|必須。挿入する OOXML を指定します。|
|insertLocation|InsertLocation|必須。有効な値は、'Before' または 'After' です。|

#### 戻り値
[Range](range.md)

### insertParagraph(paragraphText: string, insertLocation:InsertLocation)
指定した位置に、段落を挿入します。有効な insertLocation の値は、'Before' または 'After' です。

#### 構文
```js
inlinePictureObject.insertParagraph(paragraphText, insertLocation);
```

#### パラメータ
| パラメーター   | 型|説明|
|:---------------|:--------|:----------|
|paragraphText|string|必須。挿入する段落テキスト。|
|insertLocation|InsertLocation|必須。有効な値は、'Before' または 'After' です。|

#### 戻り値
[Paragraph](paragraph.md)

### insertText(text: string, insertLocation: InsertLocation)
テキストを本文の指定された位置に挿入します。insertLocation の値は、'Before' か 'After' になります。

#### 構文
```js
inlinePictureObject.insertText(text, insertLocation);
```

#### パラメータ
| パラメーター   | 型|説明|
|:---------------|:--------|:----------|
|text|string|必須。挿入するテキスト。|
|insertLocation|InsertLocation|必須。有効な値は、'Before' または 'After' です。|

#### 戻り値
[Range](range.md)

### select(selectionMode: SelectionMode)
画像を選択し、その本文に Word の UI を移動します。selectionMode 値は、'Select'、'Start'、'End' のいずれかになります。

#### 構文
```js
inlinePictureObject.select(selectionMode);
```

#### パラメータ
| パラメーター   | 型|説明|
|:---------------|:--------|:----------|
|selectionMode|SelectionMode|省略可能。選択モードは、'Select'、'Start'、'End' のいずれかになります。'Select' が既定値です。|

#### 戻り値
void

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
void

## サポートの詳細

実行時のチェックで[要件セット](https://msdn.microsoft.com/EN-US/library/office/mt590206.aspx)を使用して、アプリケーションが Word のホスト バージョンによってサポートされていることを確かめます。Office ホスト アプリケーションとサーバーの要件の詳細については、「[Office アドインを実行するための要件](https://msdn.microsoft.com/EN-US/library/office/dn833104.aspx)」を参照してください。 
