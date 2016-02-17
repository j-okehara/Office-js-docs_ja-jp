# ContentControlCollection オブジェクト (JavaScript API for Word)

ContentControl オブジェクトのコレクションが含まれます。コンテンツ コントロールは、特定の種類のコンテンツのコンテナーとして機能する、ラベルを付けることのできる、境界線で区切られたドキュメント内の領域です。個々のコンテンツ コントロールには、画像、表、または書式設定されたテキストの段落などを内容として格納できます。現在、リッチ テキストのコンテンツ コントロールのみがサポートされています。

_適用対象:Word 2016、Word for iPad、Word for Mac_

## プロパティ
| プロパティ   | 型|説明
|:---------------|:--------|:----------|
|Items|[ContentControl[]](contentcontrol.md)|contentControl オブジェクトのコレクション。読み取り専用です。|

_プロパティのアクセスの[例](#property-access-examples)を参照してください。_

## 関係
なし


## メソッド

| メソッド   | 戻り値の型|説明|
|:---------------|:--------|:----------|
|[getById(id: number)](#getbyidid-number)|[ContentControl](contentcontrol.md)|コンテンツ コントロールの識別子によってコンテンツ コントロールを取得します。|
|[getByTag(tag: string)](#getbytagtag-string)|[ContentControlCollection](contentcontrolcollection.md)|指定されたタグを含むコンテンツ コントロールを取得します。|
|[getByTitle(title: string)](#getbytitletitle-string)|[ContentControlCollection](contentcontrolcollection.md)|指定されたタイトルを含むコンテンツ コントロールを取得します。|
|[load(param: object)](#loadparam-object)|void|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|

## メソッドの詳細

### getById(id: number)
コンテンツ コントロールの識別子によってコンテンツ コントロールを取得します。

#### 構文
```js
contentControlCollectionObject.getById(id);
```

#### パラメータ
| パラメーター   | 型|説明|
|:---------------|:--------|:----------|
|id|number|必須。コンテンツ コントロールの識別子。|

#### 戻り値
[ContentControl](contentcontrol.md)

#### 例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
	
	// Create a proxy object for the content control that contains a specific id.
	var contentControl = context.document.contentControls.getById(30086310);
		
	// Queue a command to load the text property for a content control. 
	context.load(contentControl, 'text');
	
	// Synchronize the document state by executing the queued commands, 
	// and return a promise to indicate task completion.
	return context.sync().then(function () {
		console.log('The content control with that Id has been found in this document.'); 
	});  
})
.catch(function (error) {
	console.log('Error: ' + JSON.stringify(error));
	if (error instanceof OfficeExtension.Error) {
		console.log('Debug info: ' + JSON.stringify(error.debugInfo));
	}
});
```

### getByTag(tag: string)
指定されたタグを含むコンテンツ コントロールを取得します。

#### 構文
```js
contentControlCollectionObject.getByTag(tag);
```

#### パラメータ
| パラメーター   | 型|説明|
|:---------------|:--------|:----------|
|tag|string|必須。コンテンツ コントロールに設定するタグ。|

#### 戻り値
[ContentControlCollection](contentcontrolcollection.md)

#### 例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the content controls collection that contains a specific tag.
    var contentControlsWithTag = context.document.contentControls.getByTag('Customer-Address');
        
    // Queue a command to load the text property for all of content controls with a specific tag. 
    context.load(contentControlsWithTag, 'text');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        if (contentControlsWithTag.items.length === 0) {
            console.log("There isn't a content control with a tag of Customer-Address in this document.");
        } else {
            console.log('The first content control with the tag of Customer-Address has this text: ' + contentControlsWithTag.items[0].text);    
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

#### 追加情報
[Word-Add-in-DocumentAssembly][contentControls.getByTag] サンプルは、getByTag メソッドを使う別の例を示しています。


### getByTitle(title: string)
指定されたタイトルを含むコンテンツ コントロールを取得します。

#### 構文
```js
contentControlCollectionObject.getByTitle(title);
```

#### パラメータ
| パラメーター   | 型|説明|
|:---------------|:--------|:----------|
|タイトル|string|必須。コンテンツ コントロールのタイトル。|

#### 戻り値
[ContentControlCollection](contentcontrolcollection.md)

#### 例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the content controls collection that contains a specific title.
    var contentControlsWithTitle = context.document.contentControls.getByTitle('Enter Customer Address Here');
        
    // Queue a command to load the text property for all of content controls with a specific title. 
    context.load(contentControlsWithTitle, 'text');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        if (contentControlsWithTitle.items.length === 0) {
            console.log("There isn't a content control with a title of 'Enter Customer Address Here' in this document.");
        } else {
            console.log('The first content control with the title of 'Enter Customer Address Here' has this text: ' + contentControlsWithTitle.items[0].text);    
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

#### 追加情報
[Word-Add-in-DocumentAssembly][contentControls.getByTitle] サンプルは、getByTitle メソッドを使う別の例を示しています。

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
    
    // Create a proxy object for the content controls collection.
    var contentControls = context.document.contentControls;
    
    // Queue a command to load the id property for all of the content controls. 
    context.load(contentControls, 'id');
     
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        if (contentControls.items.length === 0) {
            console.log('No content control found.');
        }
        else {
            // Queue a command to load the properties on the first content control. 
            contentControls.items[0].load(  'appearance,' +
                                            'cannotDelete,' +
                                            'cannotEdit,' +
                                            'color,' +
                                            'id,' +
                                            'placeHolderText,' +
                                            'removeWhenEdited,' +
                                            'title,' +
                                            'text,' +
                                            'type,' +
                                            'style,' +
                                            'tag,' +
                                            'font/size,' +
                                            'font/name,' +
                                            'font/color');             
        
            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    console.log('Property values of the first content control:' + 
                        '   ----- appearance: ' + contentControls.items[0].appearance + 
                        '   ----- cannotDelete: ' + contentControls.items[0].cannotDelete +
                        '   ----- cannotEdit: ' + contentControls.items[0].cannotEdit +
                        '   ----- color: ' + contentControls.items[0].color +
                        '   ----- id: ' + contentControls.items[0].id +
                        '   ----- placeHolderText: ' + contentControls.items[0].placeholderText +
                        '   ----- removeWhenEdited: ' + contentControls.items[0].removeWhenEdited +
                        '   ----- title: ' + contentControls.items[0].title +
                        '   ----- text: ' + contentControls.items[0].text +
                        '   ----- type: ' + contentControls.items[0].type +
                        '   ----- style: ' + contentControls.items[0].style +
                        '   ----- tag: ' + contentControls.items[0].tag +
                        '   ----- font size: ' + contentControls.items[0].font.size +
                        '   ----- font name: ' + contentControls.items[0].font.name +
                        '   ----- font color: ' + contentControls.items[0].font.color);
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

[Silly stories](https://aka.ms/sillystorywordaddin) サンプル アドインは、**load** メソッドを使用して **tag** プロパティと **title** プロパティを含むコンテンツ コントロールのコレクションを読み込む方法を示しています。

## サポートの詳細

実行時のチェックで[要件セット](https://msdn.microsoft.com/EN-US/library/office/mt590206.aspx)を使用して、アプリケーションが Word のホスト バージョンによってサポートされていることを確かめます。Office ホスト アプリケーションとサーバーの要件の詳細については、「[Office アドインを実行するための要件](https://msdn.microsoft.com/EN-US/library/office/dn833104.aspx)」を参照してください。 


[contentControls.getByTag]: https://github.com/OfficeDev/Word-Add-in-DocumentAssembly/blob/master/WordAPIDocAssemblySampleWeb/App/Home/Home.js#L300 "get by tag" [contentControls.getByTitle]: https://github.com/OfficeDev/Word-Add-in-DocumentAssembly/blob/master/WordAPIDocAssemblySampleWeb/App/Home/Home.js#L331 "get by title"

