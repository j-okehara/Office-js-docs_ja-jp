# セクション オブジェクト (JavaScript API for Word)

Word 文書内のセクションを表します。

_適用対象:Word 2016、Word for iPad、Word for Mac_

## プロパティ
なし

## 関係
| リレーションシップ | 型|説明|
|:---------------|:--------|:----------|
|body|[Body](body.md)|セクションの本文を取得します。これには、headerfooter およびその他のセクション メタデータは含まれません。読み取り専用です。|

## メソッド

| メソッド   | 戻り値の型|説明|
|:---------------|:--------|:----------|
|[getFooter(type: HeaderFooterType)](#getfootertype-headerfootertype)|[Body](body.md)|セクションのフッターの 1 つを取得します。|
|[getHeader(type: HeaderFooterType)](#getheadertype-headerfootertype)|[Body](body.md)|セクションのヘッダーの 1 つを取得します。|
|[load(param: object)](#loadparam-object)|void|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|

## メソッドの詳細

### getFooter(type: HeaderFooterType)
セクションのフッターの 1 つを取得します。

#### 構文
```js
sectionObject.getFooter(type);
```

#### パラメータ
| パラメーター   | 型|説明|
|:---------------|:--------|:----------|
|型|HeaderFooterType|必須。返されるフッターの型。この値は 'primary'、'firstPage'、または 'evenPages' です。|

#### 戻り値
[Body](body.md)

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
		
		// Create a proxy object the primary footer of the first section. 
		// Note that the footer is a body object.
		var myFooter = mySections.items[0].getFooter("primary");
		
		// Queue a command to insert text at the end of the footer.
		myFooter.insertText("This is a footer.", Word.InsertLocation.end);
		
		// Queue a command to wrap the header in a content control.
		myFooter.insertContentControl();
							  
		// Synchronize the document state by executing the queued commands, 
		// and return a promise to indicate task completion.
		return context.sync().then(function () {
			console.log("Added a footer to the first section.");
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
### getHeader(type: HeaderFooterType)
セクションのヘッダーの 1 つを取得します。

#### 構文
```js
sectionObject.getHeader(type);
```

#### パラメータ
| パラメーター   | 型|説明|
|:---------------|:--------|:----------|
|型|HeaderFooterType|必須。返されるヘッダーの型。この値は 'primary'、'firstPage'、または 'evenPages' です。|

#### 戻り値
[Body](body.md)

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
