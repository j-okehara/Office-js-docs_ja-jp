# Section オブジェクト (JavaScript API for OneNote)

_適用対象:OneNote Online_   


OneNote セクションを表します。 セクションには、ページを含めることができます。

## プロパティ

| プロパティ     | 型   |説明|フィードバック|
|:---------------|:--------|:----------|:-------|
|clientUrl|string|セクションのクライアント url です。 読み取り専用です。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-section-clientUrl)|
|id|string|セクションの ID を取得します。 読み取り専用です。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-section-id)|
|name|string|セクションの名前を取得します。 読み取り専用です。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-section-name)|

_プロパティのアクセスの[例](#例)を参照してください。_

## リレーションシップ
| リレーションシップ | 型   |説明| フィードバック|
|:---------------|:--------|:----------|:-------|
|notebook|[Notebook](notebook.md)|セクションを含むノートブックを取得します。 読み取り専用です。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-section-notebook)|
|pages|[PageCollection](pagecollection.md)|セクション内のページのコレクションです。 読み取り専用です。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-section-pages)|
|parentSectionGroup|[セクション グループ](sectiongroup.md)|セクションを含むセクション グループを取得します。 セクションがノートブックの直接の子である場合は ItemNotFound をスローします。 読み取り専用です。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-section-parentSectionGroup)|
|parentSectionGroupOrNull|[セクション グループ](sectiongroup.md)|セクションを含むセクション グループを取得します。 セクションがノートブックの直接の子である場合は null を返します。 読み取り専用です。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-section-parentSectionGroupOrNull)|

## メソッド

| メソッド           | 戻り値の型    |説明| フィードバック|
|:---------------|:--------|:----------|:-------|
|[addPage(title: string)](#addpagetitle-string)|[Page](page.md)|セクションの末尾に新しいページを追加します。|[実行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-section-addPage)|
|[copyToNotebook(destinationNotebook:Notebook)](#copytonotebookdestinationnotebook-notebook)|[Section](section.md)|指定したノートブックに、このセクションをコピーします。|[実行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-section-copyToNotebook)|
|[copyToSectionGroup(destinationSectionGroup: SectionGroup)](#copytosectiongroupdestinationsectiongroup-sectiongroup)|[Section](section.md)|指定したセクション グループに、このセクションをコピーします。|[実行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-section-copyToSectionGroup)|
|[insertSectionAsSibling(location: string、title: string)](#insertsectionassiblinglocation-string-title-string)|[セクション](section.md)|現在のセクションの前か後に、新しいセクションを挿入します。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-section-insertSectionAsSibling)|
|[load(param: object)](#loadparam-object)|void|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-section-load)|

## メソッドの詳細


### addPage(title: string)
セクションの末尾に新しいページを追加します。

#### 構文
```js
sectionObject.addPage(title);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|title|string|新しいページのタイトルです。|

#### 戻り値
[Page](page.md)

#### 例
```js
OneNote.run(function (context) {
            
    // Queue a command to add a page to the current section.
    var page = context.application.getActiveSection().addPage("Wish list");
            
    // Queue a command to load the id and title of the new page. 
    // This example loads the new page so it can read its properties later.           
    page.load('id,title');
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
             
            // Display the properties.       
            console.log("Page name: " + page.title);
            console.log("Page ID: " + page.id);

        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```


### copyToNotebook(destinationNotebook:Notebook)
指定したノートブックに、このセクションをコピーします。

#### 構文
```js
sectionObject.copyToNotebook(destinationNotebook);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|destinationNotebook|ノートブック|このセクションのコピー先のノートブック。|

#### 戻り値
[セクション](section.md)

#### 例
```js
OneNote.run(function (context) {
    var app = context.application;
    
    // Gets the active Notebook.
    var notebook = app.getActiveNotebook();
    
    // Gets the active Section.
    var section = app.getActiveSection();
    
    var newSection;
    
    return context.sync()
        .then(function() {
            newSection = section.copyToNotebook(notebook);
            newSection.load('id');
            return context.sync();
        })
        .then(function() {
            console.log(newSection.id);
        });
})
.catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```


### copyToSectionGroup(destinationSectionGroup: SectionGroup)
指定したセクション グループに、このセクションをコピーします。

#### 構文
```js
sectionObject.copyToSectionGroup(destinationSectionGroup);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|destinationSectionGroup|セクション グループ|このセクションのコピー先のセクション グループ。|

#### 戻り値
[セクション](section.md)

#### 例
```js
OneNote.run(function (ctx) {
    var app = ctx.application;
    
    // Gets the active Notebook.
    var notebook = app.getActiveNotebook();
    
    // Gets the active Section.
    var section = app.getActiveSection();
    
    var newSection;
    
    return ctx.sync()
        .then(function() {
            var firstSectionGroup = notebook.sectionGroups.items[0];
            newSection = section.copyToSectionGroup(firstSectionGroup);
            newSection.load('id');
            return ctx.sync();
        })
        .then(function() {
            console.log(newSection.id);
        });
})
.catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```


### insertSectionAsSibling(location: string、title: string)
現在のセクションの前か後に、新しいセクションを挿入します。

#### 構文
```js
sectionObject.insertSectionAsSibling(location, title);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|location|string|現在のセクションを基準にした新しいセクションの場所です。使用可能な値は次のとおりです。Before、After。|
|title|string|新しいセクションの名前を指定します。|

#### 戻り値
[セクション](section.md)

#### 例
```js
OneNote.run(function (context) {
            
    // Queue a command to insert a section after the current section.
    var section = context.application.getActiveSection().insertSectionAsSibling("After", "New section");
            
    // Queue a command to load the id and name of the new section. 
    // This example loads the new section so it can read its properties later.           
    section.load('id,name');
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
             
            // Display the properties.       
            console.log("Section name: " + section.name);
            console.log("Section ID: " + section.id);
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```


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
### プロパティのアクセスの例

**id**
```js
OneNote.run(function (context) {
        
    // Get the current section.
    var section = context.application.getActiveSection();
            
    // Queue a command to load the section. 
    // For best performance, request specific properties.           
    section.load("id");
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            console.log("Section ID: " + section.id);
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

**name と notebook**
```js
OneNote.run(function (context) {
        
    // Get the current section.
    var section = context.application.getActiveSection();
            
    // Queue a command to load the section with the specified properties. 
    section.load("name,notebook/name");
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            console.log("Section name: " + section.name);
            console.log("Parent notebook name: " + section.notebook.name);
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

**parentSectionGroupOrNull**
```js
OneNote.run(function (context) {
    // Queue a command to add a page to the current section.
    var section = context.application.getActiveSection();
    section.load('clientUrl,notebook');
    var sectionGroup = section.parentSectionGroupOrNull;
    
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            if(sectionGroup.isNull === false)
            {
                // If a parent section group exists, queue a command to add a section in it!
                sectionGroup.addSection("NewSectionInSectionGroup");
            }
            return context.sync();
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```
    
