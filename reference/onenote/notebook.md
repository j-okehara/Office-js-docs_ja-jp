# Notebook オブジェクト (JavaScript API for OneNote)

_適用対象:OneNote Online_   


OneNote ノートブックを表します。 ノートブックには、セクション グループとセクションが含まれます。

## プロパティ

| プロパティ     | 型   |説明|フィードバック|
|:---------------|:--------|:----------|:-------|
|clientUrl|string|ノートブックのクライアント URL。 読み取り専用です。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-notebook-clientUrl)|
|id|string|ノートブックの ID を取得します。 読み取り専用です。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-notebook-id)|
|name|string|ノートブックの名前を取得します。 読み取り専用です。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-notebook-name)|

_プロパティのアクセスの[例](#例)を参照してください。_

## リレーションシップ
| リレーションシップ | 型   |説明| フィードバック|
|:---------------|:--------|:----------|:-------|
|sectionGroups|[SectionGroupCollection](sectiongroupcollection.md)|ノートブック内のセクション グループ。 読み取り専用です。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-notebook-sectionGroups)|
|sections|[SectionCollection](sectioncollection.md)|ノートブックのセクション。 読み取り専用です。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-notebook-sections)|

## メソッド

| メソッド           | 戻り値の型    |説明| フィードバック|
|:---------------|:--------|:----------|:-------|
|[addSection(name:String)](#addsectionname-string)|[セクション](section.md)|ノートブックの末尾に新しいセクションを追加します。|[実行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-notebook-addSection)|
|[addSectionGroup(name:String)](#addsectiongroupname-string)|[セクション グループ](sectiongroup.md)|ノートブックの末尾に新しいセクション グループを追加します。|[実行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-notebook-addSectionGroup)|
|[load(param: object)](#loadparam-object)|void|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-notebook-load)|

## メソッドの詳細


### addSection(name:String)
ノートブックの末尾に新しいセクションを追加します。

#### 構文
```js
notebookObject.addSection(name);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|name|String|新しいセクションの名前を指定します。|

#### 戻り値
[セクション](section.md)

#### 例
```js          
OneNote.run(function (context) {

    // Gets the active notebook.
    var notebook = context.application.getActiveNotebook();

    // Queue a command to add a new section. 
    var section = notebook.addSection("Sample section");
    
    // Queue a command to load the new section. This example reads the name property later.
    section.load("name");

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function() {
            console.log("New section name is " + section.name);
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
}); 
```


### addSectionGroup(name:String)
ノートブックの末尾に新しいセクション グループを追加します。

#### 構文
```js
notebookObject.addSectionGroup(name);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|name|String|新しいセクションの名前を指定します。|

#### 戻り値
[セクション グループ](sectiongroup.md)

#### 例
```js          
OneNote.run(function (context) {

    // Gets the active notebook.
    var notebook = context.application.getActiveNotebook();

    // Queue a command to add a new section group.
    var sectionGroup = notebook.addSectionGroup("Sample section group");

    // Queue a command to load the new section group.
    sectionGroup.load();

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function() {
            console.log("New section group name is " + sectionGroup.name);
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
        
    // Get the current notebook.
    var notebook = context.application.getActiveNotebook();
            
    // Queue a command to load the notebook. 
    // For best performance, request specific properties.           
    notebook.load('id');
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            console.log("Notebook ID: " + notebook.id);
            
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

**name**
```js
OneNote.run(function (context) {
        
    // Get the current notebook.
    var notebook = context.application.getActiveNotebook();
            
    // Queue a command to load the notebook. 
    // For best performance, request specific properties.           
    notebook.load('name');
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            console.log("Notebook name: " + notebook.name);
            
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

**sectionGroups**
```js          
OneNote.run(function (context) {

    // Get the section groups in the notebook. 
    var sectionGroups = context.application.getActiveNotebook().sectionGroups;

    // Queue a command to load the sectionGroups. 
    sectionGroups.load("name");

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function() {
            $.each(sectionGroups.items, function(index, sectionGroup) {
                console.log("Section group name: " + sectionGroup.name);
            });
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

**sections**
```js
OneNote.run(function (context) {

    // Gets the active notebook.
    var notebook = context.application.getActiveNotebook();
    
    // Queue a command to get immediate child sections of the notebook. 
    var childSections = notebook.sections;

    // Queue a command to load the childSections. 
    context.load(childSections);

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function() {
            $.each(childSections.items, function(index, childSection) {
                console.log("Immediate child section name: " + childSection.name);
            });            
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});   
```

