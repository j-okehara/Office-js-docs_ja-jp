# SectionGroup オブジェクト (JavaScript API for OneNote)

_適用対象:OneNote Online_   


OneNote セクション グループを表します。 セクション グループに含めることができるのは、セクションとその他のセクション グループです。

## プロパティ

| プロパティ     | 型   |説明|フィードバック|
|:---------------|:--------|:----------|:-------|
|clientUrl{|string|セクション グループのクライアント url です。 読み取り専用です。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionGroup-clientUrl{)|
|id|string|セクション グループの ID を取得します。 読み取り専用です。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionGroup-id)|
|name|string|セクション グループの名前を取得します。 読み取り専用です。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionGroup-name)|

_プロパティのアクセスの[例](#例)を参照してください。_

## リレーションシップ
| リレーションシップ | 型   |説明| フィードバック|
|:---------------|:--------|:----------|:-------|
|notebook|[Notebook](notebook.md)|セクション グループを含むノートブックを取得します。 読み取り専用です。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionGroup-notebook)|
|parentSectionGroup|[セクション グループ](sectiongroup.md)|セクション グループを含むセクション グループを取得します。 セクション グループがノートブックの直接の子である場合は ItemNotFound をスローします。 読み取り専用です。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionGroup-parentSectionGroup)|
|parentSectionGroupOrNull|[セクション グループ](sectiongroup.md)|セクション グループを含むセクション グループを取得します。 セクション グループがノートブックの直接の子である場合は null を返します。 読み取り専用です。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionGroup-parentSectionGroupOrNull)|
|sectionGroups|[SectionGroupCollection](sectiongroupcollection.md)|セクション グループ内のセクション グループのコレクションです。 読み取り専用です。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionGroup-sectionGroups)|
|sections|[SectionCollection](sectioncollection.md)|セクション グループ内のセクションのコレクションです。 読み取り専用です。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionGroup-sections)|

## メソッド

| メソッド           | 戻り値の型    |説明| フィードバック|
|:---------------|:--------|:----------|:-------|
|[addSection(title:String)](#addsectiontitle-string)|[セクション](section.md)|セクション グループの末尾に新しいセクションを追加します。|[実行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionGroup-addSection)|
|[addSectionGroup(name:String)](#addsectiongroupname-string)|[セクション グループ](sectiongroup.md)|この sectionGroup の末尾に新しいセクション グループを追加します。|[実行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionGroup-addSectionGroup)|
|[load(param: object)](#loadparam-object)|void|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionGroup-load)|

## メソッドの詳細


### addSection(title:String)
セクション グループの末尾に新しいセクションを追加します。

#### 構文
```js
sectionGroupObject.addSection(title);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|title|String|新しいセクションの名前を指定します。|

#### 戻り値
[セクション](section.md)

#### 例
```js
OneNote.run(function (context) {

    // Get the section groups that are direct children of the current notebook.
    var sectionGroups = context.application.getActiveNotebook().sectionGroups;
    
    // Queue a command to load the section groups.
    // For best performance, request specific properties.
    sectionGroups.load("id");

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function() {
            
            // Add a section to each section group.
            $.each(sectionGroups.items, function(index, sectionGroup) {
                sectionGroup.addSection("Agenda");
            });
            
            // Run the queued commands.
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


### addSectionGroup(name:String)
この sectionGroup の末尾に新しいセクション グループを追加します。

#### 構文
```js
sectionGroupObject.addSectionGroup(name);
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
    var sectionGroup;
    var nestedSectionGroup;

    // Gets the active notebook.
    var notebook = context.application.getActiveNotebook();

    // Queue a command to add a new section group.
    var sectionGroups = notebook.sectionGroups;

    // Queue a command to load the new section group.
    sectionGroups.load();

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function(){
            sectionGroup = sectionGroups.items[0];
            sectionGroup.load();
            return context.sync();
        })
        .then(function(){
            nestedSectionGroup = sectionGroup.addSectionGroup("Sample nested section group");
            nestedSectionGroup.load();
            return context.sync();
        })
        .then(function() {
            console.log("New nested section group name is " + nestedSectionGroup.name);
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
        
    // Get the parent section group that contains the current section.
    var sectionGroup = context.application.getActiveSection().parentSectionGroup;
            
    // Queue a command to load the section group. 
    // For best performance, request specific properties.           
    sectionGroup.load("id,name");
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            
            // Write the properties.
            console.log("Section group name: " + sectionGroup.name);
            console.log("Section group ID: " + sectionGroup.id);
            
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
        
    // Get the parent section group that contains the current section.
    var sectionGroup = context.application.getActiveSection().parentSectionGroup;
            
    // Queue a command to load the section group with the specified properties.           
    sectionGroup.load("name,notebook/name"); 
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {

            // Write the properties.
            console.log("Section group name: " + sectionGroup.name);
            console.log("Parent notebook name: " + sectionGroup.notebook.name);
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

    // Get the section groups that are direct children of the current notebook.
    var sectionGroups = context.application.getActiveNotebook().sectionGroups;

    // Queue a command to load the section groups.
    // For best performance, request specific properties.
    sectionGroups.load("name");
    
    // Get the child section groups of the first section group in the notebook.
    var nestedSectionGroups = sectionGroups._GetItem(0).sectionGroups;
    
    // Queue a command to load the ID and name properties of the child section groups.
    nestedSectionGroups.load("id,name");
    
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function() {
            
            // Write the properties for each child section group.
            $.each(nestedSectionGroups.items, function(index, sectionGroup) {
                console.log("Section group name: " + sectionGroup.name);  
                console.log("Section group ID: " + sectionGroup.id);  
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

    // Get the sections that are siblings of the current section.
    var sections = context.application.getActiveSection().parentSectionGroup.sections;

    // Queue a command to load the section groups.
    // For best performance, request specific properties.
    sections.load("id,name");
    
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function() {
            
            // Write the properties for each section.
            $.each(sections.items, function(index, section) {
                console.log("Section name: " + section.name);  
                console.log("Section ID: " + section.id);  
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

