# SectionCollection オブジェクト (JavaScript API for OneNote)

_適用対象:OneNote Online_  


セクションのコレクションを表します。

## プロパティ

| プロパティ     | 型   |説明|フィードバック|
|:---------------|:--------|:----------|:-------|
|count|int|コレクション内のセクションの数を取得します。読み取り専用です。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionCollection-count)|
|Items|[Section[]](section.md)|セクション オブジェクトのコレクション。読み取り専用です。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionCollection-items)|

_プロパティのアクセスの[例](#例)を参照してください。_

## リレーションシップ
なし


## メソッド

| メソッド           | 戻り値の型    |説明| フィードバック|
|:---------------|:--------|:----------|:-------|
|[getByName(name: string)](#getbynamename-string)|[SectionCollection](sectioncollection.md)|指定した名前のセクションのコレクションを取得します。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionCollection-getByName)|
|[getItem(index: number または string)](#getitemindex-number-または-string)|[セクション](section.md)|ID やコレクション内のインデックスで、セクションを取得します。読み取り専用です。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionCollection-getItem)|
|[getItemAt(index: number)](#getitematindex-number)|[セクション](section.md)|コレクション内での位置を基にセクションを取得します。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionCollection-getItemAt)|
|[load(param: object)](#loadparam-object)|void|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionCollection-load)|

## メソッドの詳細


### getByName(name: string)
指定した名前のセクションのコレクションを取得します。

#### 構文
```js
sectionCollectionObject.getByName(name);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|name|string|セクションの名前。|

#### 戻り値
[SectionCollection](sectioncollection.md)

#### 例
```js
OneNote.run(function (context) {

    // Get the sections in the current notebook.
    var sections = context.application.getActiveNotebook().sections;

    // Queue a command to load the sections. 
    // For best performance, request specific properties.
    sections.load("id"); 
    
    // Get the sections with the specified name.
    var groceriesSections = sections.getByName("Groceries");
    
    // Queue a command to load the sections with the specified name.
    groceriesSections.load("id,name");

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {

            // Iterate through the collection or access items individually by index.
            if (groceriesSections.items.length > 0) {
                console.log("Section name: " + groceriesSections.items[0].name);
                console.log("Section ID: " + groceriesSections.items[0].id);
            }
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

### getItem(index: number または string)
ID やコレクション内のインデックスで、セクションを取得します。読み取り専用です。

#### 構文
```js
sectionCollectionObject.getItem(index);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|index|number または string|セクションの ID、またはコレクション内のセクションのインデックスの場所です。|

#### 戻り値
[セクション](section.md)

### getItemAt(index: number)
コレクション内での位置を基にセクションを取得します。

#### 構文
```js
sectionCollectionObject.getItemAt(index);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|index|number|取得するオブジェクトのインデックス値。0 を起点とする番号になります。|

#### 戻り値
[セクション](section.md)

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

**Items**
```js
OneNote.run(function (context) {

    // Get the sections in the current notebook.
    var sections = context.application.getActiveNotebook().sections;

    // Queue a command to load the sections. 
    // For best performance, request specific properties.
    sections.load("name"); 

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            
            // Iterate through the collection or access items individually by index, for example: sections.items[0]
            $.each(sections.items, function(index, section) {
                if (section.name === "Homework") {
                    section.addPage("Biology");
                    section.addPage("Spanish");
                    section.addPage("Computer Science");
                }
            });
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

