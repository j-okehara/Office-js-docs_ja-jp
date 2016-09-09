# SectionGroupCollection オブジェクト (JavaScript API for OneNote)

_適用対象:OneNote Online_  


セクション グループのコレクションを表します。

## プロパティ

| プロパティ     | 型   |説明|フィードバック|
|:---------------|:--------|:----------|:-------|
|count|int|コレクション内のセクション グループの数を取得します。読み取り専用です。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionGroupCollection-count)|
|Items|[SectionGroup[]](sectiongroup.md)|sectionGroup オブジェクトのコレクション。読み取り専用です。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionGroupCollection-items)|

_プロパティのアクセスの[例](#例)を参照してください。_

## リレーションシップ
なし


## メソッド

| メソッド           | 戻り値の型    |説明| フィードバック|
|:---------------|:--------|:----------|:-------|
|[getByName(name: string)](#getbynamename-string)|[SectionGroupCollection](sectiongroupcollection.md)|指定した名前のセクション グループのコレクションを取得します。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionGroupCollection-getByName)|
|[getItem(index: number または string)](#getitemindex-number-または-string)|[セクション グループ](sectiongroup.md)|ID やコレクション内のインデックスで、セクション グループを取得します。読み取り専用です。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionGroupCollection-getItem)|
|[getItemAt(index: number)](#getitematindex-number)|[セクション グループ](sectiongroup.md)|コレクション内での位置を基にセクション グループを取得します。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionGroupCollection-getItemAt)|
|[load(param: object)](#loadparam-object)|void|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|[検索](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionGroupCollection-load)|

## メソッドの詳細


### getByName(name: string)
指定した名前のセクション グループのコレクションを取得します。

#### 構文
```js
sectionGroupCollectionObject.getByName(name);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|name|string|セクション グループの名前。|

#### 戻り値
[SectionGroupCollection](sectiongroupcollection.md)

#### 例
```js
OneNote.run(function (context) {

    // Get the section groups that are direct children of the current notebook.
    var sectionGroups = context.application.getActiveNotebook().sectionGroups;

    // Queue a command to load the section groups. 
    // For best performance, request specific properties.
    sectionGroups.load("id"); 

    // Get the section groups with the specified name.
    var labsSectionGroups = sectionGroups.getByName("Labs");

    // Queue a command to load the section groups with the specified properties.
    labsSectionGroups.load("id,name"); 
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {

            // Iterate through the collection or access items individually by index.
            if (labsSectionGroups.items.length > 0) {
                console.log("Section group name: " + labsSectionGroups.items[0].name);
                console.log("Section group ID: " + labsSectionGroups.items[0].id);
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
ID やコレクション内のインデックスで、セクション グループを取得します。読み取り専用です。

#### 構文
```js
sectionGroupCollectionObject.getItem(index);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|index|number または string|セクション グループの ID、またはコレクション内のセクション グループのインデックスの場所です。|

#### 戻り値
[セクション グループ](sectiongroup.md)

### getItemAt(index: number)
コレクション内での位置を基にセクション グループを取得します。

#### 構文
```js
sectionGroupCollectionObject.getItemAt(index);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|index|number|取得するオブジェクトのインデックス値。0 を起点とする番号になります。|

#### 戻り値
[セクション グループ](sectiongroup.md)

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

    // Get the section groups that are direct children of the current notebook.
    var sectionGroups = context.application.getActiveNotebook().sectionGroups;

    // Queue a command to load the section groups. 
    // For best performance, request specific properties.
    sectionGroups.load("name"); 

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            
            // Iterate through the collection or access items individually by index, for example: sectionGroups.items[0]
            $.each(sectionGroups.items, function(index, sectionGroup) {
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

