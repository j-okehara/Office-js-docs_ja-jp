# Application オブジェクト (JavaScript API for OneNote)

_適用対象:OneNote Online_


グローバルにアドレス可能な OneNote オブジェクト (ノートブック、アクティブなノートブック、アクティブなセクションなど) すべてを含む最上位のオブジェクトを表します。

## プロパティ

なし

## 関係
| リレーションシップ | 型   |説明| フィードバック|
|:---------------|:--------|:----------|:-------|
|notebooks|[NotebookCollection](notebookcollection.md)|OneNote アプリケーション インスタンスで開いているノートブックのコレクションを取得します。OneNote Online では、ノートブックはアプリケーション インスタンスで一度に 1 つだけ開かれます。読み取り専用です。|[実行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-application-notebooks)|

## メソッド

| メソッド           | 戻り値の型    |説明| フィードバック|
|:---------------|:--------|:----------|:-------|
|[getActiveNotebook()](#getactivenotebook)|[Notebook](notebook.md)|存在する場合はアクティブなノートブックを取得します。 アクティブなノートブックがない場合は、ItemNotFound をスローします。|[実行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-application-getActiveNotebook)|
|[getActiveNotebookOrNull()](#getactivenotebookornull)|[Notebook](notebook.md)|存在する場合はアクティブなノートブックを取得します。 アクティブなノートブックがない場合は、null を返します。|[実行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-application-getActiveNotebookOrNull)|
|[getActiveOutline()](#getactiveoutline)|[Outline](outline.md)|存在する場合はアクティブなアウトラインを取得します。アクティブなアウトラインがない場合は、ItemNotFound をスローします。|[実行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-application-getActiveOutline)|
|[getActiveOutlineOrNull()](#getactiveoutlineornull)|[Outline](outline.md)|存在する場合はアクティブなアウトラインを取得します。アクティブなアウトラインがない場合は、null を返します。|[実行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-application-getActiveOutlineOrNull)|
|[getActivePage()](#getactivepage)|[Page](page.md)|存在する場合はアクティブなページを取得します。 アクティブなページがない場合は、ItemNotFound をスローします。|[実行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-application-getActivePage)|
|[getActivePageOrNull()](#getactivepageornull)|[Page](page.md)|存在する場合はアクティブなページを取得します。 アクティブなページがない場合は、null を返します。|[実行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-application-getActivePageOrNull)|
|[getActiveSection()](#getactivesection)|[Section](section.md)|存在する場合はアクティブなセクションを取得します。 アクティブなセクションがない場合は、ItemNotFound をスローします。|[実行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-application-getActiveSection)|
|[getActiveSectionOrNull()](#getactivesectionornull)|[Section](section.md)|存在する場合はアクティブなセクションを取得します。 アクティブなセクションがない場合は、null を返します。|[実行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-application-getActiveSectionOrNull)|
|[load(param: object)](#loadparam-object)|void|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|[実行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-application-load)|
|[navigateToPage(page:Page)](#navigatetopagepage-page)|void|アプリケーション インスタンスで指定されたページを開きます。|[実行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-application-navigateToPage)|
|[navigateToPageWithClientUrl(url: string)](#navigatetopagewithclienturlurl-string)|[Page](page.md)|指定されたページを取得し、アプリケーション インスタンスで開きます。|[実行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-application-navigateToPageWithClientUrl)|

## メソッドの詳細


### getActiveNotebook()
存在する場合はアクティブなノートブックを取得します。 アクティブなノートブックがない場合は、ItemNotFound をスローします。

#### 構文
```js
applicationObject.getActiveNotebook();
```

#### パラメーター
なし

#### 戻り値
[Notebook](notebook.md)

#### 例
```js
OneNote.run(function (context) {
        
    // Get the active notebook.
    var notebook = context.application.getActiveNotebook();
            
    // Queue a command to load the notebook. 
    // For best performance, request specific properties.           
    notebook.load('id,name');
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
                    
            // Show some properties.
            console.log("Notebook name: " + notebook.name);
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


### getActiveNotebookOrNull()
存在する場合はアクティブなノートブックを取得します。 アクティブなノートブックがない場合は、null を返します。

#### 構文
```js
applicationObject.getActiveNotebookOrNull();
```

#### パラメーター
なし

#### 戻り値
[Notebook](notebook.md)

#### 例
```js
OneNote.run(function (context) {

    // Get the active notebook.
    var notebook = context.application.getActiveNotebookOrNull();

    // Queue a command to load the notebook. 
    // For best performance, request specific properties.           
    notebook.load('id,name');

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {

            // check if active notebook is set.
            if (!notebook.isNull) {
                console.log("Notebook name: " + notebook.name);
                console.log("Notebook ID: " + notebook.id);
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


### getActiveOutline()
存在する場合はアクティブなアウトラインを取得します。アクティブなアウトラインがない場合は、ItemNotFound をスローします。

#### 構文
```js
applicationObject.getActiveOutline();
```

#### パラメーター
なし

#### 戻り値
[アウトライン](outline.md)

#### 例
```js
OneNote.run(function (context) {

    // get active outline.
    var outline = context.application.getActiveOutline();

    // Queue a command to load the id of the outline.         
    outline.load('id');

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {

            // Show some properties.
            console.log("outline id: " + outline.id);
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```


### getActiveOutlineOrNull()
存在する場合はアクティブなアウトラインを取得します。アクティブなアウトラインがない場合は、null を返します。

#### 構文
```js
applicationObject.getActiveOutlineOrNull();
```

#### パラメーター
なし

#### 戻り値
[アウトライン](outline.md)

#### 例
```js
OneNote.run(function (context) {

    // get active outline.
    var outline = context.application.getActiveOutlineOrNull();

    // Queue a command to load the id of the outline.         
    outline.load('id');

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {

            if (!outline.isNull) {
                console.log("outline id: " + outline.id);
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


### getActivePage()
存在する場合はアクティブなページを取得します。 アクティブなページがない場合は、ItemNotFound をスローします。

#### 構文
```js
applicationObject.getActivePage();
```

#### パラメーター
なし

#### 戻り値
[Page](page.md)

#### 例
```js
OneNote.run(function (context) {
        
    // Get the active page.
    var page = context.application.getActivePage();
            
    // Queue a command to load the page. 
    // For best performance, request specific properties.           
    page.load('id,title');
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
                    
            // Show some properties.
            console.log("Page title: " + page.title);
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


### getActivePageOrNull()
存在する場合はアクティブなページを取得します。 アクティブなページがない場合は、null を返します。

#### 構文
```js
applicationObject.getActivePageOrNull();
```

#### パラメーター
なし

#### 戻り値
[Page](page.md)

#### 例
```js
OneNote.run(function (context) {

    // Get the active page.
    var page = context.application.getActivePageOrNull();

    // Queue a command to load the page. 
    // For best performance, request specific properties.           
    page.load('id,title');

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            
            if (!page.isNull) {
                // Show some properties.
                console.log("Page title: " + page.title);
                console.log("Page ID: " + page.id);
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


### getActiveSection()
存在する場合はアクティブなセクションを取得します。 アクティブなセクションがない場合は、ItemNotFound をスローします。

#### 構文
```js
applicationObject.getActiveSection();
```

#### パラメーター
なし

#### 戻り値
[セクション](section.md)

#### 例
```js
OneNote.run(function (context) {
        
    // Get the active section.
    var section = context.application.getActiveSection();
            
    // Queue a command to load the section. 
    // For best performance, request specific properties.           
    section.load('id,name');
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
                    
            // Show some properties.
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


### getActiveSectionOrNull()
存在する場合はアクティブなセクションを取得します。 アクティブなセクションがない場合は、null を返します。

#### 構文
```js
applicationObject.getActiveSectionOrNull();
```

#### パラメーター
なし

#### 戻り値
[セクション](section.md)

#### 例
```js
OneNote.run(function (context) {

    // Get the active section.
    var section = context.application.getActiveSectionOrNull();

    // Queue a command to load the section. 
    // For best performance, request specific properties.           
    section.load('id,name');

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            if (!section.isNull) {
                // Show some properties.
                console.log("Section name: " + section.name);
                console.log("Section ID: " + section.id);
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

### navigateToPage(page:Page)
アプリケーション インスタンスで指定されたページを開きます。

#### 構文
```js
applicationObject.navigateToPage(page);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|page|Page|開くページです。|

#### 戻り値
void

#### 例
```js        
OneNote.run(function (context) {
        
    // Get the pages in the current section.
    var pages = context.application.getActiveSection().pages;
            
    // Queue a command to load the pages. 
    // For best performance, request specific properties.           
    pages.load('id');
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
                    
            // This example loads the first page in the section.
            var page = pages.items[0];
                        
            // Open the page in the application.                    
            context.application.navigateToPage(page);
                    
            // Run the queued command.
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


### navigateToPageWithClientUrl(url: string)
指定されたページを取得し、アプリケーション インスタンスで開きます。

#### 構文
```js
applicationObject.navigateToPageWithClientUrl(url);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|url|string|開くページのクライアント URL です。|

#### 戻り値
[Page](page.md)

#### 例
```js
OneNote.run(function (context) {

    // Get the pages in the current section.
    var pages = context.application.getActiveSection().pages;

    // Queue a command to load the pages. 
    // For best performance, request specific properties.           
    pages.load('clientUrl');

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {

            // This example loads the first page in the section.
            var page = pages.items[0];

            // Open the page in the application.                    
            context.application.navigateToPageWithClientUrl(page.clientUrl);

            // Run the queued command.
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
