

# ProjectDocument オブジェクト
Office アドインが対話するプロジェクト ドキュメント (アクティブ プロジェクト) を表す抽象クラス。

|||
|:-----|:-----|
|**ホスト:**|Project|
|**で追加**|1.0|

```js
Office.context.document
```


## メンバー


**メソッド**


|**名前**|**説明**|
|:-----|:-----|
|[addHandlerAsync メソッド](../../reference/shared/projectdocument.addhandlerasync.md)|**ProjectDocument** オブジェクトのイベントのイベント ハンドラーを非同期的に追加します。|
|[getMaxResourceIndexAsync メソッド](../../reference/shared/projectdocument.getmaxresourceindexasync.md)|現在のプロジェクトでリソースのコレクションの最大インデックスを非同期に取得します。|
|[getMaxTaskIndexAsync メソッド](../../reference/shared/projectdocument.getmaxtaskindexasync.md)|現在のプロジェクトでタスクのコレクションの最大インデックスを非同期に取得します。|
|[getProjectFieldAsync メソッド](../../reference/shared/projectdocument.getprojectfieldasync.md)|アクティブなプロジェクトの指定したフィールドの値を非同期に取得します。|
|[getResourceByIndexAsync メソッド](../../reference/shared/projectdocument.getresourcebyindexasync.md)|リソースのコレクション内に指定のインデックスがあるリソースの GUID を非同期に取得します。|
|[getResourceFieldAsync メソッド](../../reference/shared/projectdocument.getresourcefieldasync.md)|指定したリソースの指定したフィールドの値を非同期に取得します。|
|[getSelectedDataAsync メソッド](../../reference/shared/projectdocument.getselecteddataasync.md)|ガント チャート内で現在選択されている 1 つ以上のセルに格納されているデータを非同期に取得します。|
|[getSelectedResourceAsync メソッド](../../reference/shared/projectdocument.getselectedresourceasync.md)|選択されているリソースの GUID を非同期に取得します。|
|[getSelectedTaskAsync メソッド](../../reference/shared/projectdocument.getselectedtaskasync.md)|選択されているタスクの GUID を非同期に取得します。|
|[getSelectedViewAsync メソッド](../../reference/shared/projectdocument.getselectedviewasync.md)|アクティブ ビューのビューの種類と名前を非同期に取得します。|
|[getTaskAsync メソッド](../../reference/shared/projectdocument.gettaskasync.md)|タスク名、タスクに割り当てられているリソース、およびタスクの ID を同期済みの SharePoint タスク リストから非同期的に取得します。|
|[getTaskByIndexAsync メソッド](../../reference/shared/projectdocument.gettaskbyindexasync.md)|タスクのコレクション内に指定のインデックスがあるタスクの GUID を非同期に取得します。|
|[getTaskFieldAsync メソッド](../../reference/shared/projectdocument.gettaskfieldasync.md)|指定したタスクの指定したフィールドの値を非同期に取得します。|
|[getWSSUrlAsync メソッド](../../reference/shared/projectdocument.getwssurlasync.md)|同期済みの SharePoint タスク リストの URL を非同期に取得します。|
|[removeHandlerAsync メソッド](../../reference/shared/projectdocument.removehandlerasync.md)|**ProjectDocument** オブジェクトのイベントのイベント ハンドラーを非同期的に削除します。|
|[setResourceFieldAsync メソッド](../../reference/shared/projectdocument.setresourcefieldasync.md)|指定したリソースの指定したフィールドの値を非同期に設定します。|
|[setTaskFieldAsync メソッド](../../reference/shared/projectdocument.settaskfieldasync.md)|指定したタスクの指定したフィールドの値を非同期に設定します。|

**Events**


|**名前**|**説明**|
|:-----|:-----|
|[ResourceSelectionChanged イベント](../../reference/shared/projectdocument.resourceselectionchanged.event.md)|アクティブ プロジェクト内でリソースの選択が変更されるときに発生します。|
|[TaskSelectionChanged イベント](../../reference/shared/projectdocument.taskselectionchanged.event.md)|アクティブ プロジェクト内でタスクの選択が変更されるときに発生します。|
|[ViewSelectionChanged イベント](../../reference/shared/projectdocument.viewselectionchanged.event.md)|アクティブなプロジェクトでアクティブ ビューが変更されたときに発生します。|

## 注釈

スクリプトで  **ProjectDocument** オブジェクトを直接呼び出したりインスタンス化したりしないでください。


## 例

次の例では、アドインを初期化してから、Project ドキュメントのコンテキストで入手可能な [Document](../../reference/shared/document.md) オブジェクトのプロパティを取得します。Project ドキュメントは開いたアクティブなプロジェクトです。 **ProjectDocument** オブジェクトのメンバーにアクセスするには、 **ProjectDocument** メソッドとイベントのコード例に示されているように、 **Office.context.document** オブジェクトを使用します。

この例では、アドインに jQuery ライブラリへの参照が指定されており、ページ本文の 内容 div で次のページ コントロールが定義されていることを想定しています。




```HTML
<span id="message"></span>
```




```js
(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {

            // Get information about the document.
            showDocumentProperties();
        });
    };

    // Get the document mode and the URL of the active project.
    function showDocumentProperties() {
        var output = String.format(
            'The document mode is {0}.<br/>The URL of the active project is {1}.',
            Office.context.document.mode,
            Office.context.document.url);
        $('#message').html(output);
    }
})();
```


## サポートの詳細


次の表で、大文字 Y は、このオブジェクトは、対応する Office ホスト アプリケーションでサポートされていることを示します。空のセルは、Office ホスト アプリケーションでこのオブジェクトをサポートしないことを示します。

Office ホスト アプリケーションとサーバーの要件の詳細については、「[Office アドインを実行するための要件](../../docs/overview/requirements-for-running-office-add-ins.md)」をご覧ください。


||**Windows デスクトップ版 Office**|**Office Online (ブラウザー)**|
|:-----|:-----|:-----|
|**Project**|Y||

|||
|:-----|:-----|
|**アプリの種類**|作業ウィンドウ|
|**ライブラリ**|Office.js|
|**名前空間**|Office|

## サポート履歴


|**変更内容**|**1.1**|
|:-----|:-----|
|1.0|導入|

## 関連項目



#### その他の技術情報


[Project 用の作業ウィンドウ アドイン](../../docs/project/project-add-ins.md)
[Document オブジェクト](../../reference/shared/document.md)

