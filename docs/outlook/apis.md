
# Outlook アドインの API

Outlook アドインで API を使用するには、Office.js ライブラリの場所、要件セット、スキーマ、アクセス許可を指定する必要があります。

## Office.js ライブラリ

Outlook アドイン API を操作するには、Office.js の JavaScript API を使用する必要があります。 ライブラリの CDN は _https://appsforoffice.microsoft.com/lib/1/hosted/Office.js_ です。 Office ストアに送信されるアドインは、この CDN で Office.js を参照する必要があります。ローカル参照は使用できません。 

CDN は、アドインの UI を実装する Web ページ (.html、.aspx、または .php ファイル) の **head** タグ内の **script** タグの **src** 属性で宣言します。


```HTML
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

新しい API が追加されても、Office.js への URL は同じままになります。URL 内のバージョンは、既存の API の動作を分割する場合にのみ変更されます。

> **重要:**Office ホスト アプリケーションのアドインを開発する場合は、ページの `<head>` セクションの内側から JavaScript API for Office を参照します。 これにより、あらゆる body 要素の前に API が完全に初期化されます。 Office ホストでは、アクティブ化の 5 秒以内にアドインを初期化する必要があります。 このしきい値を超えるとアドインが応答なしと宣言され、ユーザーにエラー メッセージが表示されます。  

## 要件セット

すべての Outlook API は、メールボックス要件セットに属しています。メールボックス要件セットには複数のバージョンがあり、リリースされる API の新しいセットはそれぞれのセットの上位バージョンに属します。すべての Outlook クライアントが、リリースされる最新の API のセットをサポートするわけではありませんが、ある要件セットのサポートを宣言している Outlook クライアントは、その要件セットのすべての API をサポートします。 

どの Outlook クライアントにアドインを表示するかを制御するには、最小の要件セットのバージョンをマニフェストで指定します。たとえば、要件セットのバージョン 1.3 を指定すると、最小バージョンの 1.3 をサポートしていない Outlook クライアントにはアドインが表示されなくなります。 

要件セットを指定しても、そのバージョンの API にアドインを限定することにはなりません。要件セット v1.1 を指定しているアドインが、v1.3 をサポートする Outlook クライアントで実行されると、そのアドインは v1.3 の API を使用できます。要件セットでは、どの Outlook クライアントにアドインを表示するかのみを制御します。

マニフェストで指定した要件セットよりも上位の要件セットの API が使用できるかどうかを確認する場合は、標準の JavaScript を使用できます。


```js
if (item.somePropertyOrFunction) {
   item.somePropertyOrFunction...  
}
```

> **注:** このような確認は、マニフェストで指定された要件セットのバージョンに存在する API には必要ありません。

目的のシナリオに不可欠な API セット (その機能なしではアドインが動作しない) をサポートする最小要件セットを指定する必要があります。要件セットは、マニフェストの **Requirements** 要素、**Sets** 要素、**Set** 要素で指定する必要があります。詳細については、「[Outlook アドインのマニフェスト](../outlook/manifests/manifests.md)」および「[Outlook API 要件セットについて](..\..\reference\outlook\tutorial-api-requirement-sets.md)」を参照してください。

**Methods** 要素は Outlook アドインには適用されないため、特定のメソッドについてのサポートは宣言できません。


## アクセス許可

アドインには、そのアドインが必要とする API を使用するための適切なアクセス許可が必要になります。アクセス許可には、4 つのレベルがあります。詳細については、「[Outlook アドインのアクセス許可モデルを理解する](../outlook/understanding-outlook-add-in-permissions.md)」を参照してください。


|**権限レベル**|**説明**|
|:-----|:-----|
|Restricted|エンティティは使用できますが、正規表現は使用できません。|
|アイテムの読み取り|_Restricted_ で許可されているものに加えて、以下のものが許可されます。<ul><li>正規表現</li><li>Outlook アドイン API の読み取りアクセス</li><li>アイテムのプロパティとコールバック トークンの取得</li></ul>|
|Read/write|_Read item_ で許可される内容に加えて、次に示す内容が許可されます。<ul><li>完全な Outlook アドイン API のアクセス (ただし、<b>makeEwsRequestAsync</b> を除く)</li><li>アイテムのプロパティの設定</li></ul>|
|メールボックスの読み取り/書き込み|_読み取り/書き込み_ で許可されているものに加えて、以下のものが許可されます。<ul><li>アイテムおよびフォルダーの作成、読み取り、書き込み</li><li>アイテムの送信</li><li>[makeEwsRequestAsync](../../reference/outlook/Office.context.mailbox.md#makeewsrequestasyncdata-callback-usercontext) の呼び出し</li></ul>|
通常は、アドインで必要になる最小限のアクセス許可を指定する必要があります。アクセス許可は、マニフェストの **Permissions** 要素で宣言されます。詳細については、「[Outlook アドインのマニフェスト](../outlook/manifests/manifests.md)」を参照してください。セキュリティ上の問題については、「[Outlook アドインに関するプライバシー、アクセス許可、セキュリティ](../outlook/../../docs/develop/privacy-and-security.md)」を参照してください。


## その他のリソース

- [Outlook アドインのマニフェスト](../outlook/manifests/manifests.md)

- [Outlook API 要件セットについて](../../reference/outlook/tutorial-api-requirement-sets.md)
    
- [Outlook アドインに関するプライバシー、アクセス許可、セキュリティ](../outlook/../../docs/develop/privacy-and-security.md)
    
