# <a name="uidisplaydialogasync-method"></a>UI.displayDialogAsync メソッド

Office ホストでダイアログ ボックスを表示します。 

## <a name="requirements"></a>要件

|ホスト|導入バージョン|最終変更バージョン|
|:---------------|:--------|:----------|
|Word、Excel、PowerPoint|1.1|1.1|
|Outlook|メールボックス 1.4|メールボックス 1.4|

このメソッドは、DialogAPI [要件セット](../../docs/overview/specify-office-hosts-and-api-requirements.md)で使用できます。DialogAPI 要件セットを指定するには、マニフェストで次のようにします。

```xml
 <Requirements> 
   <Sets DefaultMinVersion="1.1"> 
     <Set Name="DialogAPI"/> 
   </Sets> 
 </Requirements> 

```

実行時にこの API を検出するには、次のコードを使用します。

```js
 if (Office.context.requirements.isSetSupported('DialogAPI', 1.1)) 
    {  
         // Use Office UI methods; 
    } 
 else 
     { 
         // Alternate path 
     } 
```



### <a name="supported-platforms"></a>サポートされるプラットフォーム
現在、DialogAPI 要件セットは次のプラットフォームでサポートされています。

  - Windows デスクトップ版 Office 2013 (ビルド 15.0.4855.1000 以降)
  - Windows デスクトップ版 Office 2016 (ビルド 16.0.6741.0000 以降)
  - Office for IPad (ビルド 1.22 以降)
  - Office for Mac (ビルド 15.20 以降) 

その他のプラットフォームは近日公開します。 

## <a name="syntax"></a>構文

```js
Office.context.ui.displayDialogAsync(startAddress, options, callback);
```
##<a name="examples"></a>例

**displayDialogAsync** メソッドを使用する簡単な例については、GitHub の「[Office アドイン ダイアログ API の例](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example/)」を参照してください。

認証シナリオを示す例については、以下を参照してください。

- [Microsoft Graph ASP.Net の PowerPoint アドイン グラフの挿入](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)
- [Office アドイン Auth0](https://github.com/OfficeDev/Office-Add-in-Auth0)
- [Excel アドイン ASP.NET QuickBooks](https://github.com/OfficeDev/Excel-Add-in-ASPNET-QuickBooks)
- [Office アドインのサーバー認証サンプル ASP.net MVC 用](https://github.com/dougperkes/Office-Add-in-AspNetMvc-ServerAuth/tree/Office2016DisplayDialog)
- [Office アドイン Office 365 のクライアント認証 AngularJS 用](https://github.com/OfficeDev/Word-Add-in-AngularJS-Client-OAuth)


 
## <a name="parameters"></a>パラメーター

| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|startAddress|string|ダイアログ ボックスで最初に開く HTTPS(TLS) URL を受け取ります。 <ul><li>最初のページは、親ページと同じドメインにある必要があります。初期ページが読み込まれた後、他のドメインに移動できます。</li><li>[office.context.ui.messageParent](officeui.messageparent.md) を呼び出すすべてのページも、親ページと同じドメインに存在する必要があります。</li></ul>|
|options|object|省略可能です。ダイアログの動作を定義する options オブジェクトを指定できます。|
|callback|object|ダイアログ作成の試行を処理するコールバック メソッドを指定できます。|
    
### <a name="configuration-options"></a>構成オプション
ダイアログ ボックスでは次の構成オプションを使用できます。


| プロパティ     | 型   |説明|
|:---------------|:--------|:----------|
|**width**|object|省略可能。現在の表示のパーセンテージとして、ダイアログ ボックスの幅を定義します。既定値は 80% です。最小解像度は 250 ピクセルです。|
|**height**|object|省略可能。現在の表示のパーセンテージとして、ダイアログ ボックスの高さを定義します。既定値は 80% です。最小解像度は 150 ピクセルです。|
|**displayInIframe**|object|省略可能。ダイアログ ボックスを Office Online クライアントの IFrame 内に表示する必要があるかどうかを指定します。デスクトップ クライアントでは、この設定は無視されます。指定可能な値は次のいずれかです。<ul><li>False (既定値) - ダイアログは、新しいブラウザー ウィンドウ (ポップアップ) として表示されます。IFrame に表示できない認証ページに推奨されます。 </li><li>True - ダイアログは、IFrame のフローティング オーバーレイとして表示されます。これは、ユーザー エクスペリエンスとパフォーマンスに最適です。</li>|


## <a name="callback-value"></a>コールバック値
_callback_ パラメーターに渡した関数が実行されると、[AsyncResult](../../reference/shared/asyncresult.md) オブジェクトを受け取ります。このオブジェクトには、コールバック関数の唯一のパラメーターからアクセスできます。

**displayDialogAsync** メソッドに渡されるコールバック関数では、**AsyncResult** オブジェクトのプロパティを使用して次の情報を返すことができます。



|**プロパティ**|**使用目的**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|[Dialog](../../reference/shared/officeui.dialog.md) オブジェクトにアクセスします。|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|操作の成功または失敗を判断します。|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|操作が失敗した場合、エラーに関する情報を提供する [Error](../../reference/shared/error.md) オブジェクトにアクセスします。|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|ユーザー定義のオブジェクトまたは値を _asyncContext_ パラメーターとして渡した場合、そのオブジェクトまたは値にアクセスします。|

### <a name="errors-from-displaydialogasync"></a>displayDialogAsync のエラー

一般的なプラットフォームやシステムのエラーのほかに、**displayDialogAsync** の呼び出しに特有の次のエラーがあります。

|**コード番号**|**意味**|
|:-----|:-----|
|12004|`displayDialogAsync` に渡される URL のドメイン は信頼されていません。ドメインは、ホスト ページと同じドメイン (プロトコルとポート番号を含む) にするか、またはアドイン マニフェストの `<AppDomains>` セクションで登録する必要があります。|
|12005|`displayDialogAsync` に渡される URL には HTTP プロトコルを使用します。HTTPS が必要です。(Office の一部のバージョンでは、12004 で返されるのと同じエラー メッセージが 12005 でも返されます。)|
|12007|ダイアログ ボックスは、作業ウィンドウで既に開いています。作業ウィンドウ アドインで一度に開けるダイアログ ボックスは 1 つだけです。|



## <a name="design-considerations"></a>設計上の考慮事項
ダイアログ ボックスの設計には次のような考慮事項が適用されます。

- Office アドインが開くことのできるダイアログ ボックスは、一度に 1 つだけです。
- ユーザーは、すべてのダイアログ ボックスを移動およびサイズ変更できます。
- すべてのダイアログ ボックスは、画面の中央に開かれます。
- ダイアログ ボックスは、ホスト アプリケーションの前面に、作成された順序で表示されます。

ダイアログ ボックスは次のような場合に使用します。

- ユーザーの資格情報を収集する認証ページを表示します。
- ShowTaspane または ExecuteAction コマンドから、エラー/進行状況/入力画面を表示します。
- ユーザーがタスクの完了に利用できる表示領域を一時的に拡大します。

ドキュメントとの対話にはダイアログ ボックスを使用しないでください。代わりに作業ウィンドウを使用してください。 

ダイアログ ボックスの作成に使用できるデザイン パターンについては、GitHub の Office アドイン UX デザイン パターン リポジトリの「[クライアント ダイアログ](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/Patterns/Client_Dialog.md)」を参照してください。
