
# iPad 用の Office アドインを設計する


次の表に、Office for iPad で実行する Office アドインを設計する際に実行するタスクの一覧を示します。


|**タスク**|**説明**|**資料**|
|:-----|:-----|:-----|
|アドインを更新して、Office.js バージョン 1.1 をサポートします。|Office アドイン プロジェクトで使用する JavaScript ファイル (Office.js ファイルとアプリに固有の .js ファイル) とアドイン マニフェスト検証ファイルをバージョン 1.1 に更新します。|[JavaScript API for Office の変更点](../../reference/what's-changed-in-the-javascript-api-for-office.md)|
|UI デザインのベスト プラクティスを適用します。|アドイン UI を iOS エクスペリエンスとシームレスに統合します。|[iOS の設計](https://developer.apple.com/library/ios/documentation/UserExperience/Conceptual/MobileHIG/)|
|アドイン デザインのベスト プラクティスを適用します。|アドインが明確な価値を提供し、魅力的であり、一貫して機能することを確認します。|[Office アドイン開発のベスト プラクティス](../../docs/overview/add-in-development-best-practices.md)|
|タッチ用にアドインを最適化します。|マウスとキーボードに加え、タッチ入力に対して、UI が素早く応答するようにします。|[UX 設計原則を適用する](https://msdn.microsoft.com/ja-jp/library/mt590883.aspx#Anchor_3)|
|アドインを無料にします。|iPad 上の Office は、ユーザー数を拡大して、サービスを促進できるチャネルです。これらの新しいユーザーは、お客様になる可能性があります。|[検証ポリシー 10.8](http://msdn.microsoft.com/ja-jp/library/cd90836a-523e-42f5-ab02-5123cdf9fefe%28Office.15%29.aspx)|
|アドインを商目的で使用しないようにします。|アドインには、アプリ内購入、試用版の提供、有料版へのアップセルを目的とする UI、またはユーザーが他のコンテンツ、アプリ、アドインを購入または取得できるすべてのオンライン ストアへのリンクが含まれていてはいけません。またプライバシー ポリシーと使用条件のページにも、商用の UI またはストアへのリンクがないことが必要です。|[検証ポリシー 3.4](http://msdn.microsoft.com/ja-jp/library/cd90836a-523e-42f5-ab02-5123cdf9fefe%28Office.15%29.aspx)|
|アドインを Office ストアに再送信します。|販売者ダッシュボードで、**[このアドインを iPad の Office アドイン カタログで利用できる状態にする]** チェック ボックスをオンにして、[Apple ID] ボックスに Apple 開発者 ID を入力します。[Office ストア アプリケーション プロバイダー契約](https://sellerdashboard.microsoft.com/Assets/Content/Agreements/en-US/Office_Store_Seller_Agreement_20120927.md)を確認して、契約を十分に理解します。|[Office ストアに Office アドインと SharePoint アドインおよび Office 365 Web アプリを提出する](http://msdn.microsoft.com/ja-jp/library/ff075782-1303-4517-91cc-b3d730e9b9ae%28Office.15%29.aspx)|
他のプラットフォームで実行されている Office アプリケーション用にアドインをそのまま保持することができます。また、アドインが実行されているブラウザーとデバイスに基づく別の UI も提供できます。iPad 上でアドインが実行されているかどうかを検出するためには、次の API を使用できます。 

- var isTouchEnabled = [Office.context.touchEnabled](../../reference/shared/office.context.touchenabled.md)
    
- var allowCommerce = [Office.context.commerceAllowed](../../reference/shared/office.context.commerceallowed.md)
    

## iOS および Mac 用 Office アドイン開発のベスト プラクティス

iOS 上で実行するアドインを開発するための次のベスト プラクティスを適用します。


-  **アドインの開発に Visual Studio を使用する。**
    
    If you develop your add-in with Visual Studio, you can [set breakpoints and debug its code](../get-started/create-and-debug-office-add-ins-in-visual-studio.md#Test) in an Office host application running on Windows, before sideloading your add-in on the iPad or Mac. Because an add-in that runs in Office for iOS or Office for Mac supports the same APIs as an add-in running in Office for Windows, your add-in's code should run the same way on both platforms.
    
-  **アドインのマニフェストまたはランタイム チェックを使用して API の要件を指定する。**
    
    When you specify API requirements in your add-in's manifest, Office will determine if the host application supports those API members. If the API members are available in the host, then your add-in will be available in that host application. Alternatively, you can perform a runtime check to determine if a method is available in the host before using it in your add-in. Runtime checks ensure that your add-in is always available in the host, and provides additional functionality if the methods are available. For more information, see [Specify Office hosts and API requirements](../../docs/overview/specify-office-hosts-and-api-requirements.md).
    
一般的なアドイン開発のベスト プラクティスについては、「[Office アドイン開発のベスト プラクティス](../../docs/overview/add-in-development-best-practices.md)」を参照してください。


## その他の技術情報
<a name="bk_addresources"></a>


- [iPad または Mac で Office アドインをサイドロードする](../../docs/testing/sideload-an-office-add-in-on-ipad-and-mac.md)
    
- [iPad と Mac で Office アドインをデバッグする](../../docs/testing/debug-office-add-ins-on-ipad-and-mac.md)
    

