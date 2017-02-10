
# <a name="package-your-add-in-using-visual-studio-to-prepare-for-publishing"></a>発行のための準備として Visual Studio を使用してアドインをパッケージ化する

Office アドイン パッケージには、アドインを発行する際に使用する XML ファイルが含まれます。 プロジェクトの Web アプリケーション ファイルは、別々に発行する必要があります。


## <a name="deploy-your-web-project-and-package-your-add-in-by-using-visual-studio-2015"></a>Visual Studio 2015 を使用して Web プロジェクトを展開しアドインをパッケージ化する



### <a name="to-deploy-your-web-project"></a>Web プロジェクトを展開するには


1. [ **ソリューション エクスプローラー**] で、アドイン プロジェクトのショートカット メニューを開き、 [ **発行**] を選択します。
    
    [**アドインの発行**] ページが表示されます。
    
2. **[現在のプロファイル]** ドロップダウン リストで、プロファイルを選択するか、または **[新規…]** を選択して新しいプロファイルを作成します。
    
     >**注** 発行プロファイルは、配置先となるサーバー、サーバーへのログオンに必要な資格情報、配置するデータベース、その他の配置オプションを指定します。

    [**新規...**] を選択した場合、[**発行プロファイルの作成**] ウィザードが表示されます。このウィザードを使用して、Microsoft Azure などの Web サイトをホストするプロバイダーから発行プロファイルをインポートするか、新しいプロファイルを作成するかして、次の手順でサーバー、資格情報、その他の設定を追加することができます。
    
    発行プロファイルのインポートまたは新しい発行プロファイルの作成の詳細については、「[発行プロファイルの作成](http://msdn.microsoft.com/en-us/library/dd465337.aspx#creating_a_profile)」を参照してください。
    
3. [ **アドインを発行する**] ページで、 [ **Web プロジェクトの配置**] リンクを選択します。
    
    The  **Publish Web** dialog box appears. For more information about using this wizard, see [How to: Deploy a Web Project using On-Click Publishing in Visual Studio](http://msdn.microsoft.com/en-us/library/dd465337.aspx).
    

### <a name="to-package-your-add-in"></a>アドインをパッケージ化するには


1. [ **アドインを発行する**] ページで、 [ **アドインのパッケージ化**] リンクをクリックします。
    
    [**Office アドインおよび SharePoint アドインを発行する**] ウィザードが表示されます。
    
2. [ **Web サイトがホストされている場所**] ドロップダウン リストで、アドイン のコンテンツ ファイルをホストする Web サイトの URL を選択または入力して、[ **完了**] を選択します。
    
    You have to specify an address that begins with the HTTPS prefix to complete this wizard. In general, using an HTTPS endpoint for your website is the best approach, but it is not required if you don't plan to publish your add-in to the Office Store. After the package is created, you can open the manifest in Notepad and replace the HTTPS prefix of your website with an HTTP prefix. For more information, see [Why do my add-ins have to be SSL-secured?](http://msdn.microsoft.com/en-us/library/jj591603#bk_q7). 
    
     >**注** Azure の Web サイトは自動的に HTTPS エンドポイントを提供します。

    Visual Studio は、アドインの発行に必要なファイルを生成して、発行の出力フォルダーを開きます。 
    
Office ストアへのアドインの提出を予定している場合は、 [ **検証チェックの実行**] リンクをクリックして、アドインの受け入れが阻害される問題点を識別します。アドインをストアに提出する前に、すべての問題に対処してください。

XML マニフェストを適切な場所にアップロードして[アドインを発行](../publish/publish.md)できるようになりました。XML マニフェストは  `OfficeAppManifests` フォルダーの `app.publish` にあります。たとえば、次のようになります。

 `%UserProfile%\Documents\Visual Studio 2015\Projects\MyApp\bin\Debug\app.publish\OfficeAppManifests`


## <a name="additional-resources"></a>その他のリソース



- [Office アドインを発行する](../publish/publish.md)
    
- [Office ストアに Office アドインと SharePoint アドインおよび Office 365 Web アプリを提出する](http://msdn.microsoft.com/library/ff075782-1303-4517-91cc-b3d730e9b9ae%28Office.15%29.aspx)
    
