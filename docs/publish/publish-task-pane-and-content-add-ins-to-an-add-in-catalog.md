
# <a name="publish-task-pane-and-content-add-ins-to-a-sharepoint-catalog"></a>作業ウィンドウ アドインとコンテンツ アドインを SharePoint カタログに発行する

>**重要!**SharePoint のアドイン カタログでは、アドイン コマンドなど、[アドイン マニフェスト](../overview/add-in-manifests.md)の VersionOverride ノードに実装されているアドイン機能はサポートされていません。 

>クラウド環境またはハイブリッド環境をターゲットにしている場合は、[管理センター (プレビュー)](publish/publish.md#office-365-admin-center-preview-deployment) 経由の**一元展開**によって、アドインを発行することをお勧めします。

アドイン カタログは、Office アドインと SharePoint アドインのドキュメント ライブラリをホストする SharePoint Web アプリケーションまたは SharePoint Online テナンシーの専用サイト コレクションです。管理者は、組織の Office アドイン マニフェスト ファイルをアドイン カタログにアップロードできます。管理者がアドイン カタログを信頼できるカタログとして登録すると、ユーザーは Office クライアント アプリケーションで挿入 UI からアドインを挿入できます。

SharePoint カタログは Office 2016 for Mac ではサポートされていません。Office アドインを Mac クライアントに展開するには、それを [Office ストア](http://msdn.microsoft.com/library/ff075782-1303-4517-91cc-b3d730e9b9ae%28Office.15%29.aspx)に提出する必要があります。   

## <a name="to-set-up-an-add-in-catalog-on-sharepoint"></a>SharePoint 上でアドイン カタログをセットアップするには

1. **中央管理サイト**を参照 (**[スタート]** > **[すべてのプログラム]** > **[Microsoft SharePoint 2013 製品]** > **[SharePoint 2013 サーバーの全体管理]**の順にクリック) します。
    
2. 左側の作業ウィンドウで、 [ **アドイン**] を選択します。
    
3. [ **アドイン**] ページの [ **アドイン管理**] で、[ **アドイン カタログの管理**] を選択します。
    
4. [ **アドイン カタログの管理**] ページの  **Web アプリケーション セレクター**で正しい Web アプリケーションが選択されていることを確認します。
    
5. [ **サイトの設定の表示**] を選択します。
    
6. [ **サイトの設定**] ページで、[ **サイト コレクション管理者**] を選択してサイト コレクション管理者を指定してから、[ **OK**] を選択します。
    
7. ユーザーにサイト アクセス許可を付与するには、[ **サイトの権限**] を選択してから、[ **アクセス許可の付与**] を選択します。
    
8. [ **アプリ カタログ サイトの共有**] ダイアログ ボックスで、1 人以上のサイト ユーザーを指定して、それらに適切なアクセス許可を設定し、必要に応じて他のオプションを設定してから、[  **共有**] を選択します。
    
9. アドインを Office アドイン アドイン カタログに追加するには、[ **Office アドイン**] を選択します。

## <a name="to-set-up-an-add-in-catalog-on-office-365"></a>Office 365 でアドイン カタログをセットアップするには

1. [Office 365 管理センター] ページで、 **[管理]**、 **[SharePoint]** の順にクリックします。
    
2. 左側の作業ウィンドウで、[ **アドイン**] を選択します。
    
3. [ **アドイン**] ページで、[ **アドイン カタログ**] を選択します。
    
4. [ **アドイン カタログ サイト**] ページで、[ **OK**] を選択して既定のオプションを受け入れ、新しいアドイン カタログ サイトを作成します。
    
5. [ **アドイン カタログ サイト コレクションの作成**] ページで、アドイン カタログ サイトのタイトルを指定します。
    
6. Web サイト アドレスを指定します。
    
7. [ **記憶域のクォータ**] を可能な限り小さい値に設定します (現在は 110)。このサイト コレクションにはアドイン パッケージだけをインストールしますが、パッケージは非常に小さなものです。
    
8. [ **サーバー リソース クォータ**] を 0 (ゼロ) に設定します。(サーバー リソース クォータは、パフォーマンスが低いサンドボックス ソリューションのスロットルに関連していますが、このアドインのカタログ サイトにはサンドボックス ソリューションをインストールしません。)
    
9. [ **OK**] をクリックします。
    
アドインをアドイン カタログ サイトに追加するために、作成したばかりのサイトを参照します。左側のナビゲーション ウィンドウで、 [ **Office アドイン**] を選択してから、Office アドイン マニフェスト ファイルをアップロードするために、[ **新しいアドイン**] を選択します。    

## <a name="publish-to-an-add-in-catalog"></a>アドイン カタログへの発行


1. アドイン カタログを参照します。

    1- SharePoint サーバーの全体管理メイン ページを開きます。
    
    2- **[アドイン]** を選択します。
    
    3- **[アドイン カタログの管理]** を選択します。
    
    4- 表示されたリンクを選択し、左側のナビゲーション バーで **[Office アドイン]** を選択します。
    
2. **[新しいアイテムの追加]** リンクを選択します。
    
3. **[参照]** を選択し、アップロードする [[マニフェスト]](../../docs/overview/add-in-manifests.md) を指定します。
    
    このカタログのコンテンツおよび作業ウィンドウのアドインが **[Office アドイン]** ダイアログ ボックスから使用できるようになりました。これらにアクセスするには、**[挿入]** タブで **[個人用アドイン]** を選択して、**[自分の所属組織]** を選択します。
    
アドイン マニフェストを Office アドイン カタログにアップロードすると、ユーザーは次の操作を行ってアドインにアクセスできます。


1. Office アプリケーションで、 **[ファイル]**  >  **[オプション]**  >  **[セキュリティ センター]**  >  **[セキュリティ センターの設定]**  >  **[信頼できるアドイン カタログ]** の順に移動します。
    
2. アドイン カタログの  _親 SharePoint サイト コレクション_ の URL を指定します。たとえば、Office アドイン カタログの URL が次のような場合:
    
    `https:// _domain_ /sites/ _AddinCatalogSiteCollection_ /AgaveCatalog`
    
    単に親サイト コレクションの URL を指定します:
    
    `https:// _domain_ /sites/ _AddinCatalogSiteCollection_`
    
3. Office アプリケーションを閉じ、もう一度開きます。アドイン カタログが  **[Office アドイン]** ダイアログ ボックスに表示されます。
    
または、管理者はグループ ポリシーを使用することにより SharePoint 上の Office アドイン カタログを指定できます。詳細については、TechNet の「[Office アドインの概要](https://technet.microsoft.com/en-us/library/jj219429.aspx)」にある「グループ ポリシーを使用して、ユーザーが Office アドインをインストールおよび使用する方法を管理する」のセクションをご参照ください。

