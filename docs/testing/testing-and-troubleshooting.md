
# <a name="troubleshoot-user-errors-with-office-add-ins"></a>Office アドインでのユーザー エラーのトラブルシューティング

時折、ユーザーは開発した Office アドインの問題に遭遇することがあります。たとえば、アドインが読み込みに失敗したり、アクセスできないなどです。この記事の情報は、ユーザーが Office アドインを使用する際に遭遇する一般的な問題を解決するために用いることができます。 

また、 [Fiddler](http://www.telerik.com/fiddler) を使用して、アドインの問題を特定してデバッグすることもできます。

ユーザーの問題を解決した後、 [Office ストアでカスタマー レビューに直接返信することができます](https://msdn.microsoft.com/library/jj635874.aspx)。

## <a name="common-errors-and-troubleshooting-steps"></a>一般的なエラーとトラブルシューティングの手順

次の表は、ユーザーが遭遇する可能性がある一般的なエラー メッセージとエラーを解決するためにユーザーが実行できる手順を示しています。



|**エラー メッセージ**|**解決策**|
|:-----|:-----|
|アプリのエラー: カタログに到達できませんでした|ファイアウォールの設定を確認します。「カタログ」は、Office ストアを指します。このメッセージは、ユーザーが Office ストアにアクセスできないことを示します。|
|アプリのエラー: このアプリは起動できませんでした。このダイアログを閉じて問題を無視するか、 [再起動] をクリックしてやり直してください。|Office の最新の更新プログラムがインストールされていることを確認するか、[Office 2013 更新プログラム ](https://support.microsoft.com/en-us/kb/2986156/)をダウンロードします。|
|エラー: オブジェクトがプロパティまたはメソッド 'defineProperty' をサポートしていません|Internet Explorerが互換モードで実行されていないことを確認します。 [ツール] >  **[互換表示設定]** に移動します。|
|ブラウザーのバージョンがサポートされていないため、アプリを読み込めませんでした。サポートされているブラウザーのバージョンの一覧についてはここをクリックしてください。|ブラウザーが HTML5 のローカル ストレージをサポートしていることを確認するか、Internet Explorer の設定をリセットします。サポートされているブラウザーの詳細については、「[Office アドインを実行するための要件](../../docs/overview/requirements-for-running-office-add-ins.md)」を参照してください。|

## <a name="outlook-add-in-doesnt-work-correctly"></a>Outlook アドインが正常に機能しない

Windows で実行している Outlook アドインが正常に機能しない場合は、Internet Explorer でスクリプトのデバッグを有効にしてみてください。 


- [ツール] >  **[インターネット オプション]** > **[詳細]** に移動します。
    
- **[参照]** で、 **[スクリプトのデバッグを無効にする (Internet Explorer)]** と **[スクリプトのデバッグを無効にする (その他)]** の各チェックボックスをオフにします。
    
これらの設定は、問題のトラブルシューティングを行う場合にのみチェックボックスをオフにすることをお勧めします。チェックボックスをオフにしたままにすると、参照時にメッセージが表示されます。問題が解決したら、 **[スクリプトのデバッグを無効にする (Internet Explorer)]** と **[スクリプトのデバッグを無効にする (その他)]** の各チェックボックスをオンにしてください。


## <a name="add-in-doesnt-activate-in-office-2013"></a>Office 2013 でアドインがアクティブにならない

ユーザーが次の手順を実行したときに、アドインがアクティブにならない場合があります。


1. Office 2013 で自分の Microsoft アカウントでサインインする。
    
2. 自分の Microsoft アカウントの 2 段階検証を有効にする。
    
3. アドインを挿入しようとする際に、メッセージに従って ID の確認を行う。
    
Office の最新の更新プログラムがインストールされていることを確認するか、[Office 2013 更新プログラム](https://support.microsoft.com/en-us/kb/2986156/)をダウンロードしてください。

## <a name="add-in-doesnt-load-in-task-pane-or-other-issues-with-the-add-in-manifest"></a>アドインが作業ウィンドウで読み込まれない、または他のアドイン マニフェストの問題

「[マニフェストの問題を検証し、トラブルシューティングする](troubleshoot-manifest.md)」を参照して、アドインのマニフェストの問題をデバッグしてください。

## <a name="additional-resources"></a>追加リソース



- [Office Online でアドインをデバッグする](../testing/debug-add-ins-in-office-online.md)
    
- [iPad または Mac で Office アドインをサイドロードする](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)
    
- [iPad と Mac で Office アドインをデバッグする](../testing/debug-office-add-ins-on-ipad-and-mac.md)
    
- [Visual Studio での Office アドインの作成とデバッグ](../../docs/get-started/create-and-debug-office-add-ins-in-visual-studio.md)
    
- [テスト用に Outlook アドインを展開してインストールする](../outlook/testing-and-tips.md)
    
- [マニフェストの問題を検証し、トラブルシューティングする](troubleshoot-manifest.md)
    
