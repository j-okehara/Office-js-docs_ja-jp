
# Office アドインを実行するための要件


この記事では、Office アドインを実行するためのソフトウェアとデバイスの要件について説明します。

>**注:**現時点での Office アドインのサポート状況について、概要は「[Office アドインを使用できるホストおよびプラットフォーム](http://dev.office.com/add-in-availability)」ページを参照してください。 


## サーバーの要件

Office アドインをインストールおよび実行できるようにするには、まずアドインの UI とコードのマニフェストと Web ページ ファイルを、適切なサーバーの場所に展開する必要があります。

すべての種類のアドイン (コンテンツ、Outlook、作業ウィンドウの、アドインとアドイン コマンド) で、アドインの Web ページ ファイルを Web サーバーや [Microsoft Azure](../publish/host-an-office-add-in-on-microsoft-azure.md) などの Web ホスティング サービスに展開する必要があります。


 >**注：** Visual Studio でアドインの開発とデバッグを行う場合、Visual Studio は IIS Express を使ってローカルにアドインの Web ページ ファイルを展開して実行するため、追加の Web サーバーは不要になります。同様に、ブラウザーで Napa を使って開発とデバッグを行う場合は、Napa へのサインインに使うアカウントに関連付けられたストレージからアドインの Web ページのファイルを展開して実行します。

サポートされている Office ホスト アプリケーション (Access Web アプリ、Word、Excel、PowerPoint、または Project) のコンテンツ アドインと作業ウィンドウ アドインでは、アドインの XML マニフェスト ファイルをアップロードするために、SharePoint の [アドイン カタログ](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)も必要になります。

Outlook アドインをテストおよび実行するには、ユーザーの Outlook 電子メール アカウントが、Office 365、Exchange Online、またはオンプレミスのインストールから使用できる Exchange 2013 以降のバージョン上に存在する必要があります。ユーザーまたは管理者は、サーバー上に Outlook アドインのマニフェスト ファイルをインストールします。

 >**注:** Outlook の POP および IMAP 電子メール アカウントは、Office アドイン をサポートしていません。




## クライアントの要件:Windows のデスクトップとタブレット

Windows ベースのデスクトップ、ノート PC、または タブレット デバイス上で実行されるサポート対象の Office デスクトップ クライアントまたは Web クライアント向けの Office アドインを開発するには、以下のソフトウェアが必要です。


- Windows x86 および x64 デスクトップおよび Surface Pro などのタブレット:

    - Windows 7 以降のバージョンで実行している Office 2013 以降のバージョンの、32 ビットまたは 64 ビット バージョン。

    - Excel 2013、Outlook 2013、PowerPoint 2013、Project Professional 2013、Project 2013 SP1、Word 2013、またはそれ以降の Office クライアントのバージョン (特にこれらの Office デスクトップ クライアントを対象として Office アドインをテストまたは実行する場合)。Office デスクトップ クライアントはオンプレミスでインストールすることも、クイック実行によってクライアント コンピューターにインストールすることもできます。

- Internet Explorer 9 以降。インストールが必要ですが、既定のブラウザーである必要はありません。Office アドインをサポートするために、ホストとして機能する Office クライアントは Internet Explorer 9 以上を構成しているブラウザー コンポーネントを使用します。

- 次のいずれかの既定のブラウザー: Internet Explorer 9、Safari 5.0.6、Firefox 5、Chrome 13、またはこれらのブラウザーのそれ以降のバージョン。

- メモ帳などの HTML および JavaScript エディター、 [Visual Studio および Microsoft Developer Tools](https://www.visualstudio.com/features/office-tools-vs)、またはサードパーティの Web 開発ツール。


## クライアントの要件: OS X デスクトップ

Outlook for Mac は Office 365 に付属していて、Outlook アドインをサポートします。Outlook アドインを Outlook for Mac で実行するための要件は、Outlook for Mac そのものの要件と同じです。オペレーティング システムは、少なくとも OS X v10.10 "Yosemite" である必要があります。Outlook for Mac はレイアウト エンジンとして WebKit を使用して、アドイン ページを表示するので、追加のブラウザーの依存関係はありません。

次は、Office アドインをサポートする Office for Mac の最小クライアント バージョンです。
- Word for Mac バージョン 15.18 (160109) 
- Excel for Mac バージョン 15.19 (160206) 
- PowerPoint for Mac バージョン 15.24 (160614)

## クライアントの要件:Office Online Web クライアントと SharePoint のブラウザー サポート

Internet Explorer 9、Chrome 13、Firefox 5、Safari 5.0.6、またはこれらのブラウザーの以降のバージョンなど ECMAScript 5.1、HTML5、および CSS3 をサポートする任意のブラウザー。


## クライアントの要件: Windows 以外のスマートフォンおよび タブレット

特に、スマートフォンや Windows 以外のタブレット デバイス上のブラウザーで動作する デバイス用 OWA と Outlook Web App の場合、Outlook アドインをテストおよび実行するのに以下のソフトウェアが必要です。


| ホスト アプリケーション | デバイス | オペレーティング システム | Exchange アカウント | モバイル ブラウザー |
|:-----|:-----|:-----|:-----|:-----|
|OWA for Android|Android スマートフォン。技術的には、「 [Android OS](https://developer.android.com/guide/practices/screens_support.html)」によって「小型」または「標準」に分類されるデバイス。|Android 4.4 KitKat 以降|Office 365 for Business または Exchange Online の最新の更新プログラムが対象|Android 用のネイティブ アドイン、ブラウザーは適用外|
|OWA for iPad|iPad 2 以降|iOS 6 以降|Office 365 for Business または Exchange Online の最新の更新プログラムが対象|iOS 用のネイティブ アドイン、ブラウザーは適用外|
|OWA for iPhone|iPhone 4S 以降|iOS 6 以降|Office 365 for Business または Exchange Online の最新の更新プログラムが対象|iOS 用のネイティブ アドイン、ブラウザーは適用外|
|Outlook Web App|iPhone 4 以降、iPad 2 以降、iPod Touch 4 以降|iOS 5 以降|Office 365、Exchange Online、または Exchange Server 2013 以降の社内型が対象|Safari|


## その他のリソース

- [Office アドイン プラットフォームの概要](../../docs/overview/office-add-ins.md)
- [Office アドインを使用できるホストおよびプラットフォーム](http://dev.office.com/add-in-availability)

