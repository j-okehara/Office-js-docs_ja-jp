
# <a name="deploy-and-publish-your-office-add-in"></a>Office アドインを展開し、発行する

さまざまな方法を利用し、テスト目的またはユーザーに配布する目的で、Office アドインを展開できます。

|**メソッド**|**Use...**|
|:---------|:------------|
|[サイドロード](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)|開発プロセスの一環として、Windows、Office Online、iPad、Mac で実行されているアドインをテストします。|
|[Office 365 管理センター (プレビュー)](#office-365-admin-center-preview)|クラウド展開またはハイブリッド展開で組織内のユーザーにアドインを配布します。|
|[Office ストア]|ユーザーに配布する目的でアドインを公開します。|
|[SharePoint カタログ](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)|オンプレミス環境で、組織内のユーザーにアドインを配布します。|
|[Exchange サーバー](#outlook-add-in-deployment)|オンプレミス環境またはオンライン環境で、ユーザーに Outlook アドインを配布します。|

利用できるオプションは、対象とする Office ホストや作成するアドインの種類によって異なります。

>**注:** Office ストアにアドインを公開する予定であれば、[Office ストア検証ポリシー](https://msdn.microsoft.com/en-us/library/jj220035.aspx)に準拠していることを確認してください。たとえば、検証に合格するには、アドインは、定義したメソッドをサポートするすべてのプラットフォーム全体で機能する必要があります (詳細については、[セクション 4.12](https://dev.office.com/officestore/docs/validation-policies#4-apps-and-add-ins-behave-predictably) と「[Office アドインを使用できるホストおよびプラットフォーム](https://dev.office.com/add-in-availability)」のページを参照してください)。

## <a name="deployment-options-for-word-excel-and-powerpoint-add-ins"></a>Word、Excel、PowerPoint アドインの展開オプション

| 拡張点            | サイドロード | Office 365 管理センター (プレビュー) |Office ストア| SharePoint カタログ*  |
|:----------------|:-----------:|:------------------:|:-------------------------------:|:------------:|
| コンテンツ         | X           | X                  | X                               | X|
| 作業ウィンドウ       | X           | X                  | X                               | X|
| コマンド           | X           | X                  | X                               |  |

&#42; SharePoint カタログは、Office 2016 for Mac をサポートしません。

## <a name="deployment-options-for-outlook-add-ins"></a>Outlook アドインの展開オプション

| 拡張点     | サイドロード | Exchange サーバー | Office ストア |
|:---------|:-----------:|:---------------:|:------------:|
| メール アプリ | X           | X               | X            |
| コマンド  | X           | X               | X            |


エンド ユーザーがアドインを取得、挿入、実行する方法については、「[Office アドインの使用を開始する](https://support.office.com/en-ie/article/Start-using-your-Office-Add-in-82e665c4-6700-4b56-a3f3-ef5441996862?ui=en-US&rs=en-IE&ad=IE)」を参照してください。

## <a name="office-365-admin-center-preview-deployment"></a>Office 365 管理センター (プレビュー) の展開

Office 365 管理センターでは、管理者が Word、Excel、PowerPoint のアドインを、組織内のユーザーやグループに簡単に展開できます。管理センター経由で展開されたアドインは、ユーザーがすぐに Office アプリケーションで利用できるようになります。クライアントの構成は必要ありません。内部アドインも、ISV から提供されるアドインも、管理センターで展開することができます。

現在、管理センターは次のシナリオをサポートしています。

- 新しいアドインおよび更新されたアドインの個人、グループ、組織への一元展開。
- Windows、Office Online を含む複数のプラットフォームのサポート (Mac は準備中)。
- 英語および世界各国のテナントへの展開。
- クラウド ホスト型のアドインの展開。
- Office アプリケーションの起動時に自動的にインストール。
- ファイアウォール内でホストされるアドイン URL。
- Office ストア アドインの展開 (準備中)。

<!--
The admin center also includes a pre-deployment validation checking service.
-->

アドインの展開シナリオにおける今後の投資は Office 365 管理センターに焦点を当てていきます。組織が前提条件を満たしているのであれば、管理センターを使ってアドインを組織に展開することをお勧めします。

### <a name="prerequisites-for-admin-center-deployment"></a>管理センターの展開の前提条件 

管理センターを介してアドインを展開するには、組織が次の基準を満たしている必要があります。

- ユーザーが Office 2016 ビルド 7070 以降を実行している。
- ユーザーが職場または学校のアカウントで Office 2016 にサインインしている。
- 組織で、Azure Active Directory (AD Azure) の ID サービスを使用している。

管理センターは、以下をサポートしていません。

- Office 2013 の Word、Excel、PowerPoint を対象にしたアドイン。
- オンプレミスのディレクトリ サービス。
- SharePoint アドインの展開。
- Office Online Server へのアドインの展開。
- COM/VSTO アドインの展開。

SharePoint アドインまたは Office 2013 を対象とするアドインを展開するには、[SharePoint アドイン カタログ](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)を使います。

>**重要!**SharePoint アドイン カタログでは、[アドイン コマンド](../design/add-in-commands.md)など、アドイン マニフェストの [VersionOverride](../../reference/manifest/versionoverrides.md) ノードに実装されているアドイン機能はサポートされていません。 

COM/VSTO アドインを展開するには、ClickOnce または Windows インストーラーを使います。詳細については、「[Office ソリューションの配置](https://msdn.microsoft.com/en-us/library/bb386179.aspx)」を参照してください。

## <a name="sharepoint-catalog-deployment"></a>SharePoint カタログの展開

SharePoint アドインのカタログは、Word、Excel、PowerPoint のアドインをホストするために作成可能な特別なサイト コレクションです。SharePoint カタログは、アドイン コマンドを含むマニフェストの VersionOverrides ノードに実装されている新しいアドインの機能をサポートしていないため、可能な場合は管理センター (プレビュー) 経由の一元展開を行うことをお勧めします。SharePoint カタログによって展開したアドイン コマンドは、既定では作業ウィンドウで開かれます。

オンプレミス環境でアドインを展開する場合は、SharePoint カタログを使用します。詳細については、「[作業ウィンドウ アドインとコンテンツ アドインを SharePoint カタログに発行する](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)」を参照してください。

> **注:** SharePoint カタログは、Office 2016 for Mac をサポートしません。Office アドインを Mac クライアントに展開するには、それを [Office ストア]に提出する必要があります。 

## <a name="outlook-add-in-deployment"></a>Outlook アドインの展開

Azure AD の ID サービスを使用しないオンプレミス環境およびオンライン環境では、Exchange サーバー経由で Outlook アドインを展開することができます。 

Outlook アドインの展開には以下が必要です。

- Office 365、Exchange Online、または Exchange Server 2013 以降
- Outlook 2013 以降

アドインをテナントに割り当てるには、Exchange 管理センターを使用して、ファイルまたは URL から直接マニフェストをアップロードするか、または Office ストアからアドインを追加します。アドインを個々のユーザーに割り当てるには、Exchange PowerShell を使用する必要があります。詳細については、TechNet の「[組織の Outlook 用アプリをインストールまたは削除する](https://technet.microsoft.com/en-us/library/jj943752(v=exchg.150).aspx)」を参照してください。


## <a name="additional-resources"></a>追加リソース

- [テスト用に Outlook アドインを展開してインストールする](../outlook/testing-and-tips.md) 
- [Office ストアにアドインと Web アプリを提出する][Office ストア]
- [Office アドインの設計ガイドライン](../design/add-in-design)
- [効果的な Office ストア アドインを作成する](https://msdn.microsoft.com/en-us/library/jj635874.aspx)
- [Office アドインでのユーザー エラーのトラブルシューティング](../testing/testing-and-troubleshooting.md)

[Office ストア]: http://msdn.microsoft.com/library/ff075782-1303-4517-91cc-b3d730e9b9ae%28Office.15%29.aspx
[Office Add-in host and platform availability]: http://dev.office.com/add-in-availability
