# Outlook 用のマニフェストでアドイン コマンドを定義する

アドイン コマンドは、操作を実行する UI 要素を使用して、既定の Office UI をカスタマイズする簡単な方法を提供します。たとえば、リボンにカスタムのボタンを追加できます。 コマンドを作成する場合は、既存の作業ウィンドウ マニフェストに **[VersionOverrides](../../../reference/manifest/versionoverrides.md)** ノードを追加します。 

マニフェストが **VersionOverrides** 要素を含む場合、アドイン コマンドをサポートする Word、Excel、Outlook、PowerPoint のバージョンは、要素内の情報を使用して、アドインをロードします。 アドイン コマンドをサポートしていない以前のバージョンの Office 製品では、要素は無視されます。

クライアント アプリケーションが **VersionOverrides** ノードを認識する場合、アドインの名前はリボンに表示され、読み込み/作成ウィンドウには表示されません。 これらの場所に、アドインは表示されません。
 

## VersionOverrides ノード

[VersionOverrides](../../../reference/manifest/versionoverrides.md) 要素は、アドインによって実装されたアドイン コマンドに関する情報を格納するルート要素です。 マニフェストのスキーマ v1.1 以降でサポートされていますが、VersionOverrides v1.0 スキーマで定義されています。 

VersionOverrides 要素には、次の子要素が含まれます。

- [Description](../../../reference/manifest/description.md)
- [要件](../../../reference/manifest/requirements.md)
- [Hosts](../../../reference/manifest/hosts.md)
- [リソース](../../../reference/manifest/resources.md)

次の図は、アドイン コマンドの定義に使用する要素の階層を示しています。 

![マニフェスト内のアドイン コマンド要素の階層](../../../images/080da303-51c4-4882-b74a-7ba11517c0ad.png)

## Outlook アドイン コマンドのルールの変更点

次の変更は、マニフェストのルールに影響します。

- アクティブ化ルールは、各エントリ ポイント内に指定されるようになりました。
    
- [Rule](../../../reference/manifest/rule.md) 要素の **ItemIs** の属性が変更されました。 **ItemType** は、メッセージまたは AppointmentAttendee のどちらかにすることができます。 **FormType** の属性が削除されました。
    
- [Rule](../../../reference/manifest/rule.md) 要素の **ItemHasKnownEntity** の属性は、EntityType の文字列を受け入れるように更新されました。
    

## マニフェストのサンプル

Word、Excel、PowerPoint でアドイン コマンドを実装するサンプル マニフェストの場合は、「[Simple add-in commands sample](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/tree/master/Simple)」 (簡単なアドイン コマンドのサンプル) をご覧ください。

Outlook でアドイン コマンドを実装するサンプル マニフェストの場合は、「[Sample manifest file for an Outlook add-in](https://gist.github.com/mlafleur/95b7ac030bb7a7ae742527e85a36b095)」 (Outlook アドイン用のサンプル マニフェスト ファイル) をご覧ください。


## その他のリソース


- [Outlook のアドイン コマンド](../../outlook/add-in-commands-for-outlook.md)
    
- [Outlook アドインのマニフェスト](../../outlook/manifests/manifests.md)
    
- [Outlook アドイン コマンドのデモ サンプル](https://github.com/jasonjoh/command-demo)