
# <a name="dialog-api-requirement-sets"></a>ダイアログ API の要件セット

要件セットは、API メンバーの名前付きグループです。Office アドインは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office ホストがアドインに必要な API をサポートしているかどうかを判別します。詳しくは、「[Office のホストと API の要件を指定する](../../docs/overview/specify-office-hosts-and-api-requirements.md)」をご覧ください。

Office アドインは Office の複数のバージョンで機能します。次の表は、ダイアログ API の要件セット、その要件セットをサポートする Office ホスト アプリケーション、Office アプリケーションのビルド番号またはバージョン番号の一覧です。

|  要件セット  |  Office 2013 for Windows | Office 2016 for Windows*   |  Office 2016 for iPad  |  Office 2016 for Mac  | Office Online  |  Office Online Server  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|
| DialogApi 1.1  | ビルド 15.0.4855.1000 以降 | バージョン 1602 (ビルド 6741.0000) 以降 | 1.22 以降 | 15.20 以降| 準備中。 | バージョン 1608 (ビルド 7601.6800) 以降|

>**注:**MSI からインストールされた Office 2016 のビルド番号は、16.0.4266.1001 です。ダイアログ API を使用するには、Office の更新プログラムを実行して、最新バージョンを取得してください。 

バージョン、ビルド番号、および Office Online Server の詳細については以下を参照してください。

- [Office 365 クライアントの更新プログラム チャネル リリースのバージョン番号およびビルド番号](https://technet.microsoft.com/en-us/library/mt592918.aspx)
- [使用している Office のバージョンを確認する方法](https://support.office.com/en-us/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19?ui=en-US&rs=en-US&ad=US&fromAR=1)
- [Office 365 クライアント アプリケーションのバージョン番号およびビルド番号を確認することができます。](https://technet.microsoft.com/en-us/library/mt592918.aspx#Anchor_1)
- [Office Online Server 概要](https://technet.microsoft.com/en-us/library/jj219437(v=office.16).aspx)

## <a name="office-common-api-requirement-sets"></a>Office 共通 API の要件セット
共通 API の要件セットについて詳しくは、「[Office 共通 API の要件セット](office-add-in-requirement-sets.md)」をご覧ください。

## <a name="dialog-api-11"></a>ダイアログ API 1.1 
ダイアログ API 1.1 は、API の最初のバージョンです。API について詳しくは、[ダイアログ API](../shared/officeui.md) リファレンスのトピックをご覧ください。

## <a name="additional-resources"></a>追加リソース

- [Office のホストと API の要件を指定する](../../docs/overview/specify-office-hosts-and-api-requirements.md)
- [Office アドインの XML マニフェスト](../../docs/overview/add-in-manifests.md)

