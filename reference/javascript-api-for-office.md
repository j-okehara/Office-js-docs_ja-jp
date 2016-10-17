
# <a name="javascript-api-for-office-reference"></a>JavaScript API for Office リファレンス

JavaScript API for Office を使用すると、Office ホスト アプリケーションのオブジェクト モデルと対話する Web アプリケーションを作成できます。ユーザーのアプリケーションは、スクリプト ローダーである office.js ライブラリを参照します。Office.js ライブラリは、アドインを実行している Office アプリケーションに適用可能なオブジェクト モデルを読み込みます。次の JavaScript オブジェクト モデルを使用できます。


1. 共通 (必須) - Office 2013 で導入された API。これは、**すべての Office ホスト アプリケーション**に読み込まれ、アドイン アプリケーションを Office クライアント アプリケーションに接続します。オブジェクト モデルには、Office クライアントに固有の API と複数の Office クライアントのホスト アプリケーションに適用可能な API が含まれています。[[共有]](../reference/shared/shared-api.md) と **[outlook]** の下のすべてのコンテンツは、共通 API と見なされます。**Microsoft.Office.WebExtension** 名前空間 (コード内では既定で [Office](../reference/shared/office.md) というエイリアスを使用して参照される) には、Office アドインから Office ドキュメント、ワークシート、プレゼンテーション、メール アイテム、プロジェクトのコンテンツを操作できるスクリプトの記述に利用できるオブジェクトが含まれています。アドインが Office 2013 以降を対象にする場合、これらの共通 API を使用する必要があります。このオブジェクト モデルは、callback を使用します。

1. ホスト固有 - **Office 2016** で導入された API。このオブジェクト モデルは、Office クライアントの使用時に見られる使い慣れたオブジェクトに対応するホスト固有の厳密に型指定されたオブジェクトを提供し、Office JavaScript API の将来像を表すものです。現在、ホスト固有の API には、[Word JavaScript API](../reference/word/word-add-ins-reference-overview.md) と [Excel JavaScript API](../reference/excel/application.md) が含まれています。このオブジェクト モデルは、promise を使用します。

Office クライアントを TOC の上のドロップダウン リストから選択して、対象のホスト アプリケーションに基づいてコンテンツをフィルター処理します。

## <a name="supported-host-applications"></a>サポートされるホスト アプリケーション
* Access
* Excel
* Outlook
* PowerPoint
* Project
* Word

[サポートされるホストとその他の要件](../docs/overview/requirements-for-running-office-add-ins.md)の詳細について説明します。

## <a name="open-api-specifications"></a>Open API の仕様

新しい Office アドイン用の API の設計と開発にあたり、[Open API の仕様](openspec.md) ページでこれらに対するフィードバックの提供が可能になります。パイプラインの新機能をご確認いただき、設計の仕様に関する情報をお寄せください。

