# <a name="open-api-specifications"></a>Open API の仕様

設計中の新しい API と機能に関する詳細情報にご関心をお寄せいただきありがとうございます。コミュニティのフィードバックをいただくため、初期のバージョンの API 仕様をここでご確認いただけます。お客様からの情報によって、重要なユース ケースにぴったりの最終的な設計を確立することができます。 

ここで説明する機能は、初期設計、パブリック プレビューなど、さまざまな開発段階にあるものです。機能が一般公開されると、このページからコンテンツが削除され、新機能の詳細が含まれるようにドキュメントを更新します。 

_**重要:**以下の機能は、まだ設計とレビュー中の段階であり、まだ一般公開されていません。これらの機能と API は変更される場合があります。_

## <a name="visio-javascript-apis"></a>Visio JavaScript API
Visio Online は、Visio 図面を Web 上で表示し共有する新しい方法です。Visio JavaScript API 1.1 を使用して Visio Online の機能を拡張できます。SharePoint ページに埋め込まれた Visio 図面に対して、これらの API を使用します。現在、Visio JavaScript API は [Office アドイン](https://dev.office.com/docs/add-ins/overview/office-add-ins)には適用されませんのでご注意ください。

**詳細については、[Visio JavaScript API 1.1](https://github.com/OfficeDev/office-js-docs/tree/VisioJs_1.1_Openspec) のページを参照し、フィードバックを提供してください。**

## <a name="new-excel-javascript-apis"></a>新しい Excel JavaScript API
新しい Excel JavaScript API の設計のレビューにご参加ください。新しい API および更新された API には、カスタム XML パーツ、ピボット テーブルの更新、フィルターされたビューの範囲、イメージの範囲とテーブル、テーブルへの複数行の追加、その他の機能が含まれています。 

**詳細については [Excel JavaScript 1.3 API のページ](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec)を参照し、フィードバックを提供してください。**

## <a name="new-word-javascript-apis"></a>新しい Word JavaScript API
Word JavaScript API 1.3 更新プログラムには、この API の導入以降実装された最大の変更セットが含まれています。新しい API を使用すると、以下を行えます。 

* メモリ内のドキュメントを作成し、変更します
* リスト オブジェクトを作成し、アクセスします
* テーブル オブジェクトを作成し、アクセスします
* 範囲オブジェクトにアクセスし比較するためのオプションがさらに用意されています

Word JavaScript API のほぼすべてのオブジェクト間で、これらの変更が実装されました。この機能は、現在、またはまもなく Windows と Mac 両方のデスクトップと iPad 上の Word 2016 のプレビュー版で利用できます。クライアントを最新のビルドに毎月更新して、これら魅力的な機能の実装を開始してください。

**詳細については、[Word JS 1.3 API のページ](https://github.com/OfficeDev/office-js-docs/tree/WordJs_1.3_Openspec/word)を参照し、フィードバックを提供してください。**

## <a name="document-properties-access"></a>ドキュメント プロパティへのアクセス
ドキュメント レベルのプロパティにアクセス (get、set) する Web アドインの機能を追加する作業中です。アドインは、この機能によって、ドキュメント プロパティをカスタム ワークフローの一部として統合でき、またドキュメント プロパティの読み取り/設定ができるようなります。Word、Excel、およびおそらく PowerPoint は、この機能をサポートします。この機能は Excel REST API でも使用できます (Excel では、REST サービスをサポートします)。API が追加される場合の機能の仕方に関するユース ケースとコード スニペットを通して、設計の基本的な考え方と機能についてご紹介します。設計に関するフィードバックをお寄せください。 

**詳細については、[ドキュメント プロパティのオープン仕様のページ](https://github.com/OfficeDev/office-js-docs/tree/DocumentProperties_OpenSpec)を参照し、フィードバックを提供してください。**

