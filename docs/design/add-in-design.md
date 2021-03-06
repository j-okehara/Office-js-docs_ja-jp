# <a name="design-guidelines-for-office-add-ins"></a>Office アドインの設計ガイドライン

Office アドインは、ユーザーが Office クライアントで使用できるコンテキストに応じた機能を提供することで、Office のエクスペリエンスを拡張します。アドインにより、コストのかかるコンテキストの切り替えなしで、サード パーティの機能が Office で使用できるようになり、ユーザーの生産性は向上します。 

 アドインの UX 設計は、効率の高い自然な対話操作をユーザーに提供するために、Office とシームレスに統合する必要があります。アドインのコマンド (Office UI 拡張機能) を利用して、アドインへのアクセスを提供します。また、HTML ベースのカスタム UI を作成するときに推奨される [UI 要素](ui-elements/ui-elements.md)と[ベスト プラクティス](https://dev.office.com/docs/add-ins/overview/add-in-development-best-practices)を使用します。 
 
 
## <a name="core-office-add-in-design-principles"></a>Office アドイン設計の主な原則
どの基盤フレームワークを使用してカスタム UI を作成するにしても、アドインの設計には次の原則が適用されます。 

- **Office に合わせて、わかりやすく設計する**。アドインの機能と外観は、Office またはドキュメントのテーマを適用することも含め、Office のエクスペリエンスと調和したものであるべきです。
 
- **ユーザーの効率が向上されるようにする**。あるジョブの実行が、別のジョブの邪魔にならないようにユーザーを支援します。Office ドキュメントとアドインの間でシームレスに対話操作ができるようにします。 

- **クロムよりもコンテンツを優先する**。どのようなアクセサリ クロムよりも、アドインのコンテンツと機能を重視します。ユーザー エクスペリエンスの価値を上げない不要な UI 要素を排除して、領域の使用効率を最大化します。  

- **ユーザーによる制御を可能にする**。ユーザーが個人のエクスペリエンスを制御できるようにして、重要な決定を理解できるようにします。また、アドインで実行したアクションは、簡単に元に戻せるようにします。 

- 
  **すべてのプラットフォームおよび入力方式に対応するように設計する**。アドインは、Office がサポートしている、すべてのプラットフォームで動作するように設計します。また、アドインの UX は、あらゆるプラットフォームおよびフォーム ファクターで最適に機能する必要があります。マウス/キーボードとタッチ入力のデバイスをサポートして、カスタムの HTML UI が各種のフォーム ファクターに順応するようにします。詳細は、「[タッチ](https://msdn.microsoft.com/EN-US/library/mt590883.aspx#bk_Touch)」をご参照ください。 


## <a name="design-language"></a>デザイン言語
アドインに HTML ベースのカスタム エクスペリエンスを作成する場合は、Office デザイン言語と [Office UI Fabric](https://dev.office.com/fabric) の使用をお勧めします。すでに組織にデザイン言語が用意されている場合、最終的な結果が、Office ユーザーにとって親しみやすいエクスペリエンスになるならば、そのツールを使用してください。 


## <a name="add-in-building-blocks"></a>アドインの構築ブロック
アドインを作成するときには、次に示す 2 種類の UI 要素を使用できます。 

- [アドイン コマンド](ui-elements/ui-elements.md#add-in-commands): 使用すると、Office アプリケーションに、ネイティブな UX フックを追加できます。
- [ HTML ベースのカスタム UI](ui-elements/ui-elements.md#custom-html-based-ui): 使用すると、Office クライアントで HTML の機能を利用できるようになります。 

これらの構築ブロックを使用する方法の詳細は、「[UI 要素](ui-elements/ui-elements.md)」を参照してください。  

## <a name="ux-design-patterns"></a>UX 設計パターン

アドインのファースト クラスのユーザー エクスペリエンスを作成するために、共通の UX 設計パターンを設計するためのテンプレートが提供されています。これらのテンプレートには、説得力のある世界レベルのアドインを作成するための[ベスト プラクティス](https://dev.office.com/docs/add-ins/overview/add-in-development-best-practices)が反映されています。また、最初の実行エクスペリエンスのパターンや要素のブランド化、ユーザー通知が含まれています。テンプレートでは [Office UI Fabric](https://dev.office.com/fabric) のコンポーネントとスタイルが使用されています。また、自然に Office UI を拡張する要素が含まれています。

テンプレートにアクセスするには、「[Office Add-in UX Design Patterns](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns)」レポートを参照してください。Adobe Illustrator のファイルも使用できます。ファイルをダウンロードおよび更新して、自分の設計を反映させることができます。「[Office Add-in UX design patterns code](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)」レポートからコード ファイルを自分のアドイン プロジェクトにコピーしたり、必要に応じてカスタマイズすることもできます。 

## <a name="recommended-layouts-and-interaction-patterns"></a>推奨されるレイアウトと対話式操作のパターン
アドインの種類ごとに、推奨されるレイアウトが用意されています。これには、すべてのまとめに役立つ**エンドツーエンド**のサンプルが付属しています。アドインのレイアウト方法についての詳細は、次を参照してください。

- [作業ウィンドウ コンテナーのレイアウト](ui-elements/layout-for-task-pane-add-ins.md)
- [コンテンツ アドインのレイアウト](ui-elements/layout-for-content-add-ins.md) 
- [メール アドイン用のレイアウト](ui-elements/layouts-for-outlook-add-ins.md)

アドインと、そのアドインに対応する対話式操作のパターンに関する一般的なシナリオについては、「対話式操作のパターン」も参照してください。

## <a name="additional-resources"></a>その他のリソース

- [Office の UI ファブリック](https://dev.office.com/fabric) 

