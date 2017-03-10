-
#<a name="use-office-ui-fabric-in-office-add-ins"></a>Office アドインでの Office UI Fabric の使用

Office アドインを作成する場合は、[Office UI Fabric](https://dev.office.com/fabric) を使用して、ユーザー エクスペリエンスを作成することをお勧めします。 

Office UI Fabric は、Office と Office 365 のユーザー エクスペリエンスを構築するための JavaScript フロント エンドのフレームワークです。Fabric は、拡張や改訂が可能な視覚効果に焦点を合わせたコンポーネントであり、Office アドインで使用できます。Fabric は Office デザイン言語を使用するため、Fabric の UX コンポーネントは Office 本来の拡張機能のように表示されます。

Fabric は、次に示す複数のコンポーネントから構成されています。

- **Fabric JS (推奨)**: JavaScript のみを使用して UX コンポーネントを実装します。このバージョンの Fabric は、React フレームワークへの依存を望まない場合にお勧めします。  
- **Fabric React**: React フレームワークを使用して UX コンポーネントを実装します。
- **Fabric Core**: デザイン言語の主要要素 (アイコン、色、タイプ、グリッドなど) が含まれます。Fabric JS と Fabric React は、どちらも Fabric Core を使用します。 

次に示す手順では、Fabric JS の基本的な使用方法について説明します。  

##<a name="1-add-the-fabric-cdn-references"></a>1. Fabric CDN 参照の追加
CDN から Fabric を参照するには、次に示す HTML コードをページに追加します。

    <link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.4.0/css/fabric.min.css">
    <link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.4.0/css/fabric.components.min.css">
    <script src="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.4.0/js/fabric.min.js"></script>

これで完了です。この時点で、Fabric をアドインで使用する準備が整っています。 

##<a name="2-use-fabric-icons-and-fonts"></a>2.Fabric のアイコンとフォントの使用
アイコンは簡単に使用できます。"i" 要素を使用して、適切なクラスを参照するだけです。アイコンのサイズは、フォント サイズを変更することで制御できます。たとえば、次のコードは、themePrimary (#0078d7) 色を使用する特大の表アイコンを作成する方法を示しています。 
   
    <i class="ms-Icon ms-font-xl ms-Icon--Table ms-fontColor-themePrimary"></i>

その他の Office UI Fabric で使用可能なアイコンを見つけるには、「[アイコン](https://dev.office.com/fabric#/styles/icons)」ページの検索機能を使用します。アドインで使用するアイコンを検索するときには、アイコン名の先頭に `ms-Icon--` を追加していることを確認してください。 

Office UI Fabric で使用可能なフォントのサイズと色についての詳細は、「[文字体裁](https://dev.office.com/fabric#/styles/typography)」および「[色](https://dev.office.com/fabric#/styles/colors)」を参照してください。

##<a name="3-use-fabric-js-ux-components"></a>3. Fabric JS UX コンポーネントの使用

Fabric は、アドインで使用できるボタンやチェックボックスなど、複数の UX コンポーネントを提供しています。次に、アドインでの使用をお勧めする Fabric JS UX コンポーネントのリストを示します。アドインで Fabric コンポーネントのいずれかを使用するには、その Fabric のドキュメントへのリンクをたどって、「**このコンポーネントの使用方法**」の手順を実行してください。

> **メモ:** 追加のコンポーネントを徐々に増やしていく予定です。 

- [Breadcrumb](https://dev.office.com/fabric-js/Components/Breadcrumb/Breadcrumb.html)
- [Button](https://dev.office.com/fabric-js/Components/Button/Button.html) (アドインで小さなボタンのバリエーションの使用を検討してください。)
- [Checkbox](https://dev.office.com/fabric-js/Components/CheckBox/CheckBox.html)
- [ChoiceFieldGroup](https://dev.office.com/fabric-js/Components/ChoiceFieldGroup/ChoiceFieldGroup.html)
- [Date Picker](https://dev.office.com/fabric-js/Components/DatePicker/DatePicker.html) (アドインに日付ピッカーを実装する方法の例は、[Excel Sales Tracker](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker) コード サンプルを参照してください)
- [Dropdown](https://dev.office.com/fabric-js/Components/Dropdown/Dropdown.html)
- [Label](https://dev.office.com/fabric-js/Components/Label/Label.html)
- [Link](https://dev.office.com/fabric-js/Components/Link/Link.html)
- [List](https://dev.office.com/fabric-js/Components/List/List.html) (コンポーネントの既定のスタイルを CSS で変更することを検討してください)
- [MessageBanner](https://dev.office.com/fabric-js/Components/MessageBanner/MessageBanner.html)
- [MessageBar](https://dev.office.com/fabric-js/Components/MessageBar/MessageBar.html)
- [Overlay](https://dev.office.com/fabric-js/Components/Overlay/Overlay.html)
- [Panel](https://dev.office.com/fabric-js/Components/Panel/Panel.html)
- [Pivot](https://dev.office.com/fabric-js/Components/Pivot/Pivot.html)
- [ProgressIndicator](https://dev.office.com/fabric-js/Components/ProgressIndicator/ProgressIndicator.html)
- [Searchbox](https://dev.office.com/fabric-js/Components/SearchBox/SearchBox.html)
- [Spinner](https://dev.office.com/fabric-js/Components/Spinner/Spinner.html)
- [Table](https://dev.office.com/fabric-js/Components/Table/Table.html)
- [TextField](https://dev.office.com/fabric-js/Components/TextField/TextField.html)
- [Toggle](https://dev.office.com/fabric-js/Components/Toggle/Toggle.html)
   
## <a name="updating-your-add-in-to-use-fabric-js"></a>Fabric JS を使用するためのアドインの更新
以前のバージョンの Office UI Fabric を使用しているときに、Fabric JS への移行を考えている場合は、新しいコンポーネントをアドインに組み込んでテストする方法について理解していることを確認します。次に示す点に注意して、更新の計画に役立ててください。

- Fabric JS を使用することでコンポーネントの初期化が簡単になります。以前のバージョンの Fabric の場合は、Fabric コンポーネントの JavaScript ファイルをアドイン プロジェクト (そのファイルへの `<Script>` 参照が含まれているプロジェクト) に含めてからコンポーネントを初期化します。Fabric JS では、Fabric コンポーネントの JavaScript ファイルと、それに関連する `<Script>` 参照を含める必要はなくなりました。Fabric コンポーネントの初期化以外に必要な手順はありません。   
- いくつかのコンポーネントは、UX コンポーネントの動作を制御する関数を提供するようになりました。たとえば、チェックボックス コントロールには、チェックボックスのオン状態とオフ状態を切り替える `toggle` 関数があります。 
- 一部のアイコン クラスの名前とスタイルが更新されています。
- 最重要の変更点は、多数のコンポーネントで `<label>` 要素を使用していることです。`<label>` 要素では、コンポーネントのスタイルを制御します。`<label>` 要素を使用するように UX コードを更新することが必要になる場合があります。たとえば、Fabric JS のチェックボックスに対する `<input>` 要素のオン属性の値を変更しても、そのチェックボックスに効果は現れません。その代わりに、`check`、`unCheck`、または `toggle` 関数を使用してください。   

##<a name="next-steps"></a>次の手順
Fabric JS の使用方法を示す完全なサンプルを探しているユーザーに向けて、そのようなサンプルを提供しています。次に示すリソースを参照してください。

- [Excel Sales Tracker](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker) 

##<a name="related-resources"></a>関連リソース
以前のリリースの Fabric に関するコード サンプルやドキュメントを探している場合は、次に示す記事を参照してください。

- [UX 設計パターン (Fabric 2.6.1 を使用)](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code) 
- [Office アドイン Fabric UI サンプル (Fabric 1.0 を使用)](https://github.com/OfficeDev/Office-Add-in-Fabric-UI-Sample) 
- [Office アドインでの Fabric 2.6.1 の使用](https://dev.office.com/docs/add-ins/design/ui-elements/using-office-ui-fabric)
 

