
# <a name="referencing-the-javascript-api-for-office-library-from-its-content-delivery-network-cdn"></a>Office ライブラリの JavaScript API をそのコンテンツ配信ネットワーク (CDN) から参照する


[JavaScript API for Office](../../reference/javascript-api-for-office.md) ライブラリは、Office.js ファイルと関連するホスト アプリケーション固有のファイル (Excel-15.js や Outlook-15.js など) で構成されています。 


最も簡単に API を参照する方法は、次に示す `<script>` をページの `<head>` タグに追加して、CDN を使用することです。  

```html
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

CDN URL で `office.js` の前にある `/1/` は、Office.js のバージョン 1 内で最新の増分リリースを使用するよう指定します。JavaScript API for Office が旧バージョンとの互換性を維持するため、最新リリースはバージョン 1 で導入されていた API メンバーを引き続きサポートします。既存のプロジェクトを更新する必要がある場合は、「[JavaScript API for Office およびマニフェスト スキーマ ファイルのバージョンを更新する](../docs/develop/update-your-javascript-api-for-office-and-manifest-schema-version.md)」を参照してください。 

Office ストアから Office アドインを発行する場合は、この CDN の参照を使用する必要があります。ローカル参照は、内部の開発およびデバッグ シナリオにのみ適用できます。

> **重要:**任意の Office ホスト アプリケーションのアドインを開発する場合、ページの `<head>` セクション内から JavaScript API for Office を参照することが重要です。これにより、API はあらゆる body 要素の前に完全に初期化されます。Office ホストでは、ライセンス認証の 5 秒以内にアドインを初期化する必要があります。このしきい値を超えるとアドインが応答なしと宣言され、エラー メッセージがユーザーに表示されます。       

## <a name="additional-resources"></a>追加リソース



- [JavaScript API for Office について](../../docs/develop/understanding-the-javascript-api-for-office.md)
    
- [Office アドイン プラットフォームの概要](../../docs/overview/office-add-ins.md)
    
- [Office アドインの開発ライフ サイクル](../../docs/design/add-in-development-lifecycle.md)
    
- [JavaScript API for Office](../../reference/javascript-api-for-office.md)
    
