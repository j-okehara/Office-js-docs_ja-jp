# Word アドインの JavaScript リファレンス 

 Word アドイン用の Word JavaScript API の API リファレンスを紹介しています。

_適用対象:Word 2016、Word for iPad、Word for Mac_

## このセクションの内容

以下は、Word JavaScript API のメイン オブジェクトです。

* [Body](word-add-ins-javascript-reference/body.md):文書またはセクションの本文を表します。
* [ContentControl](word-add-ins-javascript-reference/contentcontrol.md):コンテンツのコンテナー。ラベル付けすることができる、バインドされたドキュメント内の領域で、特定の種類のコンテンツのコンテナーとして機能します。ContentControl の内容には、たとえば、書式設定されたテキストの段落や他のコンテンツ コントロールなどを格納できます。ドキュメント、ドキュメントの本文、段落、範囲、またはコンテンツ コントロールのコンテンツ コントロールのコレクションを介してコンテンツ コントロールにアクセスできます。
* [Document](word-add-ins-javascript-reference/document.md):最上位のオブジェクト。ドキュメント オブジェクトには、1 つ以上の[セクション](word-add-ins-javascript-reference/section.md)と、ドキュメントの内容を含む本文、ヘッダー/フッター情報が含まれます。
* [Font](word-add-ins-javascript-reference/font.md):本文、コンテンツ コントロール、段落、または範囲にテキストの書式設定を指定します。
* [Image](word-add-ins-javascript-reference/inlinepicture.md):段落に固定されているインライン画像を表します。
* [Paragraph](word-add-ins-javascript-reference/paragraph.md):選択部分、範囲、またはドキュメント内の 1 つの段落を表します。段落へは、選択部分、範囲、またはドキュメント内の段落コレクションを介してアクセスできます。 
* [Range](word-add-ins-javascript-reference/range.md):文書内の連続した領域を表します。選択部分を取得するとき、コンテンツを本文に挿入するとき、コンテンツをコンテンツ コントロールに挿入するとき、コンテンツを段落に挿入するとき、または検索結果を取得するときには、範囲オブジェクトを取得します。選択部分を変更しないで、範囲を定義して操作できます。
* [Section](word-add-ins-javascript-reference/section.md):さまざまなヘッダーとフッター、およびドキュメントの多様なページ レイアウト構成を定義します。セクションにはドキュメント オブジェクトからアクセスできます。 
* [Selection](word-add-ins-javascript-reference/document.md#getselection):ドキュメント オブジェクトでは、ドキュメント内のユーザーが選択した部分にアクセスできます。また何も選択されていない場合は、現在の挿入ポイントにアクセスできます。

## フィードバックをお寄せください

お客様からのフィードバックを重視しています。 

* ドキュメントを確認していだだき、ドキュメントに関する質問や問題があれば、直接このリポジトリに[問題を送信](https://github.com/OfficeDev/office-js-docs/issues)してお知らせください。
* プログラミングの経験と、今後のバージョン、コード サンプルなどで希望されるものについてお知らせください。ご提案やアイデアの入力には、[このサイト](http://officespdev.uservoice.com/)をご使用ください。

## その他の技術情報

* [Word アドイン](word-add-ins.md)
* [Word アドインのプログラミング ガイド](word-add-ins-programming-guide.md)
* [Office アドイン](https://msdn.microsoft.com/en-us/library/office/jj220060.aspx)
* [Office アドインを使う](http://dev.office.com/getting-started/addins)
* &lt;a herf="https://github.com/OfficeDev?utf8=%E2%9C%93&amp;query=Word"&gt;GitHub の Word アドイン&lt;/a&gt;
* [Word のスニペット エクスプローラー](http://officesnippetexplorer.azurewebsites.net/#/snippets/word)
