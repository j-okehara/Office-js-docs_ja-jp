# Office アドインの UX 設計パターン。 

Office アドインを設計する場合、ユーザーのアドインの UX 設計は、Office を拡張する魅力的なエクスペリエンスを提供する必要があります。優れたアドインを作成するために、アドインは初回実行時エクスペリエンス、ファーストクラスの UX エクスペリエンス、およびページ間のスムーズな移動などを提供する必要があります。クリーンでモダンな UX エクスペリエンスを提供することにより、ユーザーによるアドインの保持や採用が増加します。この記事では、設計者と開発者に UX のリソースを提供します。

* ベスト プラクティスに基づく共通の UX 設計パターンについて説明します。
* Office のファブリック コンポーネントとスタイルを実装します。
* 既定の Office UI の通常の拡張子のようなアドインを実装します。 

## Office アドインの設計サンプルのリソースは、どのような方法で使用を開始できますか。

これらの設計やコード資産を使用するための前提条件はありません。アドイン用に優れた UX の作成を開始するには:

* UX 設計パターンを確認して、アドインに重要なパターンを識別します。たとえば、最初の実行エクスペリエンスの 1 つを選択します。
* 次に、今のいずれかの操作を実行します。
	* コード ファイルをアドイン プロジェクトにコピーして、要件を満たすようにカスタマイズします。必要な設計パターンの [common.js ファイル](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/)、[資産フォルダー](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/assets)、およびコード フォルダーが必要です。次のリンクを参照してください。
	* 参照 PDF をダウンロードして、ユーザー独自の UX 設計を作成する際のガイドとして使用します。次のリンクを参照してください。
	* Adobe Illustrator ファイルをダウンロードして、ユーザー独自のアドイン設計を模擬表示するように編集します。ファイルは[ここから](https://github.com/OfficeDev/Office-Add-in-Design-Patterns/blob/master/Patterns/Source%20Files)ダウンロードできます。
 

## 初回実行時

初回実行時エクスペリエンスは、最初にアドインを開いたときにユーザーが持つエクスペリエンスです。次の一覧は、アドインに含めることができる初回実行時の設計パターンです。その下に、設計パターンの各画像を示します。

* **開始手順**: アドインの使用を開始する手順の、順序付きリストをユーザーに提供します。([PDF](https://github.com/OfficeDev/Office-Add-in-Design-Patterns/blob/master/Patterns/FirstRun_StepsToStart.pdf "PDF")、[コード](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/templates/first-run/instruction-step))
* **価値**: アドインの価値提供を明確にします。([PDF](https://github.com/OfficeDev/Office-Add-in-Design-Patterns/blob/master/Patterns/FirstRun_ValuePlacemat.pdf "PDF")、[コード](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/templates/first-run/value-placemat))
* **ビデオ**: アドインの使用を開始する前に、ユーザーにビデオを表示します。([PDF](https://github.com/OfficeDev/Office-Add-in-Design-Patterns/blob/master/Patterns/FirstRun_VideoPlacemat.pdf "PDF")、[コード](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/templates/first-run/video-placemat))
* **チュートリアル**: アドインの使用を開始する前に、ユーザーに一連の機能または情報を体験させます。([PDF](https://github.com/OfficeDev/Office-Add-in-Design-Patterns/blob/master/Patterns/FirstRun_PagingPanel.pdf "PDF")、[コード](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/templates/first-run/walkthrough))
* [Office ストア](https://msdn.microsoft.com/ja-jp/library/office/jj220033.aspx)にはユーザーにアドインの試用版を提供するシステムがありますが、試用エクスペリエンスに UI のフル コントロールが必要な場合は、次のテンプレートを使用してください。
	* **試用**: アドインの試用版で開始する方法をユーザーに示します。([PDF](https://github.com/OfficeDev/Office-Add-in-Design-Patterns/blob/master/Patterns/FirstRun_TrialVersion.pdf "PDF")、[コード](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/templates/first-run/trial-placemat))
	* **試用版の機能**: ユーザーが試用を考えている機能がアドインの試用版では使用できないことをユーザーに示します。([コード](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/templates/first-run/trial-placemat-feature))


> メモ:ユーザーに初回実行時エクスペリエンスを 1 回示すか、何回も示すかを検討することがシナリオにとって重要かどうかを検討します。たとえば、ユーザーがアドインを定期的にしか使用しない場合は、ユーザーがアドインの使用方法を忘れる可能性があります。これらのユーザーには、初回実行時エクスペリエンスを再度表示すると役に立つ可能性があります。 

 <table>
 <tr><th>開始手順</th><th>価値</th><th>ビデオ</th></tr>
 <tr><td>![instruction steps" style="width: 264px;](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/blob/master/Images/instruction.step.PNG)</td><td>![value placemat" style="width: 264px;](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/blob/master/Images/value.placemat.PNG)</td><td>![video placemat" style="width: 264px;](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/blob/master/Images/video.placemat.PNG)</td></tr>
 </table>

 <table>
 <tr><th>チュートリアルの最初のページ</th><th>試用</th><th>試用版の機能</th></tr>
 <tr><td>![walkthrough 1" style="width: 264px;](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/blob/master/Images/walkthrough1.PNG)</td><td>![trial placemat" style="width: 264px;](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/blob/master/Images/trial.placemat.PNG)</td><td>![trial placemat feature" style="width: 264px;](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/blob/master/Images/trial.placemat.feature.PNG)</td></tr>
 </table> 


## ジェネリックとブランド化

* **ランディング ページ**は、初回実行時エクスペリエンスの後またはサインイン プロセスの後でユーザーが移動する最初の場所です。([PDF](https://github.com/OfficeDev/Office-Add-in-Design-Patterns/blob/master/Helpful%20Templates/AddIn_Template_Standard_Layout.pdf "PDF")、[コード](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/templates/generic/landing-page))

<table>
 <tr><th>ランディング</th></tr>
 <tr><td>![landing page" style="width: 264px;](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/blob/master/Images/landing.page.PNG)</td></tr>
 </table>

## 通知

アドインがユーザーにエラーなどのイベントや進展を通知する方法は、いろいろあります。これらの方法を次に示します。その下に、設計パターンの各画像を示します。

* **埋め込みダイアログ**: ボタンまたはその他のコントロールを使用して、情報や必要に応じて対話型エクスペリエンスを提供するダイアログを作業ウィンドウ内に表示します。いずれか 1 つを使用して、ユーザーにアクションの確認を促すことを検討します。([PDF](https://github.com/OfficeDev/Office-Add-in-Design-Patterns/blob/master/Patterns/Embedded_Dialog.pdf "PDF")、[コード](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/templates/notifications/embedded-dialog))
* **インライン メッセージ**: エラー、成功、情報を示します。メッセージは作業ウィンドウ内の指定した場所に表示できます。たとえば、ユーザーがテキスト ボックスに不適切な書式の電子メール アドレスを入力すると、テキスト ボックスの真下にエラー メッセージが表示されます。([PDF](https://github.com/OfficeDev/Office-Add-in-Design-Patterns/blob/master/Patterns/Notification_Inline_Message.pdf "PDF")、[コード](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/templates/notifications/inline-message))
* **メッセージ バナー**: 情報や必要に応じてアクションの単純な呼び出しを、単一の行に折りたたんだり、複数の行に展開したり、非表示にできるバナーで提供します。サービスの更新またはアドインを開始するときに役に立つヒントの報告にメッセージ バナーを使用することを検討します。([PDF](https://github.com/OfficeDev/Office-Add-in-Design-Patterns/blob/master/Patterns/Notification_messagebanner.pdf "PDF")、[コード](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/templates/notifications/message-banner))
* **進行状況バー**: ユーザーが以降のアクションを行う前に完了する必要がある、実行時間の長い、同期プロセス (構成タスクなど) の進行状況を示します。これは別のスポット ページであり、アドインのブランド化を強化します。進行状況バーは、アドインに戻るまでの定期的な評価をプロセスが送信する際に使用します。([PDF](https://github.com/OfficeDev/Office-Add-in-Design-Patterns/blob/master/Patterns/Notification_progress.pdf "PDF")、[コード](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/templates/notifications/progress-bar))
* **スピナー**: 実行時間の長い、同期プロセスが進行中だが、どの程度進んでいるのかは提示されていないことを示します。これは別のスポット ページであり、アドインのブランド化を強化します。スピナーは、プロセスがどの程度進んでいるかをアドインが確実に知ることができない場合に使用します。([PDF](https://github.com/OfficeDev/Office-Add-in-Design-Patterns/blob/master/Patterns/Notification_progress.pdf "PDF")、[コード](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/templates/notifications/spinner))
* **トースト**: 数秒で消える簡単なメッセージを提供します。ユーザーがメッセージに気づかない場合があるため、トーストは重要ではない情報にのみ使用します。これは、電子メールの受信など、リモート システムでユーザーにイベントを通知する場合に優れた選択になります。([PDF](https://github.com/OfficeDev/Office-Add-in-Design-Patterns/blob/master/Patterns/Notification_toast.pdf "PDF")、[コード](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/templates/notifications/toast))

 <table>
 <tr><th>埋め込みダイアログ</th><th>インライン メッセージ</th><th>メッセージ バナー</th></tr>
 <tr><td>![embedded dialog" style="width: 264px;](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/blob/master/Images/embedded.dialog.PNG)</td><td>![inline message" style="width: 264px;](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/blob/master/Images/inline.message.PNG)</td><td>![message banner" style="width: 264px;](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/blob/master/Images/message.banner.PNG)</td></tr>
 </table>

 <table>
 <tr><th>進行状況バー</th><th>スピナー</th><th>トースト</th></tr>
 <tr><td>![progress bar" style="width: 264px;](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/blob/master/Images/progress.bar.PNG)</td><td>![spinner" style="width: 264px;](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/blob/master/Images/spinner.PNG)</td><td>![toast" style="width: 264px;](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/blob/master/Images/toast.PNG)</td></tr>
 </table>

## 既知の問題

* アドイン プロジェクトの外部でコード ファイルを実行すると、JavaScript エラーがスローされます。 
	* 解決方法:これらのファイルが Office アドイン プロジェクトに追加されていることを確認します。 
	
## その他の技術情報

* [Office アドインの設計のベスト プラクティス](https://dev.office.com/docs/add-ins/design/add-in-development-best-practices)
* [Office UI Fabric](http://dev.office.com/fabric/)
