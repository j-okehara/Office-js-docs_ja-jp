
# Office アドインのリソースの制限とパフォーマンスの最適化



ユーザーのベスト エクスペリエンスを実現するために、Office アドイン実行時の CPU コアとメモリの使用量、および信頼性を一定の範囲内に保つ必要があります。Outlook アドインでは、これに加えて正規表現の評価の応答時間を一定以内に保つ必要があります。これらの実行時のリソース使用量の制限は、Windows と OS X 用の Office クライアントに適用され、Office Online、Outlook Web App、デバイス用 OWA には適用されません。また、デスクトップやモバイル デバイス上のアドインについても、アドインの設計と実装でリソース使用量を最適化することによって、そのパフォーマンスを最適化できます。

## アドインのリソース使用量の制限


実行時のリソース使用量の制限は、すへての種類の Office アドインに適用されます。このような制限は、ユーザーのパフォーマンスの向上およびサービス拒否攻撃の影響緩和にも役立ちます。想定される一連のデータを使用して対象のホスト アプリケーションで Office アドインをテストし、次に示す制限の範囲内でパフォーマンスを調整してください。


-  **CPU コアの使用率**: 単一の CPU コアの使用率しきい値 90%、既定の 5 秒間隔で 3 回観測。
    
    CPU コアの使用率を確認するホスト リッチ クライアントの既定の間隔は、5 秒間隔です。 ホスト クライアントでアドインの CPU コアの使用率がしきい値を超えたことを検知した場合、ユーザーがアドインの実行を継続するかどうかを確認するメッセージが表示されます。 ユーザーが継続することを選択した場合、編集セッション中にホスト クライアントがユーザーにもう一度確認することはありません。 ユーザーが CPU を集中的に使用するアドインを実行する場合、この警告メッセージの表示を減らすには、管理者は **AlertInterval** レジストリ キーを使用する必要がある可能性があります。
    
-  **メモリ使用量**: デバイスの利用可能な物理メモリに基づいて動的に決定される、既定のメモリ使用量しきい値。
    
    既定では、ホスト リッチ クライアントが、デバイスの物理メモリの使用率が利用可能なメモリの 80% を超えたことを検知した場合、クライアントはコンテンツ アドインおよびタスク ウィンドウ アドインのドキュメント レベル、および Outlook アドインのメールボックス レベルで、アドインのメモリ使用率の監視を開始します。 既定の 5 秒間隔で、ドキュメントまたはメールボックス レベルでアドインのセットの物理メモリの使用率が 50% を超えた場合、クライアントはそのユーザーに警告します。 このメモリ使用量の制限では、タブレットなど、限られた RAM が搭載されたデバイスのパフォーマンスを確保するために、仮想メモリよりも物理メモリを使用します。 管理者は、グローバル設定として **MemoryAlertThreshold** Windows レジストリ キーを使用して、この動的設定を明示的な制限で上書きできます。また、グローバル設定として **AlertInterval** キーを使用して、警告の間隔を調整することもできます。
    
-  **クラッシュ許容度**: 既定の制限は、1 つのアドインにつき 4 回。
    
    管理者は、**RestartManagerRetryLimit** レジストリ キーを使用して、クラッシュのしきい値を調整できます。
    
-  **アプリケーションのブロッキング**: アドインが応答しないままになる時間のしきい値は 5 秒間。
    
    これは、アドインとホスト アプリケーションのユーザー エクスペリエンスに影響します。 このような場合、ホスト アプリケーションは、自動的にドキュメントまたはメールボックス (該当する場合) のアクティブなアドインをすべて再起動し、ユーザーに応答しなくなったアドインに関する警告を行います。 アドインが時間のかかるタスクを実行していて定期的に処理を発生させないときに、このしきい値に到達する場合があります。 ブロッキングが発生しないようにする手法があります。 管理者は、このしきい値を上書きすることはできません。
    
     **Outlook アドイン**
    
    Outlook アドインが前述の CPU コア使用率、メモリ使用量、またはクラッシュ許容度のしきい値を超えると、そのアドインは Outlook で無効化されます。Exchange 管理センターにはそのアプリの無効状態が表示されます。
    
     >**注** Outlook Web App および デバイス用 OWA ではなく、Outlook リッチ クライアントによってのみ、リソース配分状況を監視する場合でも、リッチ クライアントが Outlook アドインを無効化すると、このアドインは Outlook Web App および デバイス用 OWA の使用でも無効化されます。

    CPU コア、メモリ、および信頼性ルールだけでなく、Outlook アドインは次のアクティブ化のルールを監視する必要があります。
    
      -  **正規表現の応答時間**: Outlook で Outlook アドインのマニフェスト内のすべての正規表現を評価する時間の既定のしきい値は 1,000 ミリ秒。このしきい値を超えると、Outlook は後で評価を再試行します。
    
        Windows レジストリでグループ ポリシーまたはアプリケーションに固有の設定を使用すると、管理者は **OutlookActivationAlertThreshold** 設定でこの既定のしきい値の 1,000 ミリ秒を調整することができます。 詳細については、「[Office の開発](http://msdn.microsoft.com/library/da14ec8c-5075-4035-a951-fc3c2b15c04b%28Office.15%29.aspx)」を参照してください。
    
  -  **正規表現の再評価**: Outlook でマニフェスト内の正規表現を再評価する既定の制限は 3 回。適用されるしきい値 (既定の 1,000 ミリ秒、または Windows レジストリに **OutlookActivationAlertThreshold** 設定が存在する場合はその設定で指定された値) を 3 回とも超えて評価に失敗すると、その Outlook アドインは Outlook で無効化されます。Exchange 管理センターには無効状態が表示され、そのアドインは Outlook リッチ クライアント、Outlook Web App、および デバイス用 OWA で使用できなくなります。
    
    Windows レジストリでグループ ポリシーまたはアプリケーションに固有の設定を使用すると、管理者は **OutlookActivationManagerRetryLimit** 設定の評価を再試行する時間の数字を調整することができます。 詳細については、「[Office の開発](http://msdn.microsoft.com/library/da14ec8c-5075-4035-a951-fc3c2b15c04b%28Office.15%29.aspx)」を参照してください。
    

    **作業ウィンドウ アドインとコンテンツ アドイン**
    
    コンテンツ アドインまたは作業ウィンドウ アドインが前述の CPU コア使用率、メモリ使用量、またはクラッシュ許容度のしきい値を超えると、対応するホスト アプリケーションにユーザーへの警告が表示されます。この時点で、ユーザーは次のどちらかの処理を実行できます。
    
  - アドインを再起動します。
    
  - しきい値を超えたというそれ以降の警告をキャンセルします。理想的な対処としては、ユーザーはそのアドインをドキュメントから削除する必要があります。そのアドインの使用を続行すると、さらにパフォーマンスと安定性の問題が発生する可能性があります。
    

## テレメトリ ログでリソース使用量の問題を確認する


Office には、Office アドインでのリソースの使用に関する問題も含めて、ローカル コンピューター上で実行される Office ソリューションの一定のイベント (読み込む、開く、閉じる、およびエラー) の記録を保守するテレメトリ ログが用意されています。テレメトリ ログを設定してある場合は、Excel を使用して、ローカル ドライブ上の次の既定の場所にあるテレメトリ ログを開くことができます。

%Users%\ \<lt;現在のユーザー\>gt; \AppData\Local\Microsoft\Office\15.0\Telemetry

それぞれのアドインについてテレメトリ ログで追跡されるイベントごとに、そのイベントの発生日付/時刻、イベント ID、重大度、および短い説明的なタイトル、そのアドインのフレンドリ名と ID、イベントをログに記録したアプリケーションが記入されています。テレメトリ ログをリフレッシュすれば、現在の追跡済みイベントを確認できます。次の表は、テレメトリ ログで追跡された Outlook アドインの例を示しています。 



|**日付/時刻型 (Date/Time)**|**イベント ID**|**重大度**|**タイトル**|**ファイル**|**ID**|**アプリケーション**|
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
|10/8/2012 5:57:10 PM|7||アドインのマニフェストが正常にダウンロードされました|Who's Who|69cc567c-6737-4c49-88dd-123334943a22|Outlook|
|10/8/2012 5:57:01 PM|7||アドインのマニフェストが正常にダウンロードされました|LinkedIn|333bf46d-7dad-4f2b-8cf4-c19ddc78b723|Outlook|
 次の表は、通常、Office アドインについてテレメトリ ログで追跡されるイベントを示しています。



|**イベント ID**|**タイトル**|**重大度**|**説明**|
|:-----|:-----|:-----|:-----|
|7|アドインのマニフェストが正常にダウンロードされました||Office アドインのマニフェストがホスト アプリケーションによって正常に読み込まれ、読み取られました。|
|8|アドインのマニフェストがダウンロードされませんでした|重大|ホスト アプリケーションは Office アドインのマニフェスト ファイルを、SharePoint カタログ、コーポレート カタログ、Office ストアのいずれからも読み込めませんでした。|
|9|アドインのマークアップを解析できませんでした|重大|ホスト アプリケーションは Office アドインのマニフェストを読み込みましたが、アプリの HTML マークアップを読み取れませんでした。|
|10|アドインの CPU 使用率が高すぎます|重大|Office アドインは、限定された時間内に CPU リソースの 90% 超を使用しました。|
|15|アドインは文字列検索のタイムアウトのため無効になっています||Outlook アドインは電子メールの件名とメッセージを検索して、それらを正規表現で表示するかどうかを決定します。 **[File]** 列に記された Outlook アドインは、正規表現での一致を試みている最中に繰り返しタイムアウトしたため、Outlook によって無効にされました。|
|18|アドインは正常に終了しました||ホスト アプリケーションによって Office アドインが正常に閉じられました。|
|19|アドインで実行時エラーが発生しました|重大|Office アドインに、エラーの原因となる問題がありました。詳細については、エラーが発生したコンピューター上で Windows イベント ビューアーを使用して  **Microsoft Office Alerts** ログを確認してください。|
|20|アドインでライセンスを確認できませんでした|重大|Office アドインのライセンス情報を確認できないか、有効期限が切れている可能性があります。詳細については、エラーが発生したコンピューター上で Windows イベント ビューアーを使用して  **Microsoft Office Alerts** ログを確認してください。|
詳細については、「[テレメトリ ダッシュボードを展開する](http://msdn.microsoft.com/en-us/library/f69cde72-689d-421f-99b8-c51676c77717%28Office.15%29.aspx)」および「 [テレメトリ ログを使用した Office ファイルおよびカスタム ソリューションのトラブルシューティング](http://msdn.microsoft.com/library/ef88e30e-7537-488e-bc72-8da29810f7aa%28Office.15%29.aspx)」を参照してください。


## 設計および実装上のテクニック


CPU 使用率、メモリ使用量、クラッシュ許容度、UI の応答性に対するリソース制限は、リッチ クライアント上で実行される Office アドインにのみ適用されますが、サポートするすべてのクライアントおよびデバイス上でアドインが十分なパフォーマンスを発揮するためには、これらのリソース使用量およびバッテリーの使用量を最適化することが重要になります。アドインで長時間実行される処理があったり、大規模なデータ セットを処理したりする場合は、最適化が特に重要です。ここでは、CPU 使用率の高い操作やデータを大量に処理する操作を小さなチャンクに分割して、アドインで過度にリソースが消費されることを回避し、ホスト アプリケーションの応答性が保たれるようにするためのテクニックをいくつか紹介します。


- 制限のないデータセットからの大量のデータをアドインで読み取る必要があるシナリオでは、テーブルからデータを読み取る場合にページ付けを適用したり、またはより小さいサイズの読み取り操作に分割して 1 回の操作で処理するデータ量を小さくし、1 回の操作ですべてのデータを読み取ることがないようにします。 
    
    時間がかかる可能性がある操作と制限のないデータで CPU を集中的に使用する一連の入出力操作を解除することを示す、JavaScript および jQuery コード サンプルについては、「[How can I give control back (briefly) to the browser during intensive JavaScript processing?](http://stackoverflow.com/questions/210821/how-can-i-give-control-back-briefly-to-the-browser-during-intensive-javascript)」 (集中的な JavaScript の処理中にコントロールをブラウザーに戻す方法) を参照してください。 この例では、入出力の期間を制限するために、グローバル オブジェクトの [setTimeout](http://msdn.microsoft.com/en-us/library/ie/ms536753%28v=vs.85%29.aspx) メソッドを使用します。 また、ランダムな制限のないデータの代わりに、定義されたチャンク内のデータも処理します。
    
- アドインで CPU 使用率の高いアルゴリズムを使用して大量のデータを処理する場合は、Web Workers を使用してバックグラウンドで時間のかかるタスクを実行しつつ、フォアグラウンドで別のスクリプト (ユーザー インターフェイスへの進行状況の表示など) を実行できます。Web Workers は、ユーザー アクティビティをブロックせず、HTML ページの応答性を維持します。Web Workers の例については、「 [ウェブ ワーカーの基本](http://www.mdl5rocks.com/en/tutorials/workers/basics/)」を参照してください。Internet Explorer Web Workers API の詳細については、「 [Web Workers](http://msdn.microsoft.com/en-us/library/IE/hh772807%28v=vs.85%29.aspx)」を参照してください。
    
- アドインで CPU 使用率の高いアルゴリズムを使用しているが、データの入出力を小さなセットに分割できる場合は、Web サービスの作成を検討します。データを Web サービスに渡して CPU の負荷をオフロードし、非同期コールバックを待機します。
    
- 想定する最大量のデータでアドインをテストして、アドインにおける処理をその最大量までに制限します。
    

## その他のリソース



- [Office アドインのプライバシーとセキュリティ](../../docs/develop/privacy-and-security.md)
    
- [Outlook アドインのアクティブ化と JavaScript API の制限](../outlook/limits-for-activation-and-javascript-api-for-outlook-add-ins.md)
    