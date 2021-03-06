
# <a name="limits-for-activation-and-javascript-api-for-outlook-add-ins"></a>Outlook アドインのアクティブ化と JavaScript API の制限

Outlook アドインのユーザーに満足のいくエクスペリエンスを提供するには、特定のアクティブ化ルールと API の使用に関するガイドラインを理解し、制限の範囲内に収まるようにアドインを実装する必要があります。これらのガイドラインは、個々のアドインがアクティブ化ルールの処理や JavaScript API for Office の呼び出しのために Exchange Server や Outlook に異常に長い時間のかかる処理を要求して Outlook や他のアドインの全体的なユーザー エクスペリエンスに害が及ぶことがないようにするために用意されています。こうした制限は、アドイン マニフェストでのアクティブ化ルールの設計、カスタム プロパティ、ローミング設定、受信者、Exchange Web サービス (EWS) の要求および応答、非同期呼び出しの使用に適用されます。 

 >**メモ** アドインを Outlook リッチ クライアントで実行する場合は、そのアドインが一定のランタイム リソース使用制限の範囲内で実行されているかを確認する必要もあります。 


## <a name="limits-for-activation-rules"></a>アクティブ化ルールの制限


Outlook アドインのアクティブ化ルールを設計する際には、以下のガイドラインに従います。


- マニフェストのサイズを 256 KB までに制限します。この上限を超える場合は、Exchange メールボックス用にその Outlook アドインをインストールすることはできません。

- アドインで指定できるアクティブ化ルールの数は最大 15 です。この上限を超える場合、そのアドインはインストールできません。
    
- 選択したアイテムの本文に [ItemHasKnownEntity](http://msdn.microsoft.com/en-us/library/87e10fd2-eab4-c8aa-bec3-dff92d004d39%28Office.15%29.aspx) ルールを使用する場合、Outlook リッチ クライアントでは、このルールを本文の最初の 1 MB のみに適用し、この制限を超えた本文の残りの部分には適用しません。本文の最初の 1 MB の後にしか一致するものが存在しない場合、アドインはアクティブにはなりません。その可能性が高い場合は、アクティブ化の条件を再設計してください。
    
- 
  **ItemHasKnownEntity** ルールまたは [ItemHasRegularExpressionMatch](http://msdn.microsoft.com/en-us/library/bfb726cd-81b0-a8d5-644f-2ca90a5273fc%28Office.15%29.aspx) ルールで正規表現を使用する場合、任意の Outlook ホストに通常適用される以下の制限やガイドライン、および表 1、2、3 で説明するホストごとに異なる制限やガイドラインに留意してください。
    
      - アドインで指定できるアクティブ化ルールの正規表現は最大 5 つです。その制限を超えた場合は、アドインをインストールできません。
    
  - 予期される結果が **getRegExMatches** メソッド呼び出しによって返されて、それらが最初の 50 件以内に収まるように、正規表現を指定します。
    
  - 正規表現で前読みアサーションは指定できますが、後読み (?<=text) および否定後読み (?<!text) アサーションは指定できません。
    

表 1 に制限事項を示します。また、Outlook リッチ クライアントと、Outlook Web App またはデバイス用 OWA の間での正規表現のサポートの違いについても説明します。デバイスやアイテムの本文の種類によってサポートが異なることはありません。


 **表 1.正規表現のサポートの一般的な違い**


|**Outlook リッチ クライアント**|**Outlook Web App またはデバイス用 OWA**|
|:-----|:-----|
|Visual Studio の標準テンプレート ライブラリの一部として提供されている C++ 正規表現エンジンを使用します。このエンジンは ECMAScript 5 標準に準拠しています。 |JavaScript の一部である正規表現評価を使用します。これはブラウザーによって提供されるものであり、ECMAScript 5 のスーパーセットをサポートしています。|
|RegEx エンジンが異なるため、RegEx に定義済み文字クラスに基づくカスタム文字クラスが含まれていると、Outlook リッチ クライアントの結果と Outlook Web App またはデバイス用 OWA の結果では、ばらつきが生じることがあります。<br/><br/>たとえば、正規表現 "[\s\S]{0,100}" は、単一の文字列 (空白文字または空白文字以外) の任意の数字 (0 から 100) と一致します。この正規表現が返す結果は、Outlook リッチ クライアントと Outlook Web App またはデバイス用 OWA とでは異なります。回避策として、この正規表現を "(\s|\S){0,100}" に書き換える必要があります。このように書き換えると、空白文字または空白文字以外の任意の数字 (0 から 100) と一致します。<br/><br/>各正規表現は Outlook ホストごとに入念にテストしてください。安定した結果が得られない場合は正規表現を書き換えてください。 |各正規表現は Outlook ホストごとに入念にテストしてください。安定した結果が得られない場合は正規表現を書き換えてください。|
|既定では、アドインのすべての正規表現の評価は 1 秒に制限されています。この制限を超えると、再評価が最大 3 回試行されます。この再評価の制限を超えると、Outlook リッチ クライアントは、すべての Outlook ホストで同じメールボックスに対するアドインの実行を無効にします。<br/><br/>管理者は、**OutlookActivationAlertThreshold** レジストリ キーおよび **OutlookActivationManagerRetryLimit** レジストリ キーを使用して、これらの評価制限を変更できます。|Outlook リッチ クライアントと同じリソース監視設定およびレジストリ設定はサポートしていません。しかし、正規表現を使用するアドインで、Outlook リッチ クライアントでの評価に過剰な時間がかかるアドインは、すべての Outlook ホストで同じメールボックスに対して無効にされます。|

表 2 に制限事項を示します。また、Outlook のそれぞれで正規表現を適用するアイテムの本文での違いについても説明します。正規表現がアイテムの本文に適用される場合、デバイスやアイテムの本文の種類によって制限が異なることがあります。

**表 2評価対象アイテムの本文のサイズ制限**


||**Outlook リッチ クライアント**|**Outlook Web App、デバイス用 OWA、OWA for iPad、または OWA for iPhone**|**Outlook Web App**|
|:-----|:-----|:-----|:-----|
|フォーム ファクター|サポートされる任意のデバイス。|Android スマートフォン、iPad、または iPhone|Android スマートフォン、iPad、および iPhone 以外のサポートされている任意のデバイス|
|プレーン テキスト アイテムの本文|RegEx は本文のデータの最初の 1 MB に適用されますが、制限を超える残りの本文には適用されません。|本文が 16,000 文字未満の場合にのみアドインがアクティブ化されます。|本文が 500,000 文字未満の場合にのみアドインがアクティブ化されます。|
|HTML アイテムの本文|RegEx は本文のデータの最初の 512 KB に適用されますが、制限を超える残りの本文には適用されません (実際の文字数はエンコードによって異なり、1 文字あたり 1 ～ 4 バイトの範囲でばらつくことがあります)。|RegEx は最初の 64,000 文字 (HTML タグ文字を含む) に適用されますが、制限を超える残りの本文には適用されません。|本文が 500,000 文字未満の場合にのみアドインがアクティブ化されます。|

表 3 に制限事項を示します。また、正規表現の評価後に Outlook ホストのそれぞれから返る一致の違いについても説明します。デバイスの種類によってサポートが異なることはありませんが、正規表現がアイテムの本文に適用される場合は、アイテムの本文の種類によって異なることがあります。

**表 3返される一致の制限**


||**Outlook リッチ クライアント**|**Outlook Web App またはデバイス用 OWA**|
|:-----|:-----|:-----|
|一致が返される順序|**getRegExMatches** が同じアイテムに適用された同じ正規表現に対して返す一致は、Outlook リッチ クライアントと Outlook Web App または デバイス用 OWA とで異なることが予期されます。|**getRegExMatches** が一致を返す順序は、Outlook リッチ クライアントと Outlook Web App または デバイス用 OWA とで異なることが予期されます。|
|プレーン テキスト アイテムの本文|**getRegExMatches** は、最大 1,536 (1.5 KB) 文字の一致する文字列を 50 個まで返します。<br/><br/>**メモ**: **getRegExMatches** から一致した文字列が返される場合、戻り値の配列における順序は不定です。一般的に、Outlook リッチ クライアントと、Outlook Web App およびデバイス用 OWA で同じアイテムに対して同じ正規表現を適用しても、配列における一致した文字列の順序は異なったものになります。|**getRegExMatches** は、最大 3,072 (3 KB) 文字の一致する文字列を 50 個まで返します。|
|HTML アイテムの本文|**getRegExMatches** は、最大 3,072 (3 KB) 文字の一致する文字列を 50 個まで返します。<br/> <br/> **メモ**: **getRegExMatches** から一致した文字列が返される場合、戻り値の配列における順序は不定です。一般的に、Outlook リッチ クライアントと、Outlook Web App およびデバイス用 OWA で同じアイテムに対して同じ正規表現を適用しても、配列における一致した文字列の順序は異なったものになります。|**getRegExMatches** は、最大 3,072 (3 KB) 文字の一致する文字列を 50 個まで返します。|

## <a name="limits-for-javascript-api"></a>JavaScript API の制限


アクティブ化ルールに関する上記のガイドラインとは別に、表 4 に示されているように、JavaScript オブジェクト モデルでは、Outlook ホストのそれぞれによって特定の制限も適用されます。


**表 4JavaScript API for Office を使用して特定のデータを取得または設定する際の制限**


|**機能**|**制限**|**関連する API**|**説明**|
|:-----|:-----|:-----|:-----|
|カスタム プロパティ|2,500 文字|[CustomProperties](../../reference/outlook/CustomProperties.md) オブジェクト<br/> <br/>[item.loadCustomPropertiesAsync](../../reference/outlook/Office.context.mailbox.item.md) メソッド|予定またはメッセージのアイテムのすべてのカスタム プロパティに関する制限です。アドインのすべてのカスタム プロパティの合計サイズがこの上限を超える場合は、すべての Outlook ホストがエラーを返します。|
|ローミングの設定|文字数 32 KB|[RoamingSettings](../../reference/outlook/RoamingSettings.md) オブジェクト<br/><br/> [context.roamingSettings](../../reference/outlook/Office.context.md) プロパティ|アドインのすべてのローミング設定に関する制限です。設定値がこの上限を超える場合は、すべての Outlook ホストがエラーを返します。|
|よく知られているエンティティの抽出|2,000 文字|[item.getEntities](../../reference/outlook/Office.context.mailbox.item.md) メソッド<br/> <br/>[item.getEntitiesByType](../../reference/outlook/Office.context.mailbox.item.md) メソッド<br/> <br/>[item.getFilteredEntitiesByName](../../reference/outlook/Office.context.mailbox.item.md) メソッド|Exchange Server がアイテム本文の既知のエンティティを抽出する際の制限です。Exchange Server はその制限を超えるエンティティを無視します。この制限は、アドインが  **ItemHasKnownEntity** ルールを使用するかどうかに関係ありません。|
|Exchange Web サービス|文字数 1 MB|[mailbox.makeEwsRequestAsync](../../reference/outlook/Office.context.mailbox.md) メソッド|**Mailbox.makeEwsRequestAsync** 呼び出しの要求または応答に関する制限。|
|受信者|100 の受信者|[item.requiredAttendees](../../reference/outlook/Office.context.mailbox.item.md) プロパティ<br/> <br/>[item.optionalAttendees](../../reference/outlook/Office.context.mailbox.item.md) プロパティ<br/> <br/>[item.resources](../../reference/outlook/Office.context.mailbox.item.md) property<br/> <br/>[item.to](../../reference/outlook/Office.context.mailbox.item.md) プロパティ<br/> <br/>[item.cc](../../reference/outlook/Office.context.mailbox.item.md) プロパティ<br/> <br/>[Recipients.addAsync](../../reference/outlook/Recipients.md) メソッド<br/> <br/>[Recipient.getAsync](../../reference/outlook/Recipients.md) メソッド<br/> <br/>[Recipient.setAsync](../../reference/outlook/Recipients.md) メソッド|各プロパティで指定された受信者の制限|
|表示名|255 文字|[EmailAddressDetails.displayName](../../reference/outlook/simple-types.md) プロパティ<br/><br/> [Recipients](../../reference/outlook/Recipients.md) オブジェクト<br/><br/> **item.requiredAttendees** プロパティ<br/><br/> **item.optionalAttendees** プロパティ <br/><br/>**item.resources** property <br/><br/>**item.to** プロパティ <br/><br/>**item.cc** プロパティ|予定やメッセージの表示名の長さの制限。|
|件名の設定|255 文字|[mailbox.displayNewAppointmentForm](../../reference/outlook/Office.context.mailbox.md) メソッド<br/><br/> [Subject.setAsync](../../reference/outlook/Subject.md) メソッド|新しい予定のフォームの件名、または予定やメッセージの件名の設定に関する制限。|
|場所の設定|255 文字|[Location.setAsync](../../reference/outlook/Location.md) メソッド|予定または会議出席依頼の場所の設定に関する制限。|
|新しい予定のフォームの本文|文字数 32 KB|**Mailbox.displayNewAppointmentForm** メソッド|新しい予定のフォームでの本文に関する制限。|
|既存のアイテムの本文の表示|文字数 32 KB|[mailbox.displayAppointmentForm](../../reference/outlook/Office.context.mailbox.md) メソッド<br/><br/> [mailbox.displayMessageForm](../../reference/outlook/Office.context.mailbox.md) メソッド|Outlook Web App と デバイス用 OWA の場合: 既存の予定やメッセージ フォームの本文に関する制限。|
|本文の設定。|文字数 1 MB|[Body.prependAsync](../../reference/outlook/Body.md) メソッド<br/> <br/>[Body.setAsync](../../reference/outlook/Body.md)<br/><br/>[Body.setSelectedDataAsync](../../reference/outlook/Body.md) メソッド|予定またはメッセージ アイテムの本文の設定に関する制限。|
|添付ファイルの数|Outlook Web App および デバイス用 OWA で 499 ファイル|[item.addFileAttachmentAsync](../../reference/outlook/Office.context.mailbox.item.md) メソッド|アイテムに添付して送信できるファイルの数に関する制限。Outlook Web App と デバイス用 OWA では、ユーザー インターフェイスと  **addFileAttachmentAsync** によって通常は添付ファイルが 499 ファイルまでに制限されます。Outlook リッチ クライアントでは、添付ファイルの数は具体的に制限されません。ただし、すべての Outlook ホストがユーザーの Exchange Server に構成された添付ファイルのサイズの制限に従います。「添付ファイルのサイズ」は次の行を参照してください。|
|添付ファイルのサイズ|Exchange Server に依存|**item.addFileAttachmentAsync** メソッド|アイテムの添付ファイルのサイズには、管理者がユーザーのメールボックスの Exchange Server で構成できる制限があります。Outlook リッチ クライアントでは、これによってアイテムの添付ファイルの数が制限されます。Outlook Web App とデバイス用 OWA では、2 つの制限 (添付ファイルの数とすべての添付ファイルのサイズ) の小さい方によってアイテムの実際の添付ファイル数が制限されます。|
|添付ファイルのファイル名|255 文字|**item.addFileAttachmentAsync** メソッド|アイテムに追加する添付ファイルのファイル名の長さを制限します。|
|添付ファイルの URI|2048 文字|**item.addFileAttachmentAsync** メソッド|アイテムに添付ファイルとして追加するファイル名の URI に関する制限。|
|添付ファイルの ID|100 文字|[item.addItemAttachmentAsync](../../reference/outlook/Office.context.mailbox.item.md) メソッド<br/><br/> [item.removeAttachmentAsync](../../reference/outlook/Office.context.mailbox.item.md) メソッド|アイテムに追加またはアイテムから削除する添付ファイルの ID の長さに関する制限。|
|非同期呼び出し|3 つの呼び出し|**item.addFileAttachmentAsync** メソッド<br/><br/>**item.addItemAttachmentAsync** メソッド<br/><br/><br/>**item.removeAttachmentAsync** メソッド<br/><br/> [Body.getTypeAsync](../../reference/outlook/Body.md) メソッド<br/><br/>**Body.prependAsync** メソッド<br/><br/>**Body.setSelectedDataAsync** メソッド<br/><br/> [CustomProperties.saveAsync](../../reference/outlook/CustomProperties.md) メソッド<br/><br/><br/> [item.LoadCustomPropertiesAysnc](../../reference/outlook/Office.context.mailbox.item.md) メソッド<br/><br/><br/> [Location.getAsync](../../reference/outlook/Location.md) メソッド<br/><br/>**Location.setAsync** メソッド<br/><br/> [mailbox.getCallbackTokenAsync](../../reference/outlook/Office.context.mailbox.md) メソッド<br/><br/> [mailbox.getUserIdentityTokenAsync](../../reference/outlook/Office.context.mailbox.md) メソッド<br/><br/> [mailbox.makeEwsRequestAsync](../../reference/outlook/Office.context.mailbox.md) メソッド<br/><br/>**Recipients.addAsync** メソッド<br/><br/> [Recipients.getAsync](../../reference/outlook/Recipients.md) メソッド<br/><br/>**Recipients.setAsync** メソッド<br/><br/> [RoamingSettings.saveAsync](../../reference/outlook/RoamingSettings.md) メソッド<br/><br/> [Subject.getAsync](../../reference/outlook/Subject.md) method<br/><br/>**Subject.setAsync** メソッド<br/><br/> [Time.getAsync](../../reference/outlook/Time.md) method<br/><br/> [Time.setAsync](../../reference/outlook/Time.md) メソッド|Outlook Web App または デバイス用 OWA の場合: 任意の時点で同時に行える非同期呼び出しの上限数。これは、ブラウザーによってサーバーに対する非同期呼び出しの数が制限されているためです。 |

## <a name="additional-resources"></a>その他のリソース



- [テスト用に Outlook アドインを展開してインストールする](../outlook/testing-and-tips.md)
    
- [Outlook アドインに関するプライバシー、アクセス許可、セキュリティ](../outlook/../../docs/develop/privacy-and-security.md)
    
