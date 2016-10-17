
# <a name="compare-outlook-add-in-support-in-outlook-for-mac-with-other-outlook-hosts"></a>Outlook for Mac と他の Outlook ホストの Outlook アドイン サポートの比較

Outlook for Mac でも、Outlook for Windows、デバイス用 OWA、および Outlook Web App などの他のホストで行うのと同様に Outlook アドインを作成して実行することができ、ホストごとに JavaScript をカスタマイズする必要はありません。通常、アドインから JavaScript API for Office に対する同じ呼び出しは、以下の表に示す領域を除き同様の動作をします。

 >**メモ**  Outlook for Mac は、Outlook の閲覧モードでのみ JavaScript API for Office をサポートします。

|**項目**|**Outlook for Windows、デバイス用 OWA、Outlook Web App**|**Outlook for Mac**|
|:-----|:-----|:-----|
|サポート対象バージョンの office.js および Office アドインのマニフェスト スキーマ|Office.js および スキーマ v1.1 のすべての API。|<ul><li>閲覧モードで適用可能な API のみ。office.js v1.1 の新しい API や拡張された API を使用するアドインはアクティブ化できますが、新規作成モード用の API は Outlook for Mac では正しく実行されません。 </li><li>スキーマ v1.1。</li></ul>|
|定期的な予定系列のインスタンス|<ul><li>定期的な系列のマスター予定または予定インスタンスのアイテム ID および他のプロパティを取得できます。 </li><li>[mailbox.displayAppointmentForm](../../reference/outlook/Office.context.mailbox.md#displayappointmentformitemid) を使用して、定期的な系列のインスタンスまたはマスターを表示できます。</li></ul>|<ul><li>マスター予定のアイテム ID と他のプロパティを取得できますが、定期的な系列のインスタンスのアイテム ID とプロパティは取得できません。</li><li>定期的な系列のマスター予定を表示できます。アイテム ID がない場合、定期的な系列のインスタンスは表示できません。</li></ul>|
|予定出席者の受信者の種類|[EmailAddressDetails.recipientType](../../reference/outlook/simple-types.md) を使用して、出席者の受信者の種類を特定できます。|**EmailAddressDetails.recipientType** は、予定出席者に対して **undefined** を返します。|
|ホストのバージョン文字列 |実際のホストの種類によって異なる [diagnostics.hostVersion](../../reference/outlook/Office.context.mailbox.diagnostics.md) が返すバージョン文字列の形式。例:<ul><li>Outlook for Windows:15.0.4454.1002</li><li>Outlook Web App:15.0.918.2</li></ul>|Outlook for Mac 上の  **Diagnostics.hostVersion** によって戻されるバージョン文字列の例: 15.0 (140325)|
|アイテムのカスタム プロパティ|ネットワークが使用できなくなっても、アドインはキャッシュに入っているカスタム プロパティに引き続きアクセスできます。|Outlook for Mac はカスタム プロパティをキャッシュに入れないので、ネットワークが使用できなくなると、アドインはそれらのプロパティにはアクセスできなくなります。|
|添付ファイルの詳細|ホストの種類によって異なる [AttachmentDetails](../../reference/outlook/Office.context.mailbox.md) オブジェクト内のコンテンツ タイプと添付ファイルの名前。<ul><li><b>AttachmentDetails.contentType</b> の JSON の例: <b>"contentType": "image/x-png"</b>。 </li><li><b>AttachmentDetails.name</b> にはファイル名拡張子は含まれません。たとえば、添付ファイルが「RE: Summer activity」という件名のメッセージの場合、添付ファイル名を表す JSON オブジェクトは <b>"name": "RE: Summer activity"</b> になります。</li></ul>|<ul><li><b>AttachmentDetails.contentType</b> の JSON の例: <b>"contentType": "image/png"</b></li><li><b>AttachmentDetails.name</b> には、ファイル名拡張子が必ず含まれます。メール アイテムの添付ファイルの拡張子は .eml で、予定の拡張子は .ics です。添付ファイルが「RE: Summer activity」という件名の電子メールである場合、その添付ファイル名を表す JSON オブジェクトは <b>"name": "RE: Summer activity.eml"</b> になります。</li></ul>|
|**dateTimeCreated** プロパティおよび **dateTimeModified** プロパティでのタイム ゾーンを表す文字列|例: Thu Mar 13 2014 14:09:11 GMT+0800 (中国標準時)|例: Thu Mar 13 2014 14:09:11 GMT+0800 (CST)|
|**dateTimeCreated** および **dateTimeModified** の時間の精度|次に示すコードをアドインで使用している場合、最大の精度はミリ秒単位になります。<br/><pre lang="javascript">JSON.stringify(Office.context.mailbox.item, null, 4);</pre>|精度は最大で秒単位となります。|

## <a name="additional-resources"></a>その他のリソース



- [テスト用に Outlook アドインを展開してインストールする](../outlook/testing-and-tips.md)
    
