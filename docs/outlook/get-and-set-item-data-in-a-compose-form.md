
# <a name="get-and-set-item-data-in-a-compose-form-in-outlook"></a>Outlook で新規作成フォームのアイテム データを取得および設定する
新規作成のシナリオで、受信者、件名、本文、予定の場所と時刻を含む Outlook アドインのアイテムのさまざまなプロパティを取得または設定する方法について説明します。




## <a name="getting-and-setting-item-properties-for-a-compose-add-in"></a>新規作成アドインの item プロパティの取得と設定


新規作成フォームでは、閲覧フォームで公開されているのと同じ種類のプロパティのほとんど (出席者、受信者、件名、本文など) を取得でき、さらに、閲覧フォームではなく新規作成フォームのみに関連する少数の追加プロパティ (本文、bcc) を取得できます。 

これらのプロパティのほとんどで、Outlook アドインとユーザーはユーザー インターフェイスの同じプロパティを同時に変更できるため、プロパティの取得と設定のメソッドは非同期になっています。表 1 に、アイテムレベルのプロパティ、および新規作成フォームでそれらのプロパティの取得と設定を行う関連する非同期メソッドを示します。[item.itemType](../../reference/outlook/Office.context.mailbox.item.md) プロパティと [item.conversationId](../../reference/outlook/Office.context.mailbox.item.md) プロパティは、ユーザーが変更できないため例外です。閲覧フォームの場合と同様に、新規作成フォームでも、直接親オブジェクトからプログラムを使用してプロパティを取得できます。

JavaScript API for Office で item プロパティにアクセスする以外に、Exchange Web Services (EWS) を使用してアイテム レベルのプロパティにアクセスすることができます。 **ReadWriteMailbox** アクセス許可があれば、 [mailbox.makeEwsRequestAsync](../../reference/outlook/Office.context.mailbox.md) メソッドを使用して EWS 操作の [GetItem](http://msdn.microsoft.com/en-us/library/e3590b8b-c2a7-4dad-a014-6360197b68e4%28Office.15%29.aspx) および [UpdateItem](http://msdn.microsoft.com/en-us/library/5d027523-e0bc-4da2-b60b-0cb9fc1fdfe4%28Office.15%29.aspx) にアクセスし、ユーザーのメールボックスのアイテムのさらに多くのプロパティを取得および設定できます。 **makeEwsRequestAsync** は新規作成フォームと閲覧フォームの両方で利用できます。 **ReadWriteMailbox** アクセス許可、Office アドイン プラットフォームを経由した EWS へのアクセスの詳細については、「 [ユーザーのメールボックスにアクセスする Outlook アドインのためのアクセス許可を指定する](../outlook/understanding-outlook-add-in-permissions.md)」および「 [Outlook アドインから Web サービスを呼び出す](../outlook/web-services.md)」を参照してください。


**表 1.新規作成フォームで item プロパティを取得または設定するための非同期メソッド**


|**プロパティ**|**プロパティの種類**|**取得する非同期メソッド**|**設定する非同期メソッド**|
|:-----|:-----|:-----|:-----|
|[bcc](../../reference/outlook/Office.context.mailbox.item.md)|[受信者](../../reference/outlook/Recipients.md)|[Recipients.getAsync](../../reference/outlook/Recipients.md)|[Recipients.addAsync](../../reference/outlook/Recipients.md)[Recipients.setAsync](../../reference/outlook/Recipients.md)|
|[body](../../reference/outlook/Office.context.mailbox.item.md)|[本文](../../reference/outlook/Body.md)|[Body.getAsync](../../reference/outlook/Body.md)|[Body.prependAsync](../../reference/outlook/Body.md)[Body.setAsync](../../reference/outlook/Body.md)[Body.setSelectedDataAsync](../../reference/outlook/Body.md)|
|[cc](../../reference/outlook/Office.context.mailbox.item.md)|受信者|Recipients.getAsync|Recipients.addAsync Recipients.setAsync|
|[end](../../reference/outlook/Office.context.mailbox.item.md)|[時刻](../../reference/outlook/Time.md)|[Time.getAsync](../../reference/outlook/Time.md)|[Time.setAsync](../../reference/outlook/Time.md)|
|[location](../../reference/outlook/Office.context.mailbox.item.md)|[場所](../../reference/outlook/Location.md)|[Location.getAsync](../../reference/outlook/Location.md)|[Location.setAsync](../../reference/outlook/Location.md)|
|[optionalAttendees](../../reference/outlook/Office.context.mailbox.item.md)|受信者|Recipients.getAsync|Recipients.addAsync Recipients.setAsync|
|[requiredAttendees](../../reference/outlook/Office.context.mailbox.item.md)|受信者|Recipients.getAsync|Recipients.addAsync Recipients.setAsync|
|[start](../../reference/outlook/Office.context.mailbox.item.md)|時刻|Time.getAsync|Time.setAsync|
|[subject](../../reference/outlook/Office.context.mailbox.item.md)|[件名](../../reference/outlook/Subject.md)|[Subject.getAsync](../../reference/outlook/Subject.md)|[Subject.setAsync](../../reference/outlook/Subject.md)|
|[to](../../reference/outlook/Office.context.mailbox.item.md)|受信者|Recipients.getAsync|Recipients.addAsync Recipients.setAsync|



## <a name="additional-resources"></a>その他のリソース



- [新規作成フォーム用の Outlook アドインを作成する](../outlook/compose-scenario.md)
    
- [Outlook アドインのアクセス許可を理解する](../outlook/understanding-outlook-add-in-permissions.md)
    
- [Outlook アドインから Web サービスを呼び出す](../outlook/web-services.md)
    
- [読み取りまたは新規作成フォームの Outlook アイテム データを取得および設定する](../outlook/item-data.md)
    


