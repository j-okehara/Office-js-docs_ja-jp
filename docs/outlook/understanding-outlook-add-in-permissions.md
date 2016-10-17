
# <a name="understanding-outlook-add-in-permissions"></a>ユーザーのメールボックスにアクセスする Outlook アドインのためのアクセス許可を指定する

Outlook アドインでは、マニフェストに必要なアクセス許可のレベルを指定します。使用可能なレベルは  **Restricted**、 **ReadItem**、 **ReadWriteItem**、 **ReadWriteMailbox** です。これらのアクセス許可のレベルは累積的です。 **Restricted** が一番下のレベルで、上の各レベルには下のレベルのアクセス許可がすべて含まれます。 **ReadWriteMailbox** には、サポートするアクセス許可がすべて含まれます。

メール アドインが要求するアクセス許可を、Office ストア からメール アドインをインストールする前に表示できます。Exchange 管理センターで、インストールしたアドインに必要なアクセス許可を表示することもできます。


## <a name="restricted-permission"></a>制限付きアクセス許可



  **Restricted** アクセス許可は、最も基本的なレベルのアクセス許可です。マニフェストの **Permissions** 要素に [Restricted](http://msdn.microsoft.com/en-us/library/c20cdf29-74b0-564c-e178-b75d148b36d1%28Office.15%29.aspx) を指定することによってこのアクセス許可を要求できます。メール アドインのマニフェストに特定のアクセス許可の要求がない場合、Outlook はそのアドインにこのアクセス許可を既定で割り当てます。


### <a name="can-do"></a>できること


- アイテムの件名または本文から [特定のエンティティのみを取得](../outlook/match-strings-in-an-item-as-well-known-entities.md) (電話番号、アドレス、URL)。
    
- 閲覧フォームまたは新規作成フォームの現在のアイテムが特定のアイテムの種類であることを要求する [ItemIs アクティブ化ルール](../outlook/manifests/activation-rules.md#itemis-rule)を指定、または、選択したアイテムでサポートされる既知のエンティティ (電話番号、アドレス、URL) の小さなサブセットに一致する [ItemHasKnownEntity ルール](../outlook/match-strings-in-an-item-as-well-known-entities.md)を指定。
    
- ユーザーまたはアイテムに関する特定の情報に関連 **しない** プロパティーとメソッドにアクセス。(関連するメンバーのリストは、次のセクションを参照。)
    

### <a name="can't-do"></a>できないこと


- 連絡先、電子メール アドレス、会議の提案、タスクの提案のエンティティで [ItemHasKnownEntity](http://msdn.microsoft.com/en-us/library/87e10fd2-eab4-c8aa-bec3-dff92d004d39%28Office.15%29.aspx) ルールを使用。
    
- 
  [ItemHasAttachment](http://msdn.microsoft.com/en-us/library/031db7be-8a25-5185-a9c3-93987e10c6c2%28Office.15%29.aspx) ルールまたは [ItemHasRegularExpressionMatch](http://msdn.microsoft.com/en-us/library/bfb726cd-81b0-a8d5-644f-2ca90a5273fc%28Office.15%29.aspx) ルールを使用。
    
- ユーザーまたはアイテムの情報に関する次のリストに示すメンバーへのアクセス。このリストのメンバーにアクセスしようとすると、 **null** が返され、Outlook がメール アドインにアクセス許可の引き上げを要求していることを伝えるエラー メッセージが表示されます。
    
      - [item.addFileAttachmentAsync](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.addItemAttachmentAsync](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.attachments](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.bcc](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.body](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.cc](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.from](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.getRegExMatches](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.getRegExMatchesByName](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.optionalAttendees](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.organizer](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.removeAttachmentAsync](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.requiredAttendees](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.resources](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.sender](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.to](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [mailbox.getCallbackTokenAsync](../../reference/outlook/Office.context.mailbox.md)
    
  - [mailbox.getUserIdentityTokenAsync](../../reference/outlook/Office.context.mailbox.md)
    
  - [mailbox.makeEwsRequestAsync](../../reference/outlook/Office.context.mailbox.md)
    
  - [mailbox.userProfile](../../reference/outlook/Office.context.mailbox.userProfile.md)
    
  - [Body](../../reference/outlook/Body.md) およびその子メンバーすべて
    
  - [Location](../../reference/outlook/Location.md) およびその子メンバーすべて
    
  - [Recipients](../../reference/outlook/Recipients.md) およびその子メンバーすべて
    
  - [Subject](../../reference/outlook/Subject.md) およびその子メンバーすべて
    
  - [Time](../../reference/outlook/Time.md) およびその子メンバーすべて
    

## <a name="readitem-permission"></a>ReadItem アクセス許可


**ReadItem** アクセス許可は、アクセス許可モデルの次のレベルのアクセス許可です。マニフェストの **Permissions** 要素に **ReadItem** を指定すると、このアクセス許可を要求できます。


### <a name="can-do"></a>できること


- 閲覧フォームまたは [新規作成フォーム](../outlook/item-data.md)の現在のアイテムの [すべてのプロパティの読み取り](../outlook/get-and-set-item-data-in-a-compose-form.md)。たとえば、閲覧フォームの [item.to](../../reference/outlook/Office.context.mailbox.item.md) および新規作成フォームの [item.to.getAsync](../../reference/outlook/Recipients.md)。
    
- [アイテムの添付ファイルまたは完全なアイテムを取得するためにコールバック トークンを取得する](../outlook/get-attachments-of-an-outlook-item.md)。
    
- そのアイテムのアドインが設定する [カスタム プロパティの書き込み](http://msdn.microsoft.com/library/30217d63-7615-4f3f-8618-c91e4e60cd43%28Office.15%29.aspx)。
    
- アイテムの件名または本文から、サブセットだけでなく [存在する既知のエンティティをすべて取得する](../outlook/match-strings-in-an-item-as-well-known-entities.md)。
    
- 
  [ItemHasKnownEntity](../outlook/manifests/activation-rules.md#itemhasknownentity-rule) ルールの [既知のエンティティ](http://msdn.microsoft.com/en-us/library/87e10fd2-eab4-c8aa-bec3-dff92d004d39%28Office.15%29.aspx)、または [ItemHasRegularExpressionMatch](../outlook/manifests/activation-rules.md#itemhasregularexpressionmatch-rule) ルールの [正規表現](http://msdn.microsoft.com/en-us/library/bfb726cd-81b0-a8d5-644f-2ca90a5273fc%28Office.15%29.aspx)をすべて使用します。次の例は、スキーマ v1.1 に従っています。選択されたメッセージの件名または本文に既知のエンティティが 1 つ以上ある場合にアクティブ化されるルールを示しています。
    

```XML
<Permissions>ReadItem</Permissions>
    <Rule xsi:type="RuleCollection" Mode="And">
    <Rule xsi:type="ItemIs" FormType = "Read" ItemType="Message" />
    <Rule xsi:type="RuleCollection" Mode="Or">
        <Rule xsi:type="ItemHasKnownEntity" 
            EntityType="PhoneNumber" />
        <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
        <Rule xsi:type="ItemHasKnownEntity" EntityType="Url" />
        <Rule xsi:type="ItemHasKnownEntity" 
            EntityType="MeetingSuggestion" />
        <Rule xsi:type="ItemHasKnownEntity" 
            EntityType="TaskSuggestion" />
        <Rule xsi:type="ItemHasKnownEntity" 
            EntityType="EmailAddress" />
        <Rule xsi:type="ItemHasKnownEntity" EntityType="Contact" />
</Rule>
```


### <a name="can't-do"></a>できないこと

**mailbox.makeEWSRequestAsync** または次の書き込みメソッドにアクセスする。


- [item.addFileAttachmentAsync](../../reference/outlook/Office.context.mailbox.item.md)
    
- [item.addItemAttachmentAsync](../../reference/outlook/Office.context.mailbox.item.md)
    
- [item.bcc.addAsync](../../reference/outlook/Recipients.md)
    
- [item.bcc.setAsync](../../reference/outlook/Recipients.md)
    
- [item.body.prependAsync](../../reference/outlook/Body.md)
    
- [item.body.setAsync](../../reference/outlook/Body.md)
    
- [item.body.setSelectedDataAsync](../../reference/outlook/Body.md)
    
- [item.cc.addAsync](../../reference/outlook/Recipients.md)
    
- [item.cc.setAsync](../../reference/outlook/Recipients.md)
    
- [item.end.setAsync](../../reference/outlook/Time.md)
    
- [item.location.setAsync](../../reference/outlook/Location.md)
    
- [item.optionalAttendees.addAsync](../../reference/outlook/Recipients.md)
    
- [item.optionalAttendees.setAsync](../../reference/outlook/Recipients.md)
    
- [item.removeAttachmentAsync](../../reference/outlook/Office.context.mailbox.item.md)
    
- [item.requiredAttendees.addAsync](../../reference/outlook/Recipients.md)
    
- [item.requiredAttendees.setAsync](../../reference/outlook/Recipients.md)
    
- [item.start.setAsync](../../reference/outlook/Time.md)
    
- [item.subject.setAsync](../../reference/outlook/Subject.md)
    
- [item.to.addAsync](../../reference/outlook/Recipients.md)
    
- [item.to.setAsync](../../reference/outlook/Recipients.md)
    

## <a name="readwriteitem-permission"></a>ReadWriteItem アクセス許可


マニフェストの  **Permissions** 要素に **ReadWriteItem** を指定すると、このアクセス許可を要求できます。作成フォームでアクティブになり、書き込みメソッド ( **Message.to.addAsync** または **Message.to.setAsync**) を使用するメール アドインは、このレベル以上のアクセス許可を使用する必要があります。


### <a name="can-do"></a>できること


- Outlook で閲覧または新規作成されているアイテムの [すべてのアイテム レベルのプロパティを読み書き](../outlook/item-data.md)。
    
- そのアイテムで [添付ファイルを追加または削除](../outlook/add-and-remove-attachments-to-an-item-in-a-compose-form.md)。
    
- JavaScript API for Office の中でメール アドインに適用される、 **Mailbox.makeEWSRequestAsync** を除くほかのすべてのメンバーの使用。
    

### <a name="can't-do"></a>できないこと

Use  **Mailbox.makeEWSRequestAsync**.


## <a name="readwritemailbox-permission"></a>ReadWriteMailbox アクセス許可


**ReadWriteMailbox** アクセス許可は、最上位レベルのアクセス許可です。マニフェストの **Permissions** 要素に **ReadWriteMailbox** を指定すると、このアクセス許可を要求できます。

**ReadWriteItem** アクセス許可のサポートのほかに **Mailbox.makeEWSRequestAsync** を使用することで、サポートされる Exchange Web Services (EWS) の操作にアクセスして、次の作業を行うことができます。


- ユーザーのメール ボックスのアイテムのすべてのプロパティーの読み取りと書き込み。
    
- そのメール ボックスのフォルダーまたはアイテムの作成、読み取り、書き込み。
    
- そのメール ボックスからのアイテムの送信。
    
**mailbox.makeEWSRequestAsync** を使用して、次の EWS の操作にアクセスできます。


- 
  [CopyItem](http://msdn.microsoft.com/en-us/library/bcc68f9e-d511-4c29-bba6-ed535524624a%28Office.15%29.aspx)
    
- 
  [CreateFolder](http://msdn.microsoft.com/en-us/library/6f6c334c-b190-4e55-8f0a-38f2a018d1b3%28Office.15%29.aspx)
    
- 
  [CreateItem](http://msdn.microsoft.com/en-us/library/78a52120-f1d0-4ed7-8748-436e554f75b6%28Office.15%29.aspx)
    
- 
  [FindConversation](http://msdn.microsoft.com/en-us/library/2384908a-c203-45b6-98aa-efd6a4c23aac%28Office.15%29.aspx)
    
- 
  [FindFolder](http://msdn.microsoft.com/en-us/library/7a9855aa-06cc-45ba-ad2a-645c15b7d031%28Office.15%29.aspx)
    
- 
  [FindItem](http://msdn.microsoft.com/en-us/library/ebad6aae-16e7-44de-ae63-a95b24539729%28Office.15%29.aspx)
    
- 
  [GetConversationItems](http://msdn.microsoft.com/en-us/library/8ae00a99-b37b-4194-829c-fe300db6ab99%28Office.15%29.aspx)
    
- 
  [GetFolder](http://msdn.microsoft.com/en-us/library/355bcf93-dc71-4493-b177-622afac5fdb9%28Office.15%29.aspx)
    
- 
  [GetItem](http://msdn.microsoft.com/en-us/library/e3590b8b-c2a7-4dad-a014-6360197b68e4%28Office.15%29.aspx)
    
- 
  [MarkAsJunk](http://msdn.microsoft.com/en-us/library/1f71f04d-56a9-4fee-a4e7-d1034438329e%28Office.15%29.aspx)
    
- 
  [MoveItem](http://msdn.microsoft.com/en-us/library/dcf40fa7-7796-4a5c-bf5b-7a509a18d208%28Office.15%29.aspx)
    
- 
  [SendItem](http://msdn.microsoft.com/en-us/library/337b89ef-e1b7-45ed-92f3-8abe4200e4c7%28Office.15%29.aspx)
    
- 
  [UpdateFolder](http://msdn.microsoft.com/en-us/library/3494c996-b834-4813-b1ca-d99642d8b4e7%28Office.15%29.aspx)
    
- 
  [UpdateItem](http://msdn.microsoft.com/en-us/library/5d027523-e0bc-4da2-b60b-0cb9fc1fdfe4%28Office.15%29.aspx)
    
サポートされていない操作を使用すると、エラーが返されます。


## <a name="additional-resources"></a>その他のリソース



- [Outlook アドインに関するプライバシー、アクセス許可、セキュリティ](../outlook/../../docs/develop/privacy-and-security.md)
    
- [Outlook アイテム内の文字列を既知のエンティティとして照合する](../outlook/match-strings-in-an-item-as-well-known-entities.md)
    
