
# Rule 要素
このメール アドインに対して評価すべきアクティブ化ルールを指定します。

 **アドインの種類:**メール


## 構文:

 **ItemIs ルール** - 選んだアイテムが指定した型である場合に true と評価するルールを定義します。


```XML
<Rule xsi:type="ItemIs" 
   ItemType= ["Appointment" | "Message"]
   FormType=["Read" | "Edit" | "ReadOrEdit"] 
   ItemClass = "string " 
   IncludeSubClasses=["true" | "false"] />
```

 **ItemHasAttachment ルール** - アイテムに添付ファイルがある場合に true と評価するルールを定義します。




```XML
<Rule xsi:type="ItemHasAttachment"  />
```

 **ItemHasKnownEntity** - 指定したエンティティ型のテキストがアイテムの件名または本文に含まれている場合に true と評価するルールを定義します。




```XML
<Rule xsi:type="ItemHasKnownEntity" 
  EntityType=["MeetingSuggestion" | "TaskSuggestion" |"Address" | "Url" | "PhoneNumber" | "EmailAddress" | "Contact" ]
  RegExFilter="string "
  FilterName="string "
  IgnoreCase=["true | false"]/>
```

 **ItemHasRegularExpressionMatch ルール** - アイテムの指定したプロパティの中を検索し、指定した正規表現と一致するものがある場合に true と評価するルールを定義します。




```XML
<Rule xsi:type="ItemHasRegularExpressionMatch" 
    RegExName="string " 
    RegExValue="string " 
    PropertyName=["Subject" | "BodyAsPlaintext" | "BodyAsHtml" | "SenderSTMPAddress"]
    IgnoreCase=["true" | "false"]
/>
```

 **RuleCollection ルール** - ルールのコレクション、およびそれらのルールの評価時に使用する論理演算子を定義します。




```XML
<Rule xsi:type="RuleCollection" Mode=["And" | "Or"]>
   ...
</Rule>
```


## 次に含まれる:

 _[OfficeApp](../../reference/manifest/officeapp.md)_


## 属性:

 **ItemIs ルールの属性**



|**属性**|**種類**|**必須**|**説明**|
|:-----|:-----|:-----|:-----|
|ItemType|ItemType (文字列)|必須出席者|照合するアイテムの種類を指定します。次のいずれかを指定できます。

|**ItemType**|**対応する ItemClass**|
|:-----|:-----|
|Appointment|IPM.Appointment|
|Message(1)|メール メッセージ、会議出席依頼、返信、キャンセルが含まれます。|
|
|FormType|ItemFormType (文字列)|必須出席者|アプリがアイテムの読み取りまたは編集フォームで表示されるかどうかを指定します。 次のいずれかを指定できます。|

|**FormType**|**説明**|
|:-----|:-----|
|読み取り|(指定された **ItemType** の) 閲覧フォームでのみメール アドインをアクティブにするように指定します。|
|Edit|(指定された **ItemType** の) 作成フォームでのみメール アドインをアクティブにするように指定します。|
|ReadOrEdit|(指定された **ItemType** の) 閲覧フォームと作成フォームの両方でメール アドインをアクティブにするように指定します。|
|ItemClass|文字列|省略可能|照合するカスタム メッセージ クラスを指定します。詳細については、「[特定のメッセージ クラスに対して Outlook のメール アドインをアクティブにする](http://msdn.microsoft.com/library/f464a152-2dff-4fb3-bf98-c1a3639c3e80%28Office.15%29.aspx)」をご覧ください。|
|IncludeSubClasses|ブール値|省略可能|アイテムが指定したメッセージ クラスのサブクラスである場合に、このルールは true と評価する必要があるかどうかを指定します。既定値は false です。|


(1) 対応するメッセージ クラスは次のとおりです。IPM.NoteIPM.Schedule.Meeting.RequestIPM.Schedule.Meeting.NegIPM.Schedule.Meeting.PosIPM.Schedule.Meeting.TentIPM.Schedule.Meeting.Canceled

 **ItemHasAttachment ルールの属性**

なし。

 **ItemHasKnownEntity ルールの属性**



|**属性**|**種類**|**必須**|**説明**|
|:-----|:-----|:-----|:-----|
|EntityType|KnownEntityType (文字列)|必須|このルールが true と評価されるために見つける必要のあるエンティティの型を指定します。次のいずれかを指定できます。

|**KnownEntityType**|**説明**|
|:-----|:-----|
|MeetingSuggestion|パターン認識によって、イベントまたは会議を参照すると識別されるテキスト。|
|TaskSuggestion| パターン認識によって、実行可能なフレーズを含むと識別されるテキスト。|
|Address|パターン認識によって、米国の郵便番号と住所を参照すると識別されるテキスト。|
|Url|パターン認識によって、ファイル名または Web アドレスの URL を含むと識別されるテキスト。|
|PhoneNumber| パターン認識によって、北米の電話番号と識別される一連の数字。|
|EmailAddress|パターン認識によって、SMTP 形式の電子メールアドレスを含むと識別されるテキスト。|
|Contact|パターン認識によって、連絡先情報を含むと識別されるテキスト。|
|RegExFilter|文字列|省略可能|このエンティティに対してアクティブ化を実行するための正規表現を指定します。|
|FilterName|文字列|省略可能|正規表現フィルターの名前を指定します。指定すると、以後このフィルターをアドインのコード内で参照できます。|
|IgnoreCase|ブール値|省略可能|**RegExFilter** 属性で指定した正規表現の実行時に、大文字と小文字の違いを無視するように指定します。|
 **ItemHasRegularExpressionMatch ルールの属性**



|**属性**|**種類**|**必須**|**説明**|
|:-----|:-----|:-----|:-----|
|RegExName|文字列|必須|アドインのコードで参照できるように、正規表現の名前を指定します。|
|RegExValue|文字列|必須|メール アドインを表示するかどうかを判断するために評価する正規表現を指定します。 |
|PropertyName|PropertyName (文字列)|必須出席者|正規表現の評価対象となるプロパティの名前を指定します。次のいずれかを指定できます。

|**PropertyName**|**説明**|
|:-----|:-----|
|件名|アイテムの件名に対して正規表現を評価します。|
|BodyAsPlaintext|テキスト形式のアイテムの本文に対して正規表現を評価します。|
|BodyAsHtml|アイテムの本文が HTML 形式の場合に、その本文に対して正規表現を評価します。|
|SenderSTMPAddress|アイテムの送信者の SMTP アドレスに対して正規表現を評価します。|
|IgnoreCase|ブール値|省略可能|正規表現の実行時に大文字と小文字の違いを無視するように指定します。|
 **RuleCollection ルールの属性**



|**属性**|**種類**|**必須**|**説明**|
|:-----|:-----|:-----|:-----|
|Mode|string|必須出席者|このルール コレクションの評価時に使用する論理演算子を指定します。次のいずれかを指定できます。"And" または "Or"。|

## その他のリソース



- 「[特定のメッセージ クラスに対して Outlook のメール アドインをアクティブにする](http://msdn.microsoft.com/library/f464a152-2dff-4fb3-bf98-c1a3639c3e80%28Office.15%29.aspx)」および「[Outlook アドインのアクティブ化ルール](../../docs/outlook/manifests/activation-rules.md#activation-rules-for-outlook-add-ins)」
    
- [Outlook アイテム内の文字列を既知のエンティティとして照合する](../../docs/outlook/match-strings-in-an-item-as-well-known-entities.md)
    
- [正規表現アクティブ化ルールを使用して Outlook アドインを表示する](../../docs/outlook/use-regular-expressions-to-show-an-outlook-add-in.md)
    
