
# <a name="get-or-set-the-subject-when-composing-an-appointment-or-message-in-outlook"></a>Outlook で予定またはメッセージを作成するときに件名を取得または設定する

JavaScript API for Office には、ユーザーが作成する予定やメッセージの件名を取得および設定する非同期メソッド ([subject.getAsync](../../reference/outlook/Subject.md) と [subject.setAsync](../../reference/outlook/Subject.md)) が用意されています。これらのメソッドを使用する場合は、新規作成フォームでアドインをアクティブ化するようにアドイン マニフェストが Outlook 用に適切にセット アップされていることを確認してください。

**subject** プロパティは、予定とメッセージの新規作成フォームと閲覧フォームの両方で読み取りアクセスで利用できます。閲覧フォームでは、次の例に示すとおり、このプロパティに親オブジェクトから直接アクセスできます。




```js
item.subject
```

ただし、新規作成フォームでは、ユーザーとアドインの両方が同時に件名を挿入または変更できるため、次に示すように、非同期メソッド  **getAsync** を使用して件名を取得する必要があります。




```js
item.subject.getAsync
```

書き込みアクセスでは、 **subject** プロパティは新規作成フォームのみで利用でき、閲覧フォームでは利用できません。

JavaScript API for Office のほとんどの非同期メソッドと同じように、**getAsync** と **setAsync** はオプションの入力パラメーターを受け取ります。オプションの入力パラメーターを指定する方法の詳細については、「[Office アドインにおける非同期プログラミング](../../docs/develop/asynchronous-programming-in-office-add-ins.md)」を参照してください。


## <a name="to-get-the-subject"></a>件名を取得するには


このセクションでは、ユーザーが作成している予定またはメッセージの件名を取得して、その件名を表示するサンプル コードについて説明します。このサンプル コードは、以下に示すように、アドイン マニフェストのルールが、予定またはメッセージの新規作成フォームでアドインをアクティブにすることを想定しています。


```XML
<Rule xsi:type="RuleCollection" Mode="Or">
  <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit"/>
  <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit"/>
</Rule>

```

**item.subject.getAsync** を使用する場合は、非同期呼び出しの状態と結果を確認するコールバック メソッドを用意します。_asyncContext_ オプション パラメーターを使用して、コールバック メソッドに必要な引数を指定できます。コールバックの出力パラメーター _asyncResult_ を使用して、状態、結果およびエラーを取得できます。非同期呼び出しに成功すると、[AsyncResult.value](../../reference/outlook/simple-types.md) プロパティを使用して件名をプレーン テキスト文字列として取得できます。




```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Get the subject of the item being composed.
        getSubject();
    });
}

// Get the subject of the item that the user is composing.
function getSubject() {
    item.subject.getAsync(
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Successfully got the subject, display it.
                write ('The subject is: ' + asyncResult.value);
            }
        });
}

// Write to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## <a name="to-set-the-subject"></a>件名を設定するには


このセクションでは、ユーザーが作成している予定またはメッセージの件名を設定するサンプル コードについて説明します。前のサンプルと同様に、このサンプル コードは、アドイン マニフェストのルールが、予定またはメッセージの新規作成フォームでアドインをアクティブにすることを想定しています。

**item.subject.setAsync** を使用する場合は、データ パラメーターで最大 255 文字の文字列を指定します。オプションで、コールバック メソッドおよび _asyncContext_ パラメーターにそのコールバック メソッドの引数を指定できます。コールバックの _asyncResult_ 出力パラメーターで、状態、結果およびエラー メッセージを確認する必要があります。非同期呼び出しが成功すると、**setAsync** はそのアイテムの既存の件名を上書きして、指定された件名の文字列をプレーン テキストとして挿入します。




```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Set the subject of the item being composed.
        setSubject();
    });
}

// Set the subject of the item that the user is composing.
function setSubject() {
    var today = new Date();
    var subject;

    // Customize the subject with today's date.
    subject = 'Summary for ' + today.toLocaleDateString();

    item.subject.setAsync(
        subject,
        { asyncContext: { var1: 1, var2: 2 } },
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Successfully set the subject.
                // Do whatever appropriate for your scenario
                // using the arguments var1 and var2 as applicable.
            }
        });
}

// Write to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## <a name="additional-resources"></a>その他のリソース



- [Outlook で新規作成フォームのアイテム データを取得および設定する](../outlook/get-and-set-item-data-in-a-compose-form.md)
    
- [読み取りまたは新規作成フォームの Outlook アイテム データを取得および設定する](../outlook/item-data.md)
    
- [新規作成フォーム用の Outlook アドインを作成する](../outlook/compose-scenario.md)
    
- [Office アドインにおける非同期プログラミング](../../docs/develop/asynchronous-programming-in-office-add-ins.md)
    
- [Outlook の予定またはメッセージを作成するときに受信者を取得、設定、または追加する](../outlook/get-set-or-add-recipients.md)
    
- [Outlook で予定またはメッセージを作成するときに本文にデータを挿入する](../outlook/insert-data-in-the-body.md)
    
- [Outlook で予定を作成するときに場所を取得または設定する](../outlook/get-or-set-the-location-of-an-appointment.md)
    
- [Outlook で予定を作成するときに時刻を取得または設定する](../outlook/get-or-set-the-time-of-an-appointment.md)
    
