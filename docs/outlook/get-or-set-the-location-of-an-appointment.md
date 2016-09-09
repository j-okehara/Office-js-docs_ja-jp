
# Outlook で予定を作成するときに場所を取得または設定する

JavaScript API for Office には、ユーザーが作成する予定の場所を取得および設定する非同期メソッド ([getAsync](../../reference/outlook/Location.md) および [setAsync](../../reference/outlook/Location.md)) が用意されています。これらの非同期メソッドは、新規作成アドインでのみ使用できます。これらのメソッドを使用する場合は、新規作成フォームでアドインをアクティブ化するようアドイン マニフェストが Outlook 用に適切にセット アップされていることを確認してください。この手順については、「[新規作成フォーム用の Outlook アドインを作成する](../outlook/compose-scenario.md)」を参照してください。

[location](../../reference/outlook/Office.context.mailbox.item.md) プロパティは、予定の閲覧フォームと新規作成フォーム両方の読み取りアクセスで使用できます。閲覧フォームでは、次のように親オブジェクトからプロパティに直接アクセスできます。




```js
item.location
```

しかし、新規作成フォームでは、ユーザーとアドインの両方が同時に場所を挿入または変更できるため、以下に示すとおり、非同期メソッド  **getAsync** を使用して場所を取得する必要があります。




```js
item.location.getAsync
```

**location** プロパティは、予定の新規作成フォームでのみ書き込みアクセスに使用できます。閲覧フォームでは使用できません。

JavaScript API for Office のほとんどの非同期メソッドと同じように、**getAsync** と **setAsync** はオプションの入力パラメーターを受け取ります。このオプションの入力パラメーターを指定する方法の詳細については、「[Office アドインにおける非同期プログラミング](../../docs/develop/asynchronous-programming-in-office-add-ins.md)」を参照してください。


## 場所を取得するには


このセクションでは、ユーザーが新規作成する予定の場所を取得し、場所を表示するコード サンプルを示します。以下に示すとおり、このコード サンプルは、予定の新規作成フォームでアドインをアクティブ化するアドインのマニフェストのルールを想定しています。


```XML
<Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit"/>

```

**item.location.getAsync** を使用する場合は、非同期の呼び出しについての状態と結果を確認するコールバック メソッドを指定します。コールバック メソッドに対して必要な引数はすべて、オプションのパラメーター _asyncContext_ で指定できます。状態、結果、およびすべてのエラーは、コールバックの出力パラメーター _asyncResult_ を使用することで取得できます。非同期の呼び出しが正常に行われると、[AsyncResult.value](../../reference/outlook/simple-types.md) プロパティを使用することで、文字列として場所を取得できます。




```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Get the location of the item being composed.
        getLocation();
    });
}

// Get the location of the item that the user is composing.
function getLocation() {
    item.location.getAsync(
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Successfully got the location, display it.
                write ('The location is: ' + asyncResult.value);
            }
        });
}

// Write to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## 場所を設定するには


このセクションでは、ユーザーが新規作成する予定の場所を設定するコード サンプルを示します。前の例と同様、このコード サンプルは、予定の新規作成フォームでアドインをアクティブ化するアドインのマニフェストのルールを想定しています。

**item.location.setAsync** を使用する場合は、データ パラメーターに最大 255 文字の文字列を指定します。オプションでコールバック メソッドおよびコールバック メソッドの引数を _asyncContext_ パラメーターに指定することができます。コールバックの出力パラメーター _asyncResult_ で状態、結果およびすべてのエラー メッセージを確認してください。非同期の呼び出しが正常に行われると、**setAsync** が指定した場所をプレーン テキストの文字列として挿入し、そのアイテムの既存の場所を上書きします。




```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Check for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Set the location of the item being composed.
        setLocation();
    });
}

// Set the location of the item that the user is composing.
function setLocation() {
    item.location.setAsync(
        'Conference room A',
        { asyncContext: { var1: 1, var2: 2 } },
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Successfully set the location.
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


## その他のリソース



- [Outlook で新規作成フォームのアイテム データを取得および設定する](../outlook/get-and-set-item-data-in-a-compose-form.md)
    
- [読み取りまたは新規作成フォームの Outlook アイテム データを取得および設定する](../outlook/item-data.md)
    
- [新規作成フォーム用の Outlook アドインを作成する](../outlook/compose-scenario.md)
    
- [Office アドインにおける非同期プログラミング](../../docs/develop/asynchronous-programming-in-office-add-ins.md)
    
- [Outlook の予定またはメッセージを作成するときに受信者を取得、設定、または追加する](../outlook/get-set-or-add-recipients.md)
    
- [Outlook で予定またはメッセージを作成するときに件名を取得または設定する](../outlook/get-or-set-the-subject.md)
    
- [Outlook で予定またはメッセージを作成するときに本文にデータを挿入する](../outlook/insert-data-in-the-body.md)
    
- [Outlook で予定を作成するときに時刻を取得または設定する](../outlook/get-or-set-the-time-of-an-appointment.md)
    
