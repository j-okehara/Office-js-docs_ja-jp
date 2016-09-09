
# Outlook アイテムからエンティティ文字列を抽出する

この記事では、選択した Outlook アイテムの件名と本文に含まれる、サポートされる既知のエンティティの文字列インスタンスを抽出する **[エンティティの表示]** Outlook アドインを作成する方法について説明します。対象のアイテムは、予定、メール メッセージ、会議出席依頼、会議出席依頼の返信、または会議の取り消しです。サポートされるエンティティには次のようなものがあります。

- 住所: 米国の郵送先住所で、少なくとも番地、住所、都市名、州名、郵便番号の各要素の一部を含むもの。
    
- 連絡先: 住所、勤務先名など、他のエンティティのコンテキストにおける、個人の連絡先情報。
    
- 電子メール アドレス: SMTP 電子メール アドレス。
    
- 会議提案: イベントへの言及などの会議提案。予定ではなくメッセージのみが会議提案の抽出をサポートすることに注意してください。
    
- 電話番号: 北米の電話番号。
    
- タスクの提案: 通常、実行可能な語句で表現される、タスクの提案。
    
- URL。
    
これらのエンティティの大部分は、大量のデータの機械学習に基づいた自然言語認識を利用しています。このため、認識は非確定的であり、Outlook アイテムのコンテキストによって異なることがあります。ユーザーが予定、メール メッセージ、会議出席依頼、会議出席依頼の返信、または会議の取り消しの表示を選択するたびに、Outlook によってエンティティ アドインがアクティブ化されます。初期化時に、このサンプルのエンティティ アドインは、現在のアイテムからサポートされているエンティティのすべてのインスタンスを読み込みます。 

このアドインにはユーザーがエンティティの種類を選択するためのボタンがあります。ユーザーがエンティティを選択すると、アドインは選択されたエンティティのインスタンスをアドイン ウィンドウに表示します。後続の各セクションでは、エンティティ アドインの XML マニフェスト、HTML ファイル、および JavaScript ファイルの内容を示し、それぞれのエンティティの抽出をサポートするコードについて詳しく説明します。

## XML マニフェスト


エンティティ アドインには、論理 OR 演算で結合された 2 つのアクティブ化ルールがあります。 


```xml
<!-- Activate the add-in if the current item in Outlook is an email or appointment item. -->
<Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemIs" ItemType="Message"/>
    <Rule xsi:type="ItemIs" ItemType="Appointment"/>
</Rule>
```

これらのルールでは、閲覧ウィンドウまたは閲覧インスペクターの現在選択されているアイテムが予定またはメッセージ (電子メール メッセージ、会議出席依頼、会議出席依頼の返信、または会議の取り消しなど) であるときに、Outlook でこのアドインをアクティブ化することを指定しています。

エンティティ アドインのマニフェストを次に示します。このマニフェストは、Office アドイン マニフェストのスキーマ バージョン 1.1 を使用します。




```xml
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
xsi:type="MailApp">
  <Id>6880A140-1C4F-11E1-BDDB-0800200C9A68</Id>
  <Version>1.0</Version>
  <ProviderName>Microsoft</ProviderName>
  <DefaultLocale>EN-US</DefaultLocale>
  <DisplayName DefaultValue="Display entities"/>
  <Description DefaultValue=
     "Display known entities on the selected item."/>
  <Hosts>
    <Host Name="Mailbox" />
  </Hosts>
  <Requirements>
    <Sets DefaultMinVersion="1.1">
      <Set Name="Mailbox" />
    </Sets>
  </Requirements>
  <FormSettings>
    <Form xsi:type="ItemRead">
      <DesktopSettings>
        <!-- Change the following line to specify the web -->
        <!-- server where the HTML file is hosted. -->
        <SourceLocation DefaultValue=
          "http://webserver/default_entities/default_entities.html"/>
        <RequestedHeight>350</RequestedHeight>
      </DesktopSettings>
    </Form>
  </FormSettings>
  <Permissions>ReadItem</Permissions>
  <!-- Activate the add-in if the current item in Outlook is -->
  <!-- an email or appointment item. -->
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemIs" ItemType="Message"/>
    <Rule xsi:type="ItemIs" ItemType="Appointment"/>
  </Rule>
  <DisableEntityHighlighting>false</DisableEntityHighlighting>
</OfficeApp>
```


## HTML の実装


エンティティ アドインの HTML ファイルでは、ユーザーがエンティティの種類を選択するためのボタンと、表示されたエンティティのインスタンスを消去するためのボタンを指定しています。このファイルでは、後の「[JavaScript の実装](#javascript-の実装)」で説明する default_entities.js という JavaScript ファイルを指定しています。JavaScript ファイルには、それぞれのボタンに対するイベント ハンドラーが含まれています。

すべての Outlook アドインに office.js を含める必要があります。以下の HTML ファイルには、CDN に office.js のバージョン 1.1 が含まれます。 

```html
<!DOCTYPE html>
<html>
<head>
    <meta http-equiv="X-UA-Compatible" content="IE=Edge" >
    <title>standard_item_properties</title>
    <link rel="stylesheet" type="text/css" media="all" href="default_entities.css" />
    <script type="text/javascript" src="MicrosoftAjax.js"></script>
    <!-- Use the CDN reference to Office.js. -->
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
    <script type="text/javascript"  src="default_entities.js"></script>
</head>

<body>
    <div id="container">
        <div id="button">
        <input type="button" value="clear" 
            onclick="myClearEntitiesBox();">
        <input type="button" value="Get Addresses" 
            onclick="myGetAddresses();">
        <input type="button" value="Get Contact Information" 
            onclick="myGetContacts();">
        <input type="button" value="Get Email Addresses" 
            onclick="myGetEmailAddresses();">
        <input type="button" value="Get Meeting Suggestions" 
            onclick="myGetMeetingSuggestions();">
        <input type="button" value="Get Phone Numbers" 
            onclick="myGetPhoneNumbers();">
        <input type="button" value="Get Task Suggestions" 
            onclick="myGetTaskSuggestions();">
        <input type="button" value="Get URLs" 
            onclick="myGetUrls();">
        </div>
        <div id="entities_box"></div>
    </div>
</body>
</html>
```


## スタイル シート


エンティティ アドインでは、default_entities.css というオプションの CSS ファイルを使用して出力のレイアウトを指定しています。次に、この CSS ファイルの内容を示します。


```css
*
{
    color: #FFFFFF;
    margin: 0px;
    padding: 0px;
    font-family: Arial, Sans-serif;
}
html 
{
    scrollbar-base-color: #FFFFFF;
    scrollbar-arrow-color: #ABABAB; 
    scrollbar-lightshadow-color: #ABABAB; 
    scrollbar-highlight-color: #ABABAB; 
    scrollbar-darkshadow-color: #FFFFFF; 
    scrollbar-track-color: #FFFFFF;
}
body
{
    background: #4E9258;
}
input
{
    color: #000000;
    padding: 5px;
}
span
{
    color: #FFFF00;
}
div#container
{
    height: 100%;
    padding: 2px;
    overflow: auto;
}
div#container td
{
    border-bottom: 1px solid #CCCCCC;
}
td.property-name
{
    padding: 0px 5px 0px 0px;
    border-right: 1px solid #CCCCCC;
}
div#meeting_suggestions
{
    border-top: 1px solid #CCCCCC;
}
```


## JavaScript の実装


残りのセクションでは、このサンプル (default_entities.js ファイル) を使用して、ユーザーが表示中のメッセージまたは予定の件名と本文から一般的なエンティティを抽出する方法について説明します。 


## 初期化時のエンティティの抽出


[Office.initialize](../../reference/shared/office.initialize.md) イベントが発生すると、エンティティ アドインは現在のアイテムの [getEntities](../../reference/outlook/Office.context.mailbox.item.md) メソッドを呼び出します。**getEntities** メソッドは、グローバル変数 `_MyEntities` を返します。この変数は、サポートされているエンティティのインスタンスの配列です。関連する JavaScript コードを次に示します。


```js
// Global variables
var _Item;
var _MyEntities;

// The initialize function is required for all add-ins.
Office.initialize = function () {
    var _mailbox = Office.context.mailbox;
    // Obtains the current item.
    Item = _mailbox.item;
    // Reads all instances of supported entities from the subject 
    // and body of the current item.
    MyEntities = _Item.getEntities();
    
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    });
}

```


## 住所の抽出


ユーザーが **[Get Addresses]** ボタンをクリックすると、`myGetAddresses` イベント ハンドラーが `_MyEntities` オブジェクトの [addresses](../../reference/outlook/simple-types.md) プロパティから住所の配列を取得します (住所が抽出されていた場合)。抽出された各住所は、文字列として配列に格納されます。`myGetAddresses` はローカル HTML 文字列を .mdText` で生成し、抽出された住所の一覧を表示します。関連する JavaScript コードを次に示します。


```js
// Gets instances of the Address entity on the item.
function myGetAddresses()
{
    var htmlText = "";

    // Gets an array of postal addresses. Each address is a string.
    var addressesArray = _MyEntities.addresses;
    for (var i = 0; i < addressesArray.length; i++)
    {
        htmlText += "Address : <span>" + addressesArray[i] + "</span><br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}
```


## 連絡先情報の抽出


ユーザーが **[Get Contact Information]** ボタンをクリックすると、`myGetContacts` イベント ハンドラーが `_MyEntities` オブジェクトの [contacts](../../reference/outlook/simple-types.md) プロパティから連絡先の配列をそれぞれの情報と共に取得します (連絡先が抽出されていた場合)。抽出された各連絡先は、[Contact](../../reference/outlook/simple-types.md) オブジェクトとして配列に格納されます。`myGetContacts` は、各連絡先に関する詳細なデータを取得します。Outlook がアイテムから連絡先を抽出できるかどうかはコンテキスト次第であることに注意してください。電子メール メッセージの末尾の署名、または少なくとも次のいくつかの情報が連絡先の周辺に存在している必要があります。


- [Contact.personName](../../reference/outlook/simple-types.md) プロパティから取得される連絡先の名前を表す文字列。
    
- [Contact.businessName](../../reference/outlook/simple-types.md) プロパティから取得される連絡先に関連付けられた会社名を表す文字列。
    
- [Contact.phoneNumbers](../../reference/outlook/simple-types.md) プロパティから取得される、連絡先に関連付けられた電話番号の配列。各電話番号は [PhoneNumber](../../reference/outlook/simple-types.md) オブジェクトによって表されます。
    
- 電話番号配列内の **PhoneNumber** メンバーごとの、[PhoneNumber.phoneString](../../reference/outlook/simple-types.md) プロパティから取得される電話番号を表す文字列。
    
- [Contact.urls](../../reference/outlook/simple-types.md) プロパティから取得される連絡先に関連付けられた URL の配列。各 URL は配列メンバーの文字列として表されます。
    
- [Contact.emailAddresses](../../reference/outlook/simple-types.md) プロパティから取得される、連絡先に関連付けられた電子メール アドレスの配列。各電子メール アドレスは配列メンバーの文字列として表されます。
    
- [Contact.addresses](../../reference/outlook/simple-types.md) プロパティから取得される、連絡先に関連付けられた郵送先住所の配列。各郵送先住所は配列メンバーの文字列として表されます。
    
 `myGetContacts` はローカル HTML 文字列を `htmlText` で生成し、各連絡先のデータを表示します。関連する JavaScript コードを次に示します。




```js
// Gets instances of the Contact entity on the item.
function myGetContacts()
{
    var htmlText = "";

    // Gets an array of contacts and their information.
    var contactsArray = _MyEntities.contacts;
    for (var i = 0; i < contactsArray.length; i++)
    {
        // Gets the name of the person. The name is a string.
        htmlText += "Name : <span>" + contactsArray[i].personName +
            "</span><br/>";

        // Gets the company name associated with the contact.
        htmlText += "Business : <span>" + 
        contactsArray[i].businessName + "</span><br/>";

        // Gets an array of phone numbers associated with the 
        // contact. Each phone number is represented by a 
        // PhoneNumber object.
        var phoneNumbersArray = contactsArray[i].phoneNumbers;
        for (var j = 0; j < phoneNumbersArray.length; j++)
        {
            htmlText += "PhoneString : <span>" + 
                phoneNumbersArray[j].phoneString + "</span><br/>";
            htmlText += "OriginalPhoneString : <span>" + 
                phoneNumbersArray[j].originalPhoneString +
                "</span><br/>";
        }

        // Gets the URLs associated with the contact.
        var urlsArray = contactsArray[i].urls;
        for (var j = 0; j < urlsArray.length; j++)
        {
            htmlText += "Url : <span>" + urlsArray[j] + 
                "</span><br/>";
        }

        // Gets the email addresses of the contact.
        var emailAddressesArray = contactsArray[i].emailAddresses;
        for (var j = 0; j < emailAddressesArray.length; j++)
        {
           htmlText += "E-mail Address : <span>" + 
               emailAddressesArray[j] + "</span><br/>";
        }

        // Gets postal addresses of the contact.
        var addressesArray = contactsArray[i].addresses;
        for (var j = 0; j < addressesArray.length; j++)
        {
          htmlText += "Address : <span>" + addressesArray[j] + 
              "</span><br/>";
        }

        htmlText += "<hr/>";
        }

    document.getElementById("entities_box").innerHTML = htmlText;
}
```


## 電子メール アドレスの抽出


ユーザーが **[Get Email Addresses]** ボタンをクリックすると、`myGetEmailAddresses` イベント ハンドラーが `_MyEntities` オブジェクトの [emailAddresses](../../reference/outlook/simple-types.md) プロパティから SMTP 電子メール アドレスの配列を取得します (メール アドレスが抽出されていた場合)。抽出された各電子メール アドレスは、文字列として配列に格納されます。`myGetEmailAddresses` はローカル HTML 文字列を `htmlText` で生成し、抽出された電子メール アドレスの一覧を表示します。関連する JavaScript コードを次に示します。


```js
// Gets instances of the EmailAddress entity on the item.
function myGetEmailAddresses() {
    var htmlText = "";

    // Gets an array of email addresses. Each email address is a 
    // string.
    var emailAddressesArray = _MyEntities.emailAddresses;
    for (var i = 0; i < emailAddressesArray.length; i++) {
        htmlText += "E-mail Address : <span>" + emailAddressesArray[i] + "</span><br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}
```


## 会議提案の抽出


ユーザーが **[Get Meeting Suggestions]** ボタンをクリックすると、`myGetMeetingSuggestions` イベント ハンドラーが `_MyEntities` オブジェクトの [meetingSuggestions](../../reference/outlook/simple-types.md) プロパティから会議提案の配列を取得します (会議提案が抽出されていた場合)。


 >**注** **MeetingSuggestion** エンティティ型をサポートしているのはメッセージだけであり、予定ではサポートされていません。

抽出された各会議提案は、[MeetingSuggestion](../../reference/outlook/simple-types.md) オブジェクトとして配列に格納されます。`myGetMeetingSuggestions` は、各会議提案に関する次の詳細なデータを取得します。


- [MeetingSuggestion.meetingString](../../reference/outlook/simple-types.md) プロパティから取得される会議提案として識別された文字列。
    
- [MeetingSuggestion.attendees](../../reference/outlook/simple-types.md) プロパティから取得される、会議の出席者の配列。各出席者は [EmailUser](../../reference/outlook/simple-types.md) オブジェクトによって表されます。
    
- 出席者ごとの、[EmailUser.displayName](../../reference/outlook/simple-types.md) プロパティから取得される名前。
    
- 出席者ごとの、[EmailUser.emailAddress](../../reference/outlook/simple-types.md) プロパティから取得される SMTP アドレス。
    
- [MeetingSuggestion.location](../../reference/outlook/simple-types.md) プロパティから取得される、会議提案の場所を表す文字列。
    
- [MeetingSuggestion.subject](../../reference/outlook/simple-types.md) プロパティから取得される、会議提案の議題を表す文字列。
    
- [MeetingSuggestion.start](../../reference/outlook/simple-types.md) プロパティから取得される、会議提案の開始時刻を表す文字列。
    
- [MeetingSuggestion.end](../../reference/outlook/simple-types.md) プロパティから取得される会議提案の終了時刻を表す文字列。
    
 `myGetMeetingSuggestions` はローカル HTML 文字列を `htmlText` で生成し、会議提案ごとのデータを表示します。関連する JavaScript コードを次に示します。




```js
// Gets instances of the MeetingSuggestion entity on the 
// message item.
function myGetMeetingSuggestions() {
    var htmlText = "";

    // Gets an array of MeetingSuggestion objects, each array 
    // element containing an instance of a meeting suggestion 
    // entity from the current item.
    var meetingsArray = _MyEntities.meetingSuggestions;

    // Iterates through each instance of a meeting suggestion.
    for (var i = 0; i < meetingsArray.length; i++) {
        // Gets the string that was identified as a meeting suggestion.
        htmlText += "MeetingString : <span>" + meetingsArray[i].meetingString + "</span><br/>";

        // Gets an array of attendees for that instance of a 
        // meeting suggestion. Each attendee is represented 
        // by an EmailUser object.
        var attendeesArray = meetingsArray[i].attendees;
        for (var j = 0; j < attendeesArray.length; j++) {
            htmlText += "Attendee : ( ";

            // Gets the displayName property of the attendee.
            htmlText += "displayName = <span>" + attendeesArray[j].displayName + "</span> , ";

            // Gets the emailAddress property of each attendee.
            // This is the SMTP address of the attendee.
            htmlText += "emailAddress = <span>" + attendeesArray[j].emailAddress + "</span>";

            htmlText += " )<br/>";
        }

        // Gets the location of the meeting suggestion.
        htmlText += "Location : <span>" + meetingsArray[i].location + "</span><br/>";

        // Gets the subject of the meeting suggestion.
        htmlText += "Subject : <span>" + meetingsArray[i].subject + "</span><br/>";

        // Gets the start time of the meeting suggestion.
        htmlText += "Start time : <span>" + meetingsArray[i].start + "</span><br/>";

        // Gets the end time of the meeting suggestion.
        htmlText += "End time : <span>" + meetingsArray[i].end + "</span><br/>";

        htmlText += "<hr/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}
```


## 電話番号の抽出


ユーザーが **[Get Phone Numbers]** ボタンをクリックすると、`myGetPhoneNumbers` イベント ハンドラーが `_MyEntities` オブジェクトの [phoneNumbers](../../reference/outlook/simple-types.md) プロパティから電話番号の配列を取得します (電話番号が抽出されていた場合)。抽出された各電話番号は、[PhoneNumber](../../reference/outlook/simple-types.md) オブジェクトとして配列に格納されます。`myGetPhoneNumbers` は、各電話番号に関する次の詳細なデータを取得します。


- [PhoneNumber.type](../../reference/outlook/simple-types.md) プロパティから取得される電話番号の種類 (自宅の電話番号など) を表す文字列。
    
- [PhoneNumber.phoneString](../../reference/outlook/simple-types.md) プロパティから取得される、実際の電話番号を表す文字列。
    
- [PhoneNumber.originalPhoneString](../../reference/outlook/simple-types.md) プロパティから取得される電話番号として最初に識別された文字列。
    
 `myGetPhoneNumbers` はローカル HTML 文字列を `htmlText` で生成し、各電話番号のデータを表示します。関連する JavaScript コードを次に示します。




```js
// Gets instances of the phone number entity on the item.
function myGetPhoneNumbers()
{
    var htmlText = "";

    // Gets an array of phone numbers. 
    // Each phone number is a PhoneNumber object.
    var phoneNumbersArray = _MyEntities.phoneNumbers;
    for (var i = 0; i < phoneNumbersArray.length; i++)
    {
        htmlText += "Phone Number : ( ";
        // Gets the type of phone number, for example, home, office.
        htmlText += "type = <span>" + phoneNumbersArray[i].type + 
           "</span> , ";

        // Gets the actual phone number represented by a string.
        htmlText += "phone string = <span>" + 
            phoneNumbersArray[i].phoneString + "</span> , ";

        // Gets the original text that was identified in the item 
        // as a phone number. 
        htmlText += "original phone string = <span>" + 
            phoneNumbersArray[i].originalPhoneString + "</span>";

        htmlText += " )<br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}

```


## タスクの提案の抽出


ユーザーが **[Get Task Suggestions]** ボタンをクリックすると、`myGetTaskSuggestions` イベント ハンドラーが `_MyEntities` オブジェクトの [taskSuggestions](../../reference/outlook/simple-types.md) プロパティからタスクの提案の配列を取得します (タスクの提案が抽出されていた場合)。抽出された各タスクの提案は、[TaskSuggestion](../../reference/outlook/simple-types.md) オブジェクトとして配列に格納されます。`myGetTaskSuggestions` は、各タスクの提案に関する次の詳細なデータを取得します。


- [TaskSuggestion.taskString](../../reference/outlook/simple-types.md) プロパティから取得されるタスクの提案として最初に識別された文字列。
    
- [TaskSuggestion.assignees](../../reference/outlook/simple-types.md) プロパティから取得される、タスクの割り当て先の配列。各割り当て先は [EmailUser](../../reference/outlook/simple-types.md) オブジェクトによって表されます。
    
- 割り当て先ごとの、[EmailUser.displayName](../../reference/outlook/simple-types.md) プロパティから取得される名前。
    
- 割り当て先ごとの、[EmailUser.emailAddress](../../reference/outlook/simple-types.md) プロパティから取得される SMTP アドレス。
    
 `myGetTaskSuggestions` はローカル HTML 文字列を `htmlText` で生成し、タスクの提案ごとのデータを表示します。関連する JavaScript コードを次に示します。




```js
// Gets instances of the task suggestion entity on the item.
function myGetTaskSuggestions()
{
    var htmlText = "";

    // Gets an array of TaskSuggestion objects, each array element 
    // containing an instance of a task suggestion entity from 
    // the current item.
    var tasksArray = _MyEntities.taskSuggestions;

    // Iterates through each instance of a task suggestion.
    for (var i = 0; i < tasksArray.length; i++)
    {
        // Gets the string that was identified as a task suggestion.
        htmlText += "TaskString : <span>" + 
           tasksArray[i].taskString + "</span><br/>";

        // Gets an array of assignees for that instance of a task 
        // suggestion. Each assignee is represented by an 
        // EmailUser object.
        var assigneesArray = tasksArray[i].assignees;
        for (var j = 0; j < assigneesArray.length; j++)
        {
            htmlText += "Assignee : ( ";
            // Gets the displayName property of the assignee.
            htmlText += "displayName = <span>" + assigneesArray[j].displayName + 
               "</span> , ";

            // Gets the emailAddress property of each assignee.
            // This is the SMTP address of the assignee.
            htmlText += "emailAddress = <span>" + assigneesArray[j].emailAddress + 
                "</span>";

            htmlText += " )<br/>";
        }

        htmlText += "<hr/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}

```


## URL の抽出


ユーザーが **[Get URLs]** ボタンをクリックすると、`myGetUrls` イベント ハンドラーが `_MyEntities` オブジェクトの [urls](../../reference/outlook/simple-types.md) プロパティから URL の配列を取得します (URL が抽出されていた場合)。抽出された各 URL は、文字列として配列に格納されます。`myGetUrls` はローカル HTML 文字列を `htmlText` で生成し、抽出された URL の一覧を表示します。


```js
// Gets instances of the URL entity on the item.
function myGetUrls()
{
    var htmlText = "";

    // Gets an array of URLs. Each URL is a string.
    var urlArray = _MyEntities.urls;
    for (var i = 0; i < urlArray.length; i++)
    {
        htmlText += "Url : <span>" + urlArray[i] + "</span><br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}

```


## 表示されたエンティティ文字列の消去


最後に、エンティティ アドインでは表示された文字列を消去する `myClearEntitiesBox` イベント ハンドラーを指定しています。関連するコードを次に示します。


```js
// Clears the div with id="entities_box".
function myClearEntitiesBox()
{
    document.getElementById("entities_box").innerHTML = "";
}
```


## JavaScript の内容


次に、JavaScript の実装の内容全体を示します。


```js
// Global variables
var _Item;
var _MyEntities;

// Initializes the add-in.
Office.initialize = function () {
    var _mailbox = Office.context.mailbox;
    // Obtains the current item.
    _Item = _mailbox.item;
    // Reads all instances of supported entities from the subject 
    // and body of the current item.
    _MyEntities = _Item.getEntities();

    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    });
}


// Clears the div with id="entities_box".
function myClearEntitiesBox()
{
    document.getElementById("entities_box").innerHTML = "";
}

// Gets instances of the Address entity on the item.
function myGetAddresses()
{
    var htmlText = "";

    // Gets an array of postal addresses. Each address is a string.
    var addressesArray = _MyEntities.addresses;
    for (var i = 0; i < addressesArray.length; i++)
    {
        htmlText += "Address : <span>" + addressesArray[i] + 
            "</span><br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}


// Gets instances of the EmailAddress entity on the item.
function myGetEmailAddresses()
{
    var htmlText = "";

    // Gets an array of email addresses. Each email address is a 
    // string.
    var emailAddressesArray = _MyEntities.emailAddresses;
    for (var i = 0; i < emailAddressesArray.length; i++)
    {
        htmlText += "E-mail Address : <span>" + 
            emailAddressesArray[i] + "</span><br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}

// Gets instances of the MeetingSuggestion entity on the 
// message item.
function myGetMeetingSuggestions()
{
    var htmlText = "";

    // Gets an array of MeetingSuggestion objects, each array 
    // element containing an instance of a meeting suggestion 
    // entity from the current item.
    var meetingsArray = _MyEntities.meetingSuggestions;

    // Iterates through each instance of a meeting suggestion.
    for (var i = 0; i < meetingsArray.length; i++)
    {
        // Gets the string that was identified as a meeting 
        // suggestion.
        htmlText += "MeetingString : <span>" + 
            meetingsArray[i].meetingString + "</span><br/>";

        // Gets an array of attendees for that instance of a 
        // meeting suggestion.
        // Each attendee is represented by an EmailUser object.
        var attendeesArray = meetingsArray[i].attendees;
        for (var j = 0; j < attendeesArray.length; j++)
        {
            htmlText += "Attendee : ( ";
            // Gets the displayName property of the attendee.
            htmlText += "displayName = <span>" + attendeesArray[j].displayName + 
                "</span> , ";

            // Gets the emailAddress property of each attendee.
            // This is the SMTP address of the attendee.
            htmlText += "emailAddress = <span>" + attendeesArray[j].emailAddress + 
                "</span>";

            htmlText += " )<br/>";
        }

        // Gets the location of the meeting suggestion.
        htmlText += "Location : <span>" + 
            meetingsArray[i].location + "</span><br/>";

        // Gets the subject of the meeting suggestion.
        htmlText += "Subject : <span>" + 
            meetingsArray[i].subject + "</span><br/>";

        // Gets the start time of the meeting suggestion.
        htmlText += "Start time : <span>" + 
           meetingsArray[i].start + "</span><br/>";

        // Gets the end time of the meeting suggestion.
        htmlText += "End time : <span>" + 
            meetingsArray[i].end + "</span><br/>";

        htmlText += "<hr/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}


// Gets instances of the phone number entity on the item.
function myGetPhoneNumbers()
{
    var htmlText = "";

    // Gets an array of phone numbers. 
    // Each phone number is a PhoneNumber object.
    var phoneNumbersArray = _MyEntities.phoneNumbers;
    for (var i = 0; i < phoneNumbersArray.length; i++)
    {
        htmlText += "Phone Number : ( ";
        // Gets the type of phone number, for example, home, office.
        htmlText += "type = <span>" + phoneNumbersArray[i].type + 
            "</span> , ";

        // Gets the actual phone number represented by a string.
        htmlText += "phone string = <span>" + 
            phoneNumbersArray[i].phoneString + "</span> , ";

        // Gets the original text that was identified in the item 
        // as a phone number. 
        htmlText += "original phone string = <span>" + 
           phoneNumbersArray[i].originalPhoneString + "</span>";

        htmlText += " )<br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}

// Gets instances of the task suggestion entity on the item.
function myGetTaskSuggestions()
{
    var htmlText = "";

    // Gets an array of TaskSuggestion objects, each array element 
    // containing an instance of a task suggestion entity from the 
    // current item.
    var tasksArray = _MyEntities.taskSuggestions;

    // Iterates through each instance of a task suggestion.
    for (var i = 0; i < tasksArray.length; i++)
    {
        // Gets the string that was identified as a task suggestion.
        htmlText += "TaskString : <span>" + 
            tasksArray[i].taskString + "</span><br/>";

        // Gets an array of assignees for that instance of a task 
        // suggestion. Each assignee is represented by an 
        // EmailUser object.
        var assigneesArray = tasksArray[i].assignees;
        for (var j = 0; j < assigneesArray.length; j++)
        {
            htmlText += "Assignee : ( ";
            // Gets the displayName property of the assignee.
            htmlText += "displayName = <span>" + assigneesArray[j].displayName + 
                "</span> , ";

            // Gets the emailAddress property of each assignee.
            // This is the SMTP address of the assignee.
            htmlText += "emailAddress = <span>" + assigneesArray[j].emailAddress + 
                "</span>";

            htmlText += " )<br/>";
        }

        htmlText += "<hr/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}

// Gets instances of the URL entity on the item.
function myGetUrls()
{
    var htmlText = "";

    // Gets an array of URLs. Each URL is a string.
    var urlArray = _MyEntities.urls;
    for (var i = 0; i < urlArray.length; i++)
    {
        htmlText += "Url : <span>" + urlArray[i] + "</span><br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}

```


## その他のリソース



- [閲覧フォーム用の Outlook アドインを作成する](../outlook/read-scenario.md)
    
- [Outlook アイテム内の文字列を既知のエンティティとして照合する](../outlook/match-strings-in-an-item-as-well-known-entities.md)
    
- [item.getEntities メソッド](../../reference/outlook/Office.context.mailbox.item.md)
    
