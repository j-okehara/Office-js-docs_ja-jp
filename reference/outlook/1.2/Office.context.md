

# <a name="context"></a>context

## [Office](Office.md). context

Office.context 名前空間は、すべての Office アプリのアドインで使う共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office.context 名前空間の完全な一覧は、「[共有 API の Office.context リファレンス](../../shared/office.context.md)」をご覧ください。


##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](../tutorial-api-requirement-sets.md)| 1.0|
|適用可能な Outlook のモード| 作成または読み取り|

### <a name="namespaces"></a>名前空間

[メールボックス](Office.context.mailbox.md) - Microsoft Outlook と Web 上の Microsoft Outlook の Outlook アドイン オブジェクト モデルへのアクセスを提供します。

### <a name="members"></a>メンバー

####  <a name="displaylanguage-:string"></a>displayLanguage :String

Office ホスト アプリケーションの UI 用にユーザーが指定した RFC 1766 言語タグ形式のロケール (言語) を取得します。

`displayLanguage` の値は、Office ホスト アプリケーションの **[ファイル]、[選択肢]、[言語]** によって指定される現在の **[表示言語]** 設定を反映します。

##### <a name="type:"></a>型:

*   String

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](../tutorial-api-requirement-sets.md)| 1.0|
|適用可能な Outlook のモード| 作成または読み取り|

##### <a name="example"></a>例

```js
function sayHelloWithDisplayLanguage() {
  var myDisplayLanguage = Office.context.displayLanguage;
  switch (myDisplayLanguage) {
    case 'en-US':
      write('Hello!');
      break;
    case 'en-NZ':
      write('G\'day mate!');
      break;
  }
}
// Function that writes to a div with id='message' on the page.
function write(message){
  document.getElementById('message').innerText += message;
}
```

####  <a name="roamingsettings-:[roamingsettings](roamingsettings.md)"></a>roamingSettings :[RoamingSettings](RoamingSettings.md)

ユーザーのメールボックスに保存されている、メール アドインのカスタム設定や状態を表すオブジェクトを取得します。

`RoamingSettings` オブジェクトを使うと、ユーザーのメールボックスに保存されている、メール アドインのデータの保存やアクセスを実行できます。そのため、メール アドインは、このメールボックスへのアクセスに使うどのホスト クライアント アプリケーションから実行されても、このデータを使うことができます。

##### <a name="type:"></a>型:

*   [RoamingSettings](RoamingSettings.md)

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](../tutorial-api-requirement-sets.md)| 1.0|
|[最小限のアクセス許可レベル](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| 制限あり|
|適用可能な Outlook のモード| 作成または読み取り|
