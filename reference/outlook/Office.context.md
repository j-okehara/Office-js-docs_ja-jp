

# <a name="context"></a>context

## <a name="officeofficemdcontext"></a>[Office](Office.md).context

Office.context 名前空間は、すべての Office アプリのアドインで使う共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office.context 名前空間の完全な一覧は、「[共有 API の Office.context リファレンス](../shared/office.context.md)」をご覧ください。

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](./tutorial-api-requirement-sets.md)| 1.0|
|適用可能な Outlook のモード| 作成または読み取り|

### <a name="namespaces"></a>名前空間

[mailbox](Office.context.mailbox.md):Microsoft Outlook と Microsoft Outlook on the web の Outlook アドイン オブジェクト モデルへのアクセスを提供します。

### <a name="members"></a>メンバー

####  <a name="displaylanguage-string"></a>displayLanguage :String

Office ホスト アプリケーションの UI 用にユーザーが指定した RFC 1766 言語タグ形式のロケール (言語) を取得します。

`displayLanguage` の値は、Office ホスト アプリケーションの **[ファイル]、[選択肢]、[言語]** によって指定される現在の **[表示言語]** 設定を反映します。

##### <a name="type"></a>型:

*   String

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](./tutorial-api-requirement-sets.md)| 1.0|
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

####  <a name="officetheme-object"></a>officeTheme :Object

Office テーマの色のプロパティにアクセスできるようにします。

> **注:**このメンバーは、Outlook for iOS または Outlook for Android ではサポートされていません。

Office テーマの色を使うと、**[ファイル] > [Office アカウント] > [Office テーマ UI]** によってユーザーが選択した現在の Office テーマに合わせてアドインの配色を調整できます。このテーマは Office ホスト アプリケーション全体に適用されます。Office テーマの色を使うことは、メール アドインと作業ウィンドウ アドインに適しています。

##### <a name="type"></a>型:

*   Object

##### <a name="properties"></a>プロパティ:

|名前| 型| 説明|
|---|---|---|
|`bodyBackgroundColor`| String|Office テーマの本文の背景色を 16 進数の組み合わせとして取得します。|
|`bodyForegroundColor`| String|Office テーマの本文の前景色を 16 進数の組み合わせとして取得します。|
|`controlBackgroundColor`| String|Office テーマのコントロールの背景色を 16 進数の組み合わせとして取得します。|
|`controlForegroundColor`| String|Office テーマの本文のコントロール色を 16 進数の組み合わせとして取得します。|

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](./tutorial-api-requirement-sets.md)| 1.3|
|適用可能な Outlook のモード| 作成または読み取り|

##### <a name="example"></a>例

```js
function applyOfficeTheme(){
  // Get office theme colors.
  var bodyBackgroundColor = Office.context.officeTheme.bodyBackgroundColor;
  var bodyForegroundColor = Office.context.officeTheme.bodyForegroundColor;
  var controlBackgroundColor = Office.context.officeTheme.controlBackgroundColor
  var controlForegroundColor = Office.context.officeTheme.controlForegroundColor;

  // Apply body background color to a CSS class.
  $('.body').css('background-color', bodyBackgroundColor);
}
```

####  <a name="roamingsettings-roamingsettingsroamingsettingsmd"></a>roamingSettings :[RoamingSettings](RoamingSettings.md)

ユーザーのメールボックスに保存されている、メール アドインのカスタム設定や状態を表すオブジェクトを取得します。

`RoamingSettings` オブジェクトを使うと、ユーザーのメールボックスに保存されている、メール アドインのデータの保存やアクセスを実行できます。そのため、メール アドインは、このメールボックスへのアクセスに使うどのホスト クライアント アプリケーションから実行されても、このデータを使うことができます。

##### <a name="type"></a>型:

*   [RoamingSettings](RoamingSettings.md)

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](./tutorial-api-requirement-sets.md)| 1.0|
|[最小限のアクセス許可レベル](../../docs/outlook/understanding-outlook-add-in-permissions.md)| 制限あり|
|適用可能な Outlook のモード| 作成または読み取り|
