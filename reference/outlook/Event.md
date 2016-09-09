

# イベント

`event` オブジェクトは、UI を使用しないコマンド ボタンによって呼び出されるアドイン関数のパラメーターとして渡されます。オブジェクトにより、アドインはどのボタンがクリックされたかを識別し、その処理を行ったホストにシグナルを送ることができます。

たとえば、アドインのマニフェストで次のように定義されているボタンがあるとします。

```
<Control xsi:type="Button" id="eventTestButton">
  <Label resid="eventButtonLabel" />
  <Tooltip resid="eventButtonTooltip" />
  <Supertip>
    <Title resid="eventSuperTipTitle" />
    <Description resid="eventSuperTipDescription" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="blue-icon-16" />
    <bt:Image size="32" resid="blue-icon-32" />
    <bt:Image size="80" resid="blue-icon-80" />
  </Icon>
  <Action xsi:type="ExecuteFunction">
    <FunctionName>testEventObject</FunctionName>
  </Action>
</Control>
```

ボタンは、`id` 属性が `eventTestButton` に設定されており、アドインで定義された `testEventObject` 関数を呼び出します。その関数は、次のようになります。

```
function testEventObject(event) {
  // The event object implements the Event interface

  // This value will be "eventTestButton"
  var buttonId = event.source.id;

  // Signal to the host app that processing is complete.
  event.completed();
}
```

##### 要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](./tutorial-api-requirement-sets.md)| 1.3|
|[最小のアクセス許可レベル](../../docs/outlook/understanding-outlook-add-in-permissions.md)| Restricted|
|適用可能な Outlook のモード| 作成または読み取り|

### メンバー

####  ソース: オブジェクト

メソッドを呼び出したアドイン コマンド ボタンの識別子を取得します。

`source` プロパティは、次のプロパティを持つオブジェクトを返します。

| プロパティ | 説明 |
| --- | --- |
| `id` | アドイン マニフェストのアドイン コマンド ボタンを定義する、`id` 要素の `Control` 属性の値です。 |

この値は、1 つ以上のボタンが同じ関数を呼び出すものの、どのボタンがクリックされたかによって異なる操作を実行しなければならない場合に使用できます。

##### 型:

*   オブジェクト

##### 要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](./tutorial-api-requirement-sets.md)| 1.3|
|[最小のアクセス許可レベル](../../docs/outlook/understanding-outlook-add-in-permissions.md)| Restricted|
|適用可能な Outlook のモード| 作成または読み取り|

##### 例

```
// Function is used by two buttons:
// button1 and button2
function multiButton (event) {
  // Check which button was clicked
  var buttonId = event.source.id;

  if (buttonId === 'button1') {
    doButton1Action();
  else {
    doButton2Action();
  }

  event.completed();
}
```

### メソッド

####  completed()

アドインが、アドイン コマンド ボタンによりトリガーされた処理を完了したことを示します。

このメソッドは、`Action` 属性が `xsi:type` に設定された `ExecuteFunction` 要素で定義されたアドイン コマンドにより呼び出された関数の最後に呼び出される必要があります。このメソッドを呼び出すと、関数が終了したことと、関数の呼び出しに関連するすべての状態をクリーンアップできることがホスト クライアントに通知されます。たとえば、ユーザーがこのメソッドを呼び出す前に Outlook を終了すると、関数が実行中であることが Outlook により警告されます。

##### 要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](./tutorial-api-requirement-sets.md)| 1.3|
|[最小のアクセス許可レベル](../../docs/outlook/understanding-outlook-add-in-permissions.md)| Restricted|
|適用可能な Outlook のモード| 作成または読み取り|

##### 例

```
function processItem (event) {
  // Do some processing

  event.completed();
}
```