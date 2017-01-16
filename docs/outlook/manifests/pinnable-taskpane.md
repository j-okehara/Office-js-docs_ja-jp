# <a name="implement-a-pinnable-taskpane-in-outlook"></a>Outlook にピン留め可能な作業ウィンドウを実装する

アドイン コマンド用の[作業ウィンドウ](../add-in-commands-for-outlook.md#launching-a-task-pane) UX シェイプは、開いたメッセージまたは予定の右側に縦方向の作業ウィンドウを開きます。アドインは、このウィンドウを使用することで、より詳細な対話式操作 (複数フィールドの入力など) に対応した UI を提供できようになります。この作業ウィンドウは、メッセージの一覧を表示しているときに、閲覧ウィンドウに表示できます。これにより、メッセージの素早い処理が可能になります。

ただし、既定では、ユーザーが新しいメッセージを選択すると、閲覧ウィンドウ内で開いていたメッセージのアドイン作業ウィンドウは自動的に閉じられます。頻繁に使用されるアドインの場合、ユーザーはそのウィンドウを開いたままにして、メッセージごとにアドインを有効化する手間がなくなることを望むでしょう。ピン留め可能な作業ウィンドウでは、これに該当するオプションをユーザーに提供できます。

> **注**: 現時点では、ピン留め可能な作業ウィンドウは Outlook 2016 (バージョン 7628.1000) でのみサポートされます。

## <a name="support-taskpane-pinning"></a>作業ウィンドウのピン留めをサポートする

ピン留めのサポートを追加する際の最初の手順は、アドインの[マニフェスト](./manifests.md)で実行します。この手順は、作業ウィンドウのボタンについて記述する [SupportsPinning](../../../reference/manifest/action.md#supportspinning) 要素を `Action` 要素に追加することで実行します。

`SupportsPinning` 要素は、VersionOverrides v1.1 スキーマで定義されているため、v1.0 と v1.1 のどちらの場合も [VersionOverrides](../../../reference/manifest/versionoverrides.md) 要素を含める必要があります。

```xml
<!-- Task pane button -->
<Control xsi:type="Button" id="msgReadOpenPaneButton">
  <Label resid="paneReadButtonLabel" />
  <Supertip>
    <Title resid="paneReadSuperTipTitle" />
    <Description resid="paneReadSuperTipDescription" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="green-icon-16" />
    <bt:Image size="32" resid="green-icon-32" />
    <bt:Image size="80" resid="green-icon-80" />
  </Icon>
  <Action xsi:type="ShowTaskpane">
    <SourceLocation resid="readTaskPaneUrl" />
    <SupportsPinning>true</SupportsPinning>
  </Action>
</Control>
```

完全な例については、[command-demo のサンプル マニフェスト](https://github.com/jasonjoh/command-demo/blob/master/command-demo-manifest.xml)の `msgReadOpenPaneButton` コントロールを参照してください。

## <a name="handling-ui-updates-based-on-currently-selected-message"></a>現在選択されているメッセージに基づいた UI の更新を処理する

現在のアイテムに基づいて作業ウィンドウの UI または内部変数を更新するには、変更の通知を受け取るイベント ハンドラの登録が必要になります。

### <a name="implement-the-event-handler"></a>イベント ハンドラを実装する

イベント ハンドラは、オブジェクト リテラルの単一パラメーターを受け入れる必要があります。このオブジェクトの `type` プロパティは、`Office.EventType.ItemChanged` に設定されます。イベントが呼び出されたときには、既に、`Office.context.mailbox.item` オブジェクトは現在選択されているアイテムを反映するように更新されています。

```js
function itemChanged(eventArgs) {
  // Update UI based on the new current item
  UpdateTaskPaneUI(Office.context.mailbox.item);
}
```

### <a name="register-the-event-handler"></a>イベント ハンドラを登録する

[Office.context.mailbox.addHandlerAsync](https://dev.outlook.com/reference/add-ins/1.5/Office.context.mailbox.html#addHandlerAsync) メソッドを使用して、`Office.EventType.ItemChanged` イベント用のイベント ハンドラを登録します。これは、作業ウィンドウの `Office.initialize` 関数で実行する必要があります。

```js
Office.initialize = function (reason) {
  $(document).ready(function () {

    // Set up ItemChanged event
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, itemChanged);

    UpdateTaskPaneUI(Office.context.mailbox.item);
  });
};
```

## <a name="additional-resources"></a>その他のリソース

ピン留め可能な作業ウィンドウを実装するサンプル アドインについては、GitHub の [command-demo](https://github.com/jasonjoh/command-demo) を参照してください。