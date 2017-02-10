# <a name="add-support-for-add-in-commands-for-outlook-mobile"></a>Outlook Mobile のアドイン コマンドのサポートを追加する

> **注**:Outlook Mobile のアドイン コマンドは、現在 Outlook for iOS でのみサポートされています。

Outlook Mobile のアドイン コマンドを使用すると、ユーザーは、Outlook for Windows、Outlook for Mac、および Outlook on the web.にあるものと同じ機能にアクセスできます (ただし、いくつかの[制限](#code-considerations)があります)。Outlook Mobile のサポートを追加するには、アドイン マニフェストを更新する必要があります。おそらく、モバイル シナリオのコードを変更する必要もあります。

## <a name="updating-the-manifest"></a>マニフェストを更新する

Outlook Mobile でアドイン コマンドを有効にするための最初の手順は、アドイン マニフェストでの定義です。**VersionOverrides** v1.1 スキーマは、モバイル用に新しいフォーム ファクター [MobileFormFactor](../../reference/manifest/mobileformfactor.md) を定義します。

この要素には、モバイル クライアントにアドインを読み込むためのすべての情報が含まれています。これにより、モバイル エクスペリエンスに対して完全に異なる UI 要素と JavaScript ファイルを定義することができます。

次の例は、**MobileFormFactor** 要素の単一の作業ウィンドウ ボタンを示しています。

```xml
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
  ...
  <MobileFormFactor>
    <FunctionFile resid="residUILessFunctionFileUrl" />
    <ExtensionPoint xsi:type="MobileMessageReadCommandSurface">
      <Control xsi:type="MobileButton" id="TaskPane1Btn">
        <Label resid="residTaskPaneButton0Name" />
        <Icon xsi:type="bt:MobileIconList">
          <bt:Image size="25" scale="1" resid="tp0icon" />
          <bt:Image size="25" scale="2" resid="tp0icon" />
          <bt:Image size="25" scale="3" resid="tp0icon" />

          <bt:Image size="32" scale="1" resid="tp0icon" />
          <bt:Image size="32" scale="2" resid="tp0icon" />
          <bt:Image size="32" scale="3" resid="tp0icon" />

          <bt:Image size="48" scale="1" resid="tp0icon" />
          <bt:Image size="48" scale="2" resid="tp0icon" />
          <bt:Image size="48" scale="3" resid="tp0icon" />
        </Icon>
        <Action xsi:type="ShowTaskpane">
          <SourceLocation resid="residTaskpaneUrl" />
        </Action>
      </Control>
    </ExtensionPoint>
  </MobileFormFactor>
  ...
</VersionOverrides>
```

これは、[DesktopFormFactor](../../reference/manifest/desktopformfactor.md) 要素に表示される要素と非常によく似ていますが、いくつかの注目すべき違いがあります。

- [OfficeTab](../../reference/manifest/officetab.md) 要素は使用されません。
- [ExtensionPoint](../../reference/manifest/exensionpoint.md) 要素に含まれる子要素は 1 つでなければなりません。アドインがボタンを 1 つのみ追加する場合、子要素は [Control](../../reference/manifest/control.md) 要素になります。アドインがボタンを複数追加する場合、子要素は複数の `Control` 要素を含む [Group](../../reference/manifest/group.md) 要素になります。
- `Control` 要素に相当する `Menu` の種類はありません。
- [Supertip](../../reference/manifest/supertip.md) 要素は使用されません。
- アイコンの必須サイズが異なります。モバイル アドインは少なくとも 25x25、32x32 および 48x48 ピクセルのアイコンをサポートする必要があります。

## <a name="code-considerations"></a>コードに関する考慮事項

モバイル用のアドインの設計には、追加の考慮事項がいくつか導入されています。

### <a name="use-rest-instead-of-exchange-web-services"></a>Exchange Web サービスの代わりに REST を使用する

[Office.context.mailbox.makeEwsRequestAsync](../../reference/outlook/Office.context.mailbox.md) メソッドは、Outlook Mobile ではサポートされていません。可能な場合には、アドインは優先的に Office.js API から情報を取得します。Office.js API によって表示されていない情報がアドインで必要な場合、[Outlook REST APIs](https://dev.outlook.com/restapi/reference) を使用してユーザーのメールボックスにアクセスする必要があります。 

メールボックスの要件セット 1.5 では、REST API と互換性のあるアクセス トークンを要求できる [Office.context.mailbox.getCallbackTokenAsync](https://dev.outlook.com/reference/add-ins/1.5/Office.context.mailbox.html#getCallbackTokenAsync) の新しいバージョンと、ユーザーの REST API エンドポイントを検索するために使用できる新しい [Office.context.mailbox.restUrl](https://dev.outlook.com/reference/add-ins/1.5/Office.context.mailbox.html#restUrl) プロパティが導入されています。

### <a name="pinch-zoom"></a>ピンチによるズーム

既定で、ユーザーは "ピンチによるズーム" ジェスチャを使用して作業ウィンドウで拡大することができます。ご使用のシナリオでこれが該当しない場合は、HTML でピンチによるズームを無効にしてください。

### <a name="closing-taskpanes"></a>作業ウィンドウを閉じる

Outlook Mobile では、作業ウィンドウが画面全体を占めるので、既定ではユーザーが作業ウィンドウを閉じてメッセージに戻る必要があります。シナリオが完成したら、[Office.context.ui.closeContainer](https://dev.outlook.com/reference/add-ins/1.5/Office.context.ui.html#closeContainer) メソッドを使用して作業ウィンドウを閉じることを検討してください。

### <a name="compose-mode-and-appointments"></a>作成モードと予定

現在、Outlook Mobile のアドインは、メッセージ読み取り時のアクティブ化のみをサポートしています。メッセージを作成するときや、予定を表示または作成するときには、アドインはアクティブ化されません。

### <a name="unsupported-apis"></a>サポートされていない API

Outlook Mobile では、次の API はサポートされていません。

  - [Office.context.officeTheme](../../reference/outlook/Office.context.md)
  - [Office.context.mailbox.ewsUrl](../../reference/outlook/Office.context.mailbox.md)
  - [Office.context.mailbox.convertToEwsId](../../reference/outlook/Office.context.mailbox.md)
  - [Office.context.mailbox.convertToRestId](../../reference/outlook/Office.context.mailbox.md)
  - [Office.context.mailbox.displayAppointmentForm](../../reference/outlook/Office.context.mailbox.md)
  - [Office.context.mailbox.displayMessageForm](../../reference/outlook/Office.context.mailbox.md)
  - [Office.context.mailbox.displayNewAppointmentForm](../../reference/outlook/Office.context.mailbox.md)
  - [Office.context.mailbox.makeEwsRequestAsync](../../reference/outlook/Office.context.mailbox.md)
  - [Office.context.mailbox.item.dateTimeModified](../../reference/outlook/Office.context.mailbox.item.md)
  - [Office.context.mailbox.item.resources](../../reference/outlook/Office.context.mailbox.item.md)
  - [Office.context.mailbox.item.displayReplyAllForm](../../reference/outlook/Office.context.mailbox.item.md)
  - [Office.context.mailbox.item.displayReplyForm](../../reference/outlook/Office.context.mailbox.item.md)
  - [Office.context.mailbox.item.getEntities](../../reference/outlook/Office.context.mailbox.item.md)
  - [Office.context.mailbox.item.getEntitiesByType](../../reference/outlook/Office.context.mailbox.item.md)
  - [Office.context.mailbox.item.getFilteredEntitiesByName](../../reference/outlook/Office.context.mailbox.item.md)
  - [Office.context.mailbox.item.getRegexMatches](../../reference/outlook/Office.context.mailbox.item.md)
  - [Office.context.mailbox.item.getRegexMatchesByName](../../reference/outlook/Office.context.mailbox.item.md)