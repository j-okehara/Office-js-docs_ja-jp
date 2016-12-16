 

# <a name="understanding-outlook-api-requirement-sets"></a>Outlook API 要件セットについて

Outlook アドインは、[マニフェスト](https://msdn.microsoft.com/EN-US/library/office/dn592036.aspx)で[要件](https://msdn.microsoft.com/en-us/library/office/fp123693.aspx)要素を使用して、必要な API のバージョンを宣言します。Outlook アドインには、`Name` 属性が `Mailbox` に設定され、`MinVersion` 属性がアドインのシナリオをサポートする最小 API 要件セットに設定された[設定](https://msdn.microsoft.com/EN-US/library/office/dn592049.aspx)要素が常に含まれます。

たとえば、次のマニフェストのスニペットは、最小要件セットの 1.1 を表します。

```xml
<Requirements>
  <Sets>
    <Set Name="MailBox" MinVersion="1.1" />
  </Sets>
</Requirements>
```

すべての Outlook API は、`Mailbox`[要件セット](https://msdn.microsoft.com/EN-US/library/office/dn535871.aspx#SpecifyRequirementSets_intro)に属しています。`Mailbox` 要件セットには複数のバージョンがあり、リリースされる API の新しいセットはそれぞれのセットの上位バージョンに属します。すべての Outlook クライアントが最新の API のセットをサポートするわけではありませんが、Outlook クライアントが要件セットのサポートを宣言する場合は、その要件セットのすべての API がサポートされます。

マニフェストに要件セットの最小バージョンを設定することで、アドインが表示される Outlook クライアントをコントロールできます。クライアントが最小要件セットをサポートしない場合、アドインはロードされません。たとえば、要件セットのバージョン 1.3 が指定されている場合、1.3 以上をサポートしていない Outlook クライアントには表示されません。

## <a name="using-apis-from-later-requirement-sets"></a>後続の要件セットからの API の使用

要件セットを設定しても、アドインを使用できる API は制限されません。たとえば、アドインでは要件セット 1.1 が指定されていて、1.3 をサポートしている Outlook クライアントで実行されている場合、アドインは要件セット 1.3 の API を使用できます。

より新しい API を使用するために、開発者は標準の JavaScript を使用して新しい API の有無を確認できます。

```js
if (item.somePropertyOrFunction !== undefined) {
  item.somePropertyOrFunction ...
}
```

このようなチェックは、マニフェストで指定された要件セットバージョンに存在する API には必要ありません。

## <a name="choosing-a-minimum-requirement-set"></a>最小要件セットの選択

開発者は、アドインを使用するために必要な、シナリオで必須の API のセットが含まれている初期の要件セットを使用する必要があります。

## <a name="clients"></a>クライアント

以下のクライアントは、Outlook のアドインをサポートしています。

| クライアント | サポートされる API の要件セット |
| --- | --- |
| Outlook 2016 for Windows | 1.1, 1.2, 1.3, 1.4 |
| Outlook 2016 for Mac | 1.1 |
| Windows 版 Outlook 2013 | 1.1、1.2、1.3 |
| Outlook on the web (Office 365 および Outlook.com) | 1.1, 1.2, 1.3, 1.4 |
| Outlook Web App (Exchange 2013 On-Premise) | 1.1 |
| Outlook Web App (Exchange 2016 On-Premise) | 1.1, 1.2. 1.3 |
>**注** Outlook 2013 での 1.3 のサポートは、[2015 年 12 月 8 日付、Outlook 2013 用更新プログラム (KB3114349) ](https://support.microsoft.com/en-us/kb/3114349) の一部として追加されました。    
