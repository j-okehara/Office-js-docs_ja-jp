﻿# ExtensionPoint 要素

 アドインが Office UI に機能を表示するかどうかを定義します。 **ExtensionPoint** 要素は、[FormFactor](./formfactor.md) の子要素です。 

## 属性

|  属性  |  必須  |  説明  |
|:-----|:-----|:-----|
|  **xsi:type**  |  はい  | 定義される拡張点の種類。|


## Word、Excel、PowerPoint、OneNote アドイン コマンドの拡張点

- **PrimaryCommandSurface** - Office のリボン。
- **ContextMenu**Office UI で右クリックしたときに表示されるショートカット メニュー。

次の例は、 **PrimaryCommandSurface** と **ContextMenu** の属性値を持つ **ExtensionPoint** 要素を使用する方法と、各要素と併用する必要がある子要素を示しています。


 >**重要**  ID 属性を含む要素では、一意の ID を指定してください。会社の名前と ID を使用することをお勧めします。たとえば、次の形式にします。<CustomTab id="mycompanyname.mygroupname">


```XML
 <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <CustomTab id="Contoso Tab">
            <!-- If you want to use a default tab that comes with Office, remove the above CustomTab element, and then uncomment the following OfficeTab element -->
             <!-- <OfficeTab id="TabData"> -->
              <Label resid="residLabel4" />
              <Group id="Group1Id12">
                <Label resid="residLabel4" />
                <Icon>
                  <bt:Image size="16" resid="icon1_32x32" />
                  <bt:Image size="32" resid="icon1_32x32" />
                  <bt:Image size="80" resid="icon1_32x32" />
                </Icon>
                <Tooltip resid="residToolTip" />
                <Control xsi:type="Button" id="Button1Id1">

                   <!-- information about the control -->
                </Control>
                <!-- other controls, as needed -->
              </Group>
            </CustomTab>
          </ExtensionPoint>

        <ExtensionPoint xsi:type="ContextMenu">
          <OfficeMenu id="ContextMenuCell">
            <Control xsi:type="Menu" id="ContextMenu2">
                   <!-- information about the control -->
            </Control>
           <!-- other controls, as needed -->
          </OfficeMenu>
         </ExtensionPoint>
```

**子要素**
 
|**要素**|**説明**|
|:-----|:-----|
|**CustomTab**|カスタム タブをリボンに追加する場合は必須です (  **PrimaryCommandSurface** を使用)。 **CustomTab** 要素を使用する場合、 **OfficeTab** 要素は使用できません。 **id** 属性が必要です。|
|**OfficeTab**|既定の Office リボン タブを拡張する場合は必須です (**PrimaryCommandSurface** を使用)。 **OfficeTab** 要素を使用する場合、**CustomTab** 要素は使用できません。 詳細については、「[OfficeTab](officetab.md)」を参照してください。|
|**OfficeMenu**|既定のコンテキスト メニューにアドイン コマンドを追加する場合は必須です (**ContextMenu** を使用)。 **id** 属性は以下に設定する必要があります。 <br/>Excel または Word の場合は  - **ContextMenuText**。 テキストが選択され、ユーザーが選択されたテキストを右クリックしたときに、コンテキスト メニューに項目が表示されます。 <br/>Excel の場合は  - **ContextMenuCell**。 ユーザーがスプレッドシートのセルを右クリックすると、コンテキスト メニューに項目が表示されます。|
|**Group**|タブのユーザー インターフェイスの拡張点のグループ。グループには、最大 6 個のコントロールを指定できます。 **id** 属性が必要です。id は最大 125 文字の文字列です。|
|**Label**|必須。グループのラベル。 **resid** 属性は、 **String** 要素の **id** 属性の値に設定する必要があります。 **String** 要素は、 **Resources** 要素の子要素である **ShortStrings** 要素の子要素です。|
|**Icon**|必須。小さいフォーム ファクターのデバイス、または表示されるボタンが多すぎるときに使用されるグループのアイコンを指定します。 **resid** 属性は、 **Image** 要素の **id** 属性の値に設定する必要があります。 **Image** 要素は、 **Resources** 要素の子要素である **Images** 要素の子要素です。 **size** 属性は、イメージのサイズをピクセル単位で指定します。3 つのイメージのサイズ (16、32、80) が必要です。5 つのオプションのサイズ (20、24、40、48、64) もサポートされています。|
|**Tooltip**|省略可能。グループのツールヒント。 **resid** 属性は、 **String** 要素の **id** 属性の値に設定する必要があります。 **String** 要素は、 **Resources** 要素の子要素である **LongStrings** 要素の子要素です。|
|**Control**|各グループには、1 つ以上のコントロールが必要です。 **Control** 要素は、 **Button** または **Menu** のいずれかにすることができます。ボタンのコントロールのドロップダウンリストを指定するには、 **Menu** を使用します。現在、ボタンとメニューのみがサポートされています。詳しくは、「 [ボタン コントロール](#ボタン-コントロール)」および「 [メニュー コントロール](#メニュー-コントロール)」のセクションをご覧ください。<br/>**注** トラブルシューティングを簡単にするために、**Control** 要素と関連する **Resources** 子要素を一度に 1 つずつ追加することをお勧めします。

## Outlook アドイン コマンドの拡張点

- [CustomPane](#custompane) 
- [MessageReadCommandSurface](#messagereadcommandsurface) 
- [MessageComposeCommandSurface](#messagecomposecommandsurface) 
- [AppointmentOrganizerCommandSurface](#appointmentorganizercommandsurface) 
- [AppointmentAttendeeCommandSurface](#appointmentattendeecommandsurface)
- [Module](#module) ([DesktopFormFactor](./formfactor.md) でのみ使用できます。)

### CustomPane

CustomPane 拡張点には、指定したルールが満たされた場合にアクティブになるアドインを定義します。 これは閲覧フォーム専用で、水平方向のウィンドウに表示されます。 

**子要素**

|  要素 |  必須  |  説明  |
|:-----|:-----|:-----|
|  **RequestedHeight** | いいえ |  デスクトップ コンピューターで実行する場合に、表示ウィンドウに必要な高さをピクセル単位で指定します。 32 から 450 ピクセルを指定できます。  |
|  **SourceLocation**  | はい |  アドインのソース コード ファイルの URL。 これは、[Resources](./resources.md) 要素の **Url** 要素を指します。  |
|  **Rule**  | はい |  アドインをいつアクティブ化するかを指定するルールのルールまたはコレクション。 詳細については、「[Outlook アドインのアクティブ化ルール](../../outlook/manifests/activation-rules.md)」を参照してください。 |
|  **DisableEntityHighlighting**  | いいえ |  エンティティの強調表示をオフにするかどうかを指定します。 |


#### CustomPane の例
```xml
<ExtensionPoint xsi:type="CustomPane">
   <RequestedHeight>100< /RequestedHeight> 
   <SourceLocation resid="residReadTaskpaneUrl"/>
   <Rule xsi:type="RuleCollection" Mode="Or">
     <Rule xsi:type="ItemIs" ItemType="Message"/>
     <Rule xsi:type="ItemHasAttachment"/>
     <Rule xsi:type="ItemHasKnownEntity" EntityType="Address"/>
   </Rule>
</ExtensionPoint>
```

### MessageReadCommandSurface
この拡張点により、メールの閲覧ビューのコマンド サーフェスにボタンが配置されます。 Outlook デスクトップでは、これはリボンに表示されます。

**子要素**

|  要素 |  説明  |
|:-----|:-----|
|  [OfficeTab](./officetab.md) |  コマンドを既定のリボン タブに追加します。  |
|  [CustomTab](./customtab.md) |  コマンドをカスタム リボン タブに追加します。  |

#### OfficeTab の例
```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### CustomTab の例
```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```
### MessageComposeCommandSurface
この拡張点は、メールの新規作成フォームを使用してアドイン用のリボンにボタンを配置します。 

**子要素**

|  要素 |  説明  |
|:-----|:-----|
|  [OfficeTab](./officetab.md) |  コマンドを既定のリボン タブに追加します。  |
|  [CustomTab](./customtab.md) |  コマンドをカスタム リボン タブに追加します。  |

#### OfficeTab の例
```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### CustomTab の例

```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```
### AppointmentOrganizerCommandSurface

この拡張点は、会議の開催者に表示されるフォームのリボンにボタンを配置します。 

**子要素**

|  要素 |  説明  |
|:-----|:-----|
|  [OfficeTab](./officetab.md) |  コマンドを既定のリボン タブに追加します。  |
|  [CustomTab](./customtab.md) |  コマンドをカスタム リボン タブに追加します。  |

#### OfficeTab の例
```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### CustomTab の例
```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### AppointmentAttendeeCommandSurface

この拡張点は、会議の出席者に表示されるフォームのリボンにボタンを配置します。 

**子要素**

|  要素 |  説明  |
|:-----|:-----|
|  [OfficeTab](./officetab.md) |  コマンドを既定のリボン タブに追加します。  |
|  [CustomTab](./customtab.md) |  コマンドをカスタム リボン タブに追加します。  |

#### OfficeTab の例
```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### CustomTab の例
```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### Module

この拡張点は、モジュール拡張機能用のリボンにボタンを配置します。 

**子要素**

|  要素 |  説明  |
|:-----|:-----|
|  [OfficeTab](./officetab.md) |  コマンドを既定のリボン タブに追加します。  |
|  [CustomTab](./customtab.md) |  コマンドをカスタム リボン タブに追加します。  |
