﻿# OfficeMenu 要素
Office のコンテキスト メニューに追加するコントロールのコレクションを定義します。 Word、Excel、PowerPoint、OneNote アドインに適用されます。

## 属性

| 属性            | 必須 | 説明                          |
|:---------------------|:--------:|:-------------------------------------|
| [xsi:type](#xsitype) | はい      | 定義する OfficeMenu の種類。|

## 子要素
|  要素 |  必須  |  説明  |
|:-----|:-----|:-----|
|  [コントロール](#コントロール)    | はい |  1 つ以上のコントロール オブジェクトのコレクション。  |

## xsi:type
この Office アドインを追加する Office クライアント アプリケーションの組み込みメニューを指定します。

- `ContextMenuText` -  テキストが選ばれ、選ばれたテキストのコンテキスト メニューをユーザーが開いたときに (右クリック)、コンテキスト メニューに項目が表示されます。 Word、Excel、PowerPoint、OneNote に適用されます。
- `ContextMenuCell` -  ユーザーがスプレッドシートのセルのコンテキスト メニューを開くと (右クリック)、コンテキスト メニューに項目が表示されます。 Excel に適用されます。 

## コントロール

各 **OfficeMenu** 要素には、1 つ以上の [メニュー](./menu.md#menu-control) コントロールが必要です。 


## 例

```xml
<OfficeMenu id="ContextMenuCell">
    <Control xsi:type="Menu" id="myMenuID">
      <Label resid="residLabel3" />
      <Supertip>
          <Title resid="residLabel" />
          <Description resid="residToolTip" />
      </Supertip>   
      <Icon>
        <bt:Image size="16" resid="icon1_16x16" />
        <bt:Image size="32" resid="icon1_32x32" />
        <bt:Image size="80" resid="icon1_80x80" />
      </Icon>    
      <Items>
        <Item id="myMenuItemID">
          <Label resid="residLabel3"/>
          <Supertip>
            <Title resid="residLabel" />
            <Description resid="residToolTip" />
          </Supertip>
          <Icon>
            <bt:Image size="16" resid="icon1_16x16" />
            <bt:Image size="32" resid="icon1_32x32" />
            <bt:Image size="80" resid="icon1_80x80" />
          </Icon>    
          <Action xsi:type="ShowTaskpane">
            <SourceLocation resid="residTaskpaneUrl2" />    
          </Action>    
        </Item>
      </Items>
    </Control>   
</OfficeMenu>
```