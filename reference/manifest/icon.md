﻿# Icon 要素
[Button](./button.md) および [Menu](./menu-control.md) コントロールの **Image** 要素を定義します。

## 子要素
|  要素 |  必須  |  説明  |
|:-----|:-----|:-----|
|  [Image](#image)        | はい |   使用するイメージの resid         |

## Image
ボタンの画像。**resid** 属性には、**Resources** 要素の **Images** 要素にある **Image** 要素の [id](./resources.md) 属性の値を設定する必要があります。**size** 属性は、画像のサイズをピクセル単位で示します。他に 5 つのサイズ (20、24、40、48、64 ピクセル) がサポートされていますが、3 つの画像のサイズ (16、32、80 ピクセル) を必ず指定します。|


```xml
  <Icon>
    <bt:Image size="16" resid="blue-icon-16" />
    <bt:Image size="32" resid="blue-icon-32" />
    <bt:Image size="80" resid="blue-icon-80" />
  </Icon>
```  