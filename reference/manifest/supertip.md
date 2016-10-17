## <a name="supertip"></a>Supertip
豊富なヒント (タイトルと説明の両方) を定義します。コントロールの[ボタン](./button.md)と[メニュー](./menu-control.md)の両方で使用されます。 

## <a name="child-elements"></a>子要素
|  要素 |  必須  |  説明  |
|:-----|:-----|:-----|
|  [Title](#title)        | はい |   ヒントのテキストです。         |
|  [説明](#description)  | はい |  ヒントの説明です。    |

## <a name="title"></a>タイトル
必ず指定します。ヒントのテキストです。 **resid** 属性には、 **Resources** 要素の **ShortStrings** 要素にある [String](./resources.md#shortstrings) 要素の [id](./resources.md) の値を設定する必要があります。

## <a name="description"></a>説明
必ず指定します。ヒントの記述です。 **resid** 属性には、 **Resources** 要素の **LongStrings** 要素にある [String](./resources.md#longstrings) 要素の [id](./resources.md) 属性の値を設定する必要があります。

```xml
 <Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
  </Supertip>
```