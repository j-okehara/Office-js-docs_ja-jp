# <a name="getstarted-element"></a>GetStarted 要素

アドインが、Word、Excel、PowerPoint、OneNote のホストにインストールされているときに表示される吹き出しで使用される情報を提供します。**GetStarted** 要素は、[FormFactor](./formfactor.md) の子要素です。

## <a name="child-elements"></a>子要素

| 要素                       | 必須 | 説明                                        |
|:------------------------------|:--------:|:---------------------------------------------------|
| [Title](#title)               | はい      | アドインが機能を公開する場所を定義します。     |
| [説明](#description)   | はい      | JavaScript 関数を含むファイルの URL。|
| [LearnMoreUrl](#learnmoreurl) | いいえ       | アドインの詳細を説明するページの URL。   |


## <a name="title"></a>タイトル 
必須。吹き出しの一番上に使用するタイトル。**resid** 属性は [Resources](./resources.md) セクションの [ShortStrings](./resources.md#shortstrings) 要素にある有効な ID を参照します。

## <a name="description"></a>説明
必須。吹き出しの説明/本文の内容。**resid** 属性は [Resources](./resources.md) セクションの [LongStrings](./resources.md#longstrings) 要素にある有効な ID を参照します。

## <a name="learnmoreurl"></a>LearnMoreUrl
必須。ユーザーがアドインの詳細を参照できるページの URL。**resid** 属性は [Resources](./resources.md) セクションの [Urls](./resources.md#urls) 要素にある有効な ID を参照します。

> **注:** **LearnMoreUrl** は現在、Word、Excel、または PowerPoint のクライアントではレンダリングされません。これが利用可能になったときに URL がレンダリングされるよう、すべてのクライアントにこの URL を追加することをお勧めします。 
