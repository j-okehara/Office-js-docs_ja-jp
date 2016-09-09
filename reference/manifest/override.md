
# Override 要素
追加ロケールの設定の値を指定する方法を提供します。

 **アドインの種類:**コンテンツ、作業ウィンドウ、メール


## 構文:


```XML
<Override Locale="string " Value="string " />
```


## 次に含まれる:


||
|:-----|
|[CitationText](../../reference/manifest/citationtext.md)|
|[説明](../../reference/manifest/description.md)|
|[DictionaryName](../../reference/manifest/dictionaryname.md)|
|[DictionaryHomePage](../../reference/manifest/dictionaryhomepage.md)|
|[DisplayName](../../reference/manifest/displayname.md)|
|[HighResolutionIconUrl](../../reference/manifest/highresolutioniconurl.md)|
|[IconUrl](../../reference/manifest/iconurl.md)|
|[QueryUri](../../reference/manifest/queryuri.md)|
|[SourceLocation](../../reference/manifest/sourcelocation.md)|
|[SupportUrl](../../reference/manifest/supporturl.md)|

## 属性



|**属性**|**種類**|**必須**|**説明**|
|:-----|:-----|:-----|:-----|
|ロケール|string|必須|`"en-US"` などの BCP 47 言語タグの書式で、この上書きのロケールのカルチャ名を指定します。|
|Value|string|必須|指定のロケールに対して表される設定の値を指定します。|

## その他のリソース



- [Office アドインのローカライズ](../../docs/develop/localization.md#off15wecon_LocalesManifest)
    
