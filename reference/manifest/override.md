
# <a name="override-element"></a>Override 要素
追加ロケールの設定の値を指定する方法を提供します。

 **アドインの種類:**コンテンツ、作業ウィンドウ、メール


## <a name="syntax:"></a>構文:


```XML
<Override Locale="string " Value="string " />
```


## <a name="contained-in:"></a>次に含まれる:


||
|:-----|
|[CitationText](../../reference/manifest/citationtext.md)|
|[Description](../../reference/manifest/description.md)|
|[DictionaryName](../../reference/manifest/dictionaryname.md)|
|[DictionaryHomePage](../../reference/manifest/dictionaryhomepage.md)|
|[DisplayName](../../reference/manifest/displayname.md)|
|[HighResolutionIconUrl](../../reference/manifest/highresolutioniconurl.md)|
|[IconUrl](../../reference/manifest/iconurl.md)|
|[QueryUri](../../reference/manifest/queryuri.md)|
|[SourceLocation](../../reference/manifest/sourcelocation.md)|
|[SupportUrl](../../reference/manifest/supporturl.md)|

## <a name="attributes"></a>属性



|**属性**|**型**|**必須**|**説明**|
|:-----|:-----|:-----|:-----|
|Locale|string|必須|`"en-US"` などの BCP 47 言語タグの書式で、この上書きのロケールのカルチャ名を指定します。|
|Value|string|必須|指定のロケールに対して表される設定の値を指定します。|

## <a name="additional-resources"></a>その他のリソース



- [Office アドインのローカライズ](../../docs/develop/localization.md#off15wecon_LocalesManifest)
    
