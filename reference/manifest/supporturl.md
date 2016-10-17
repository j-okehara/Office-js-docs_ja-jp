
# <a name="supporturl-element"></a>SupportUrl 要素
アドインのサポート情報を提供するページの URL を指定します。

 **アドインの種類:**コンテンツ、作業ウィンドウ、メール


## <a name="syntax:"></a>構文:


```XML
<SupportUrl DefaultValue="string " />
```


## <a name="can-contain:"></a>含めることができるもの:

[Override](../../reference/manifest/override.md)


## <a name="attributes"></a>属性



|**属性**|**型**|**必須**|**説明**|
|:-----|:-----|:-----|:-----|
|DefaultValue|URL|必須|この設定の既定値を指定します。この値は、[DefaultLocale](../../reference/manifest/defaultlocale.md) 要素に指定されるロケールを対象としています。|
