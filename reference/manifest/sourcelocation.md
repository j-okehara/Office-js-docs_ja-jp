
# SourceLocation 要素
Office アドインのソース ファイルの場所を、1 から 2018 文字までの長さの URL として指定します。 ソースの場所はファイル パスではなく、HTTPS アドレスにする必要があります。

 **アドインの種類:**コンテンツ、作業ウィンドウ、メール


## 構文:


```XML
<SourceLocation DefaultValue="string " />
```


## 次に含まれる:

[DefaultSettings](../../reference/manifest/defaultsettings.md) (コンテンツ アドインおよび作業ウィンドウ アドイン)

[FormSettings](../../reference/manifest/formsettings.md) (メール アドイン)


## 含めることができるもの:

[[優先]](../../reference/manifest/override.md)


## 属性



|**属性**|**種類**|**必須**|**説明**|
|:-----|:-----|:-----|:-----|
|DefaultValue|URL|必須|[DefaultLocale](../../reference/manifest/defaultlocale.md) 要素に指定されるロケール用に、この設定の既定値を指定します。|
