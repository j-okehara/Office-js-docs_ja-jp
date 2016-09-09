
# HighResolutionIconUrl 要素
高 DPI の画面での挿入 UX と Office ストアの Office アドインを表すために使用されるイメージの URL を指定します。

 **アドインの種類:**コンテンツ、作業ウィンドウ、メール


## 構文:


```XML
<HighResolutionIconUrl DefaultValue="string " />
```


## 含めることができるもの:

[[優先]](../../reference/manifest/override.md)


## 属性



|**属性**|**種類**|**必須**|**説明**|
|:-----|:-----|:-----|:-----|
|DefaultValue|文字列 (URL)|必須|この設定の既定値を指定します。この値は、[DefaultLocale](../../reference/manifest/defaultlocale.md) 要素に指定されるロケールを対象としています。|

## 注釈

メール アドインの場合、アイコンは、**[ファイル]**  >  **[アドインの管理]** UI に表示されます。コンテンツ アドインまたは作業ウィンドウ アドインでは、アイコンは、**[挿入]**  >  **[アドイン]** UI に表示されます。

イメージは推奨解像度が 64 x 64 ピクセルであり、次のファイル形式のいずれかである必要があります。GIF、JPG、PNG、EXIF、BMP、または TIFF。 詳細については、「[効果的な Office ストア アプリおよびアドインを作成する](http://msdn.microsoft.com/library/c66a6e6b-2e96-458f-8f8c-2a499fe942c9%28Office.15%29.aspx)」の「_アプリに一貫性のあるビジュアル ID を作成する_」セクションをご覧ください。

