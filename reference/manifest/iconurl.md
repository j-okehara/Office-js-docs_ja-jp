
# <a name="iconurl-element"></a>IconUrl 要素
挿入 UX と Office ストアの Office アドインを表すために使用されるイメージの URL を指定します。

 **アドインの種類:**コンテンツ、作業ウィンドウ、メール


## <a name="syntax:"></a>構文:


```XML
<IconUrl DefaultValue="string " />
```


## <a name="can-contain:"></a>含めることができるもの:

[Override](../../reference/manifest/override.md)


## <a name="attributes"></a>属性



|**属性**|**型**|**必須**|**説明**|
|:-----|:-----|:-----|:-----|
|DefaultValue|文字列|必須|この設定の既定値を指定します。この値は、[DefaultLocale](../../reference/manifest/defaultlocale.md) 要素に指定されるロケールを対象としています。|

## <a name="remarks"></a>解説

メール アドインの場合、アイコンは、**[ファイル]**  >  **[アドインの管理]** UI (Outlook) または **[設定]**  >  **[アドインの管理]** UI (Outlook Web App) に表示されます。コンテンツ アドインまたは作業ウィンドウ アドインでは、アイコンは、**[挿入]**  >  **[アドイン]** UI に表示されます。どのアドインの種類についても、アドインを Office ストアに公開すると、アイコンは Office ストア サイトでも使用されます。

このイメージは、GIF、JPG、PNG、EXIF、BMP、TIFF のいずれかのファイル形式である必要があります。コンテンツ アプリおよび作業ウィンドウ アプリの場合、指定するイメージは 32 x 32 ピクセルである必要があります。メール アプリの場合、イメージは 64 x 64 ピクセルである必要があります。また、高 DPI 画面で実行する Office ホスト アプリケーションで使用するアイコンを [HighResolutionIconUrl](../../reference/manifest/highresolutioniconurl.md) 要素を使用して指定する必要もあります。詳細については、「_効果的な Office ストア アプリおよびアドインを作成する_」の「[アプリに一貫性のあるビジュアル ID を作成する](http://msdn.microsoft.com/library/c66a6e6b-2e96-458f-8f8c-2a499fe942c9%28Office.15%29.aspx)」セクションをご覧ください。

