
# <a name="appdomains-element"></a>AppDomains 要素
Office アドインがページを読み込むために使用する追加のドメインを指定します。

 **アドインの種類:**コンテンツ、作業ウィンドウ、メール


## <a name="syntax:"></a>構文:


```XML
<AppDomains>
   ...
</AppDomains>
```


## <a name="contained-in:"></a>次に含まれる:

[OfficeApp](../../reference/manifest/officeapp.md)


## <a name="can-contain:"></a>含めることができるもの:

[AppDomain](../../reference/manifest/appdomain.md)


## <a name="remarks"></a>注釈

**AppDomains** 要素と **AppDomain** 要素は、[SourceLocation](../../reference/manifest/sourcelocation.md) 要素で指定したドメイン以外のものを追加指定するために使用されます。詳細については、「[Office アドインの XML マニフェスト](../../docs/overview/add-in-manifests.md)」をご覧ください。

