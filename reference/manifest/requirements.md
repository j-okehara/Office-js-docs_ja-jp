
# <a name="requirements-element"></a>Requirements 要素
Office アドインをアクティブにするために必要な JavaScript API for Office の最小要件セット ([要件セット](../../docs/overview/specify-office-hosts-and-api-requirements.md#SpecifyRequirementSets_sets)またはメソッド、あるいはその両方) を指定します。

 **アドインの種類:**コンテンツ、作業ウィンドウ、メール


## <a name="syntax:"></a>構文:


```XML
<Requirements>
   ...
</Requirements>
```


## <a name="contained-in:"></a>次に含まれる:

[OfficeApp](../../reference/manifest/officeapp.md)


## <a name="can-contain:"></a>含めることができるもの:



|**Element**|**Content**|**Mail**|**TaskPane**|
|:-----|:-----|:-----|:-----|
|[Sets](../../reference/manifest/sets.md)|x|x|x|
|[Methods](../../reference/manifest/methods.md)|x||x|

## <a name="remarks"></a>注釈

要件セットの詳細については、「[Office ホストと API 要件を指定する](../../docs/overview/specify-office-hosts-and-api-requirements.md)」をご覧ください。

