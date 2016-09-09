
# Requirements 要素
Office アドインをアクティブにするために必要な JavaScript API for Office の最小要件セット ([要件セット](../../docs/overview/specify-office-hosts-and-api-requirements.md#SpecifyRequirementSets_sets)またはメソッド、あるいはその両方) を指定します。

 **アドインの種類:**コンテンツ、作業ウィンドウ、メール


## 構文:


```XML
<Requirements>
   ...
</Requirements>
```


## 次に含まれる:

[OfficeApp](../../reference/manifest/officeapp.md)


## 含めることができるもの:



|**要素**|**コンテンツ**|**メール**|**TaskPane**|
|:-----|:-----|:-----|:-----|
|[Sets](../../reference/manifest/sets.md)|x|x|x|
|[メソッド](../../reference/manifest/methods.md)|x||x|

## 注釈

要件セットの詳細については、「[Office ホストと API 要件を指定する](../../docs/overview/specify-office-hosts-and-api-requirements.md)」をご覧ください。

