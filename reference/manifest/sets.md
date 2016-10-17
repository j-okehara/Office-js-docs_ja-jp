
# <a name="sets-element"></a>Sets 要素
Office アドインをアクティブにするために必要な JavaScript API for Office の最小限のサブセットを指定します。

 **アドインの種類:**コンテンツ、作業ウィンドウ、メール


## <a name="syntax:"></a>構文:


```XML
<Sets DefaultMinVersion="n .n ">
   ...
</Sets>
```


## <a name="contained-in:"></a>次に含まれる:

[Requirements](../../reference/manifest/requirements.md)


## <a name="can-contain:"></a>含めることができるもの:

[Set](../../reference/manifest/set.md)


## <a name="attributes"></a>属性



|**属性**|**型**|**必須**|**説明**|
|:-----|:-----|:-----|:-----|
|DefaultMinVersion|文字列|省略可能|すべての子の **Set** 要素に対して、既定の [MinVersion](../../reference/manifest/set.md) 属性値を指定します。既定値は "1.1" です。|

## <a name="remarks"></a>注釈

要件セットの詳細については、「[Office ホストと API 要件を指定する](../../docs/overview/specify-office-hosts-and-api-requirements.md)」をご覧ください。

**Set** 要素の **MinVersion** 属性と **Sets** 要素の **DefaultMinVersion** 属性の詳細については、「[マニフェストで Requirements 要素を設定する](../../docs/overview/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest)」をご覧ください。

