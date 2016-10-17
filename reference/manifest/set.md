
# <a name="set-element"></a>Set 要素
Office アドインをアクティブにするために必要な JavaScript API for Office の要件セットを指定します。

 **アドインの種類:**コンテンツ、作業ウィンドウ、メール


## <a name="syntax:"></a>構文:


```XML
<Set Name="string " MinVersion="n .n ">
```


## <a name="contained-in:"></a>次に含まれる:

[Sets](../../reference/manifest/sets.md)


## <a name="attributes"></a>属性



|**属性**|**型**|**必須**|**説明**|
|:-----|:-----|:-----|:-----|
|名前|string|必須|[要件セット](../../docs/overview/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest)の名前。|
|MinVersion|文字列|省略可能|アドインに必要な API セットの最小バージョンを指定します。**DefaultMinVersion** の値が親の [Sets](../../reference/manifest/sets.md) 要素に指定されている場合は、その値を上書きします。|

## <a name="remarks"></a>解説

要件セットの詳細については、「[Office ホストと API 要件を指定する](../../docs/overview/specify-office-hosts-and-api-requirements.md#specify-office-hosts-and-api-requirements)」をご覧ください。

**Set** 要素の **MinVersion** 属性と **Sets** 要素の **DefaultMinVersion** 属性の詳細については、「[Office ホストと API 要件を指定する](../../docs/overview/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest)」をご覧ください。


 >**重要**  メール アドインでは、利用できるのは `"Mailbox"` 要件セット 1 つのみです。この要件セットには、Outlook のメール アドインでサポートされている API のサブセット全体が含まれ、メール アドインのマニフェストで `"Mailbox"` 要件セットを指定する必要があります (コンテンツ アドインと作業ウィンドウ アドインの場合とは異なり、オプションではありません)。また、メール アドインで特定のメソッドのサポートを宣言することもできません。

