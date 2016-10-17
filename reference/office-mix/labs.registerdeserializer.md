
# <a name="labs.registerdeserializer"></a>Labs.registerDeserializer

 _**適用対象:** Office 用アプリ | Office アドイン | Office Mix | PowerPoint_

指定された JSON オブジェクトをオブジェクトに逆シリアル化します。コンポーネントの作成者のみが使用する必要があります。

```
function registerDeserializer(type: string, deserialize: (json: Core.ILabObject): any): void
```


## <a name="parameters"></a>パラメーター


|**名前**|**説明**|
|:-----|:-----|
|json|The [Labs.Core.ILabObject](../../reference/office-mix/labs.core.ilabobject.md) to deserialize.|

## <a name="return-value"></a>Return value

Returns an [Labs.Core.ILabObject](../../reference/office-mix/labs.core.ilabobject.md) instance.

