
# Labs.registerDeserializer

 _**適用対象:** Office 用アプリ | Office アドイン | Office Mix | PowerPoint_

指定された JSON オブジェクトをオブジェクトに逆シリアル化します。コンポーネントの作成者のみが使用する必要があります。

```
function registerDeserializer(type: string, deserialize: (json: Core.ILabObject): any): void
```


## パラメーター


|**名前**|**説明**|
|:-----|:-----|
|json|The [Labs.Core.ILabObject](../../reference/office-mix/labs.core.ilabobject.md) to deserialize.|

## Return value

Returns an [Labs.Core.ILabObject](../../reference/office-mix/labs.core.ilabobject.md) instance.

