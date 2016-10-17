
# <a name="labs.core.ivalueinstance"></a>Labs.Core.IValueInstance

 _**適用対象:** Office 用アプリ | Office アドイン | Office Mix | PowerPoint_

値データを含む場合の [Labs.Core.IValue](../../reference/office-mix/labs.core.ivalue.md) オブジェクト インスタンス。

```
interface IValueInstance
```


## <a name="properties"></a>プロパティ


|||
|:-----|:-----|
| `valueId: string`|ID of the value which this instance represents.|
| `isHint: boolean`|Boolean  **true** if this value is considered a hint.|
| `hasValue: boolean`|インスタンス情報に値が含まれていない場合は、ブール値 (**true**) です。|
| `value?: any`|値です。このパラメーターは、非表示かどうかによって、設定される場合と設定されない場合があります。|
