
# <a name="labs.components.idynamiccomponent"></a>Labs.Components.IDynamicComponent

 _**適用対象:** Office 用アプリ | Office アドイン | Office Mix | PowerPoint_

動的コンポーネントとの対話を有効にする。

```
interface IDynamicComponent extends Labs.Core.IComponent
```


## <a name="properties"></a>プロパティ


|名前|説明|
|:-----|:-----|
| `generatedComponentTypes: string[]`|この動的コンポーネントが生成する可能性があるコンポーネントの種類を含む配列。|
| `maxComponents: number`|The maximum number of components that will be generated by this dynamic component. Or  **Labs.Components.Infinite** if there is no cap.|
