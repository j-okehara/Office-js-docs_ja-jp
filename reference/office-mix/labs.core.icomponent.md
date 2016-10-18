
# <a name="labs.core.icomponent"></a>Labs.Core.IComponent

 _**適用対象:** Office 用アプリ | Office アドイン | Office Mix | PowerPoint_

ラボのコンポーネントを表す基底クラス。

```
interface IComponent extends Core.ILabObject, Core.IUserData
```


## <a name="properties"></a>プロパティ


|||
|:-----|:-----|
| `name: string`|Name of the component.|
| `values: {[type:string]: Core.IValue[]}`|コンポーネントに関連付けられている Value プロパティ マップ。|
