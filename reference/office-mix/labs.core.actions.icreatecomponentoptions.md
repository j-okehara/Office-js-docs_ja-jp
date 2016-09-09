
# Labs.Core.Actions.ICreateComponentOptions

 _**適用対象:** Office 用アプリ | Office アドイン | Office Mix | PowerPoint_

Creates a new component.

```
interface ICreateComponentOptions extends Core.IActionOptions
```


## プロパティ


|||
|:-----|:-----|
| `componentId: string`|The component invoking the create component action.|
| `component: Core.IComponent`|作成する [Labs.Core.IComponent](../../reference/office-mix/labs.core.icomponent.md) コンポーネント|
| `correlationId?: string`|ラボのすべてのインスタンスにわたってこのコンポーネントを関連付けるための省略可能なフィールドです。 ホストが同じコンポーネントでさまざまな試行を識別できるようにします。|
