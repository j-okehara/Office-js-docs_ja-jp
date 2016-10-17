
# <a name="labs.core.actions.icreatecomponentoptions"></a>Labs.Core.Actions.ICreateComponentOptions

 _**適用対象:** Office 用アプリ | Office アドイン | Office Mix | PowerPoint_

新しいコンポーネントを作成します。

```
interface ICreateComponentOptions extends Core.IActionOptions
```


## <a name="properties"></a>プロパティ


|||
|:-----|:-----|
| `componentId: string`|The component invoking the create component action.|
| `component: Core.IComponent`|作成する [Labs.Core.IComponent](../../reference/office-mix/labs.core.icomponent.md) コンポーネント|
| `correlationId?: string`|ラボのすべてのインスタンスにわたってこのコンポーネントを関連付けるための省略可能なフィールドです。ホストが同じコンポーネントでさまざまな試行を識別できるようにします。|
