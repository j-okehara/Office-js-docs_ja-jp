
# <a name="labs.core.iaction"></a>Labs.Core.IAction

 _**適用対象:** Office 用アプリ | Office アドイン | Office Mix | PowerPoint_

指定されたラボでのユーザー相互作用であるラボのアクションを表します。

```
interface IAction
```


## <a name="properties"></a>プロパティ


|||
|:-----|:-----|
| `type: string`|The type of action taken by the user.|
| `options: Core.IActionOptions`|The [Labs.Core.IActionOptions](../../reference/office-mix/labs.core.iactionoptions.md) options sent with the action taken by the user.|
| `result: Core.IActionResult`|The [Labs.Core.IActionResult](../../reference/office-mix/labs.core.iactionresult.md) result of the action.|
| `time: number`|The time at which the action was completed, represented in milliseconds elapsed since 01 January 1970 00:00:00 UTC.|
