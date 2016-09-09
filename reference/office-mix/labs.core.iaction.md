
# Labs.Core.IAction

 _**適用対象:** Office ???? | Office ???? | Office Mix | PowerPoint_

Represents a lab action, which is an interaction that a user has with a specified lab.

```
interface IAction
```


## プロパティ


|||
|:-----|:-----|
| `type: string`|The type of action taken by the user.|
| `options: Core.IActionOptions`|The [Labs.Core.IActionOptions](../../reference/office-mix/labs.core.iactionoptions.md) options sent with the action taken by the user.|
| `result: Core.IActionResult`|The [Labs.Core.IActionResult](../../reference/office-mix/labs.core.iactionresult.md) result of the action.|
| `time: number`|The time at which the action was completed, represented in milliseconds elapsed since 01 January 1970 00:00:00 UTC.|
