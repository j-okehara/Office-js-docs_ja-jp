
# <a name="labs.core.iaction"></a>Labs.Core.IAction

 _**適用対象:** Office 用アプリ | Office アドイン | Office Mix | PowerPoint_

指定されたラボでのユーザー相互作用であるラボのアクションを表します。

```
interface IAction
```


## <a name="properties"></a>プロパティ


|||
|:-----|:-----|
| `type: string`|ユーザーが実行するアクションの種類。|
| `options: Core.IActionOptions`|The [Labs.Core.IActionOptions](../../reference/office-mix/labs.core.iactionoptions.md) options sent with the action taken by the user.|
| `result: Core.IActionResult`|The [Labs.Core.IActionResult](../../reference/office-mix/labs.core.iactionresult.md) result of the action.|
| `time: number`|アクションを完了した時刻です (1 月 1 日 1970 00:00:00 UTC からの経過時間がミリ秒で表されます)。|
