
# <a name="labs.core.actions.isubmitansweroptions"></a>Labs.Core.Actions.ISubmitAnswerOptions

 _**適用対象:** Office 用アプリ | Office アドイン | Office Mix | PowerPoint_

応答送信のアクションで使用できるオプション。

```
interface ISubmitAnswerOptions extends Core.IActionOptions
```


## <a name="properties"></a>プロパティ


|||
|:-----|:-----|
| `componentId: string`|The component associated with the submission.|
| `attemptId: string`|The attempt associated with the submission.|
| `answer: any`|The answer being submitted.|
