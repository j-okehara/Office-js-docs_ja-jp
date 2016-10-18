
# <a name="labs.core.actions.isubmitansweroptions"></a>Labs.Core.Actions.ISubmitAnswerOptions

 _**適用対象:** Office 用アプリ | Office アドイン | Office Mix | PowerPoint_

応答送信のアクションで使用できるオプション。

```
interface ISubmitAnswerOptions extends Core.IActionOptions
```


## <a name="properties"></a>プロパティ


|||
|:-----|:-----|
| `componentId: string`|送信に関連付けられているコンポーネント。|
| `attemptId: string`|送信に関連付けられている試行。|
| `answer: any`|送信される応答。|
