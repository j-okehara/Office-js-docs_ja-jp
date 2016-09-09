
# Labs.Core.Actions.ISubmitAnswerResult

 _**適用対象:** Office 用アプリ | Office アドイン | Office Mix | PowerPoint_

The result of submitting an answer for an attempt.

```
interface ISubmitAnswerResult extends Core.IActionResult
```


## プロパティ


|||
|:-----|:-----|
| `submissionId: string`|送信に関連付けられた ID です。 サーバーによって指定されます。|
| `complete: boolean`|最新の送信により試行が完了した場合は、**true** を返します。|
| `score: any`|Score information associated with the submission.|
