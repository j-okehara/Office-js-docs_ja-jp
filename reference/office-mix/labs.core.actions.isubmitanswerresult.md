
# <a name="labs.core.actions.isubmitanswerresult"></a>Labs.Core.Actions.ISubmitAnswerResult

 _**適用対象:** Office 用アプリ | Office アドイン | Office Mix | PowerPoint_

試行に対する応答の送信結果。

```
interface ISubmitAnswerResult extends Core.IActionResult
```


## <a name="properties"></a>プロパティ


|||
|:-----|:-----|
| `submissionId: string`|送信に関連付けられた ID です。サーバーによって指定されます。|
| `complete: boolean`|最新の送信により試行が完了した場合は、**true** を返します。|
| `score: any`|送信に関連付けられているスコア情報。|
