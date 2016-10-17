
# <a name="labs.components.choicecomponentsubmission"></a>Labs.Components.ChoiceComponentSubmission

 _**適用対象:** Office 用アプリ | Office アドイン | Office Mix | PowerPoint_

選択コンポーネントに関連付けられている送信を表す。

```
class ChoiceComponentSubmission
```


## <a name="properties"></a>プロパティ


|プロパティ|説明|
|:-----|:-----|
| `public var answer: Components.ChoiceComponentAnswer`|The answer ([Labs.Components.ChoiceComponentAnswer](../../reference/office-mix/labs.components.choicecomponentanswer.md)) associated with the submission.|
| `public var result: Components.ChoiceComponentResult`|The result ([Labs.Components.ChoiceComponentResult](../../reference/office-mix/labs.components.choicecomponentresult.md)) of the submission.|
| `public var time: number`|The time at which the submission was received.|

## <a name="methods"></a>メソッド




### <a name="constructor"></a>コンストラクター

 `class ChoiceComponentSubmission`

選択コンポーネントに関連付けられている送信を表す。

 **パラメーター**


|パラメーター|説明|
|:-----|:-----|
| _answer_|送信に関連付けられている応答 ([Labs.Components.ChoiceComponentAnswer](../../reference/office-mix/labs.components.choicecomponentanswer.md))。|
| _result_|送信の結果 ([Labs.Components.ChoiceComponentResult](../../reference/office-mix/labs.components.choicecomponentresult.md))。|
| _time_|送信が受信された時刻。|
