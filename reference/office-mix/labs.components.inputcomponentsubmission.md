
# <a name="labs.components.inputcomponentsubmission"></a>Labs.Components.InputComponentSubmission

 _**適用対象:** Office 用アプリ | Office アドイン | Office Mix | PowerPoint_

入力コンポーネントへの送信を表す。

```
class InputComponentSubmission
```


## <a name="properties"></a>プロパティ


|プロパティ|説明|
|:-----|:-----|
| `public var answer: Components.InputComponentAnswer`|The answer ([Labs.Components.InputComponentAnswer](../../reference/office-mix/labs.components.inputcomponentanswer.md)) associated with the submission.|
| `public var result: Components.InputComponentResult`|The result ([Labs.Components.InputComponentResult](../../reference/office-mix/labs.components.inputcomponentresult.md)) of the submission.|
| `public var time: number`|The time at which the submission was received.|

## <a name="methods"></a>メソッド




### <a name="constructor"></a>コンストラクター

 `function constructor(answer: Components.InputComponentAnswer, result: Components.InputComponentResult, time: number)`

**InputComponentSubmission** クラスの新しいインスタンスを作成します。

 **パラメーター**


|パラメーター|説明|
|:-----|:-----|
| _answer_|送信に関連付けられている応答。|
| _result_|送信の結果。|
| _time_|送信が受信された時刻。|
