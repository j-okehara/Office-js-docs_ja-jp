
# Labs.Components.ChoiceComponentAttempt

 _**適用対象:** Office 用アプリ | Office アドイン | Office Mix | PowerPoint_

選択コンポーネントでの試行を表す。

```
class ChoiceComponentAttempt extends Components.ComponentAttempt
```


## メソッド




### コンストラクター

 `function constructor(labs: Labs.LabsInternal, componentId: string, attemptId: string, values: {[type:string]: Labs.Core.IValueInstance[]})`

Creates a new instance of the  **ChoiceComponentAttempt** class.

 **パラメーター**


|**名前**|**説明**|
|:-----|:-----|
| _labs_|The [Labs.LabsInternal](http://msdn.microsoft.com/library/599fb2c4-bb16-4422-84ad-10ed85a14018.aspx) instance to use with the attempt.|
| _attemptId_|The ID associated with the attempt.|
| _values_|The values associated with the attempt.|

### timeout

 `public function timeout(callback: Labs.Core.ILabCallback<void>): void`

Indicates that the lab has timed out.

 **パラメーター**


|**名前**|**説明**|
|:-----|:-----|
| _callback_|Callback functions that fires once the server has received the timeout message.|

### getSubmissions

 `public function getSubmissions(): Components.ChoiceComponentSubmission[]`

Retrieves all submissions that have been previously submitted for a given attempt.


### submit

 `public function submit(answer: Components.ChoiceComponentAnswer, result: Components.ChoiceComponentResult, callback: Labs.Core.ILabCallback<Components.ChoiceComponentSubmission>): void`

Submits a new answer that was graded by the lab and will not use the host to compute a grade.

 **パラメーター**


|**名前**|**説明**|
|:-----|:-----|
| _answer_|The answer for the attempt.|
| _result_|The result of the submission.|
| _callback_|Callback function that fires once the submission has been received.|

### processAction

 `public function processAction(action: Labs.Core.IAction): void`

Initiates processing of the [Labs.Core.IAction](../../reference/office-mix/labs.core.iaction.md) action.

