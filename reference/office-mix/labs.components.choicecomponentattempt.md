
# <a name="labs.components.choicecomponentattempt"></a>Labs.Components.ChoiceComponentAttempt

 _**適用対象:** Office 用アプリ | Office アドイン | Office Mix | PowerPoint_

選択コンポーネントでの試行を表す。

```
class ChoiceComponentAttempt extends Components.ComponentAttempt
```


## <a name="methods"></a>メソッド




### <a name="constructor"></a>コンストラクター

 `function constructor(labs: Labs.LabsInternal, componentId: string, attemptId: string, values: {[type:string]: Labs.Core.IValueInstance[]})`

**ChoiceComponentAttempt** クラスの新しいインスタンスを作成します。

 **パラメーター**


|**名前**|**説明**|
|:-----|:-----|
| _labs_|試行に使用する [Labs.LabsInternal](http://msdn.microsoft.com/library/599fb2c4-bb16-4422-84ad-10ed85a14018.aspx) インスタンス。|
| _attemptId_|試行に関連付けられている ID。|
| _values_|試行に関連付けられている値。|

### <a name="timeout"></a>timeout

 `public function timeout(callback: Labs.Core.ILabCallback<void>): void`

ラボがタイムアウトしたことを示します。

 **パラメーター**


|**名前**|**説明**|
|:-----|:-----|
| _callback_|サーバーがタイムアウト メッセージを受信すると起動するコールバック関数。|

### <a name="getsubmissions"></a>getSubmissions

 `public function getSubmissions(): Components.ChoiceComponentSubmission[]`

Retrieves all submissions that have been previously submitted for a given attempt.


### <a name="submit"></a>submit

 `public function submit(answer: Components.ChoiceComponentAnswer, result: Components.ChoiceComponentResult, callback: Labs.Core.ILabCallback<Components.ChoiceComponentSubmission>): void`

ラボで評価され、評価の計算にホストを使用しない新しい応答を送信します。

 **パラメーター**


|**名前**|**説明**|
|:-----|:-----|
| _answer_|試行への応答。|
| _result_|送信の結果。|
| _callback_|送信が受信されると起動するコールバック関数。|

### <a name="processaction"></a>processAction

 `public function processAction(action: Labs.Core.IAction): void`

Initiates processing of the [Labs.Core.IAction](../../reference/office-mix/labs.core.iaction.md) action.

