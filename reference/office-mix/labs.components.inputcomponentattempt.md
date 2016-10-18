
# <a name="labs.components.inputcomponentattempt"></a>Labs.Components.InputComponentAttempt

 _**適用対象:** Office 用アプリ | Office アドイン | Office Mix | PowerPoint_

入力コンポーネントとの対話の試行を表す。

```
class InputComponentAttempt extends Components.ComponentAttempt
```


## <a name="methods"></a>メソッド




### <a name="constructor"></a>コンストラクター

 `function constructor(labs: Labs.LabsInternal, componentId: string, attemptId: string, values: {[type:string]: Labs.Core.IValueInstance[]})`

**InputComponentAttempt** クラスの新しいインスタンスを作成します。

 **パラメーター**


|パラメーター|説明|
|:-----|:-----|
| _labs_|試行に関連付けられているラボ ([Labs.LabsInternal](http://msdn.microsoft.com/library/599fb2c4-bb16-4422-84ad-10ed85a14018.aspx))。|
| _componentID_|試行に関連付けられているコンポーネントの ID。|
| _attemptId_|特定の試行の ID。|
| _values_|値のインスタンスを含む配列 ([Labs.Core.IValueInstance](../../reference/office-mix/labs.core.ivalueinstance.md))。|

### <a name="processaction"></a>processAction

 `public function processAction(action: Labs.Core.IAction): void`

指定の試行で取得したアクションを繰り返し、ラボの状態を生成します。

 **パラメーター**


|パラメーター|説明|
|:-----|:-----|
| _action_|ラボの状態に関連付けられているアクション。|

### <a name="getsubmissions"></a>getSubmissions

 `public function getSubmissions(): Components.InputComponentSubmission[]`

指定の試行で、以前に送信されたすべての送信を取得します。


### <a name="submit"></a>submit

 `public function submit(answer: Components.InputComponentAnswer, result: Components.InputComponentResult, callback: Labs.Core.ILabCallback<Components.InputComponentSubmission>): void`

ラボで評価され、評価の計算にホストを使用しない新しい応答を送信します。

 **パラメーター**


|パラメーター|説明|
|:-----|:-----|
| _answer_|試行に関連付けられている応答。|
| _result_|送信に関連付けられている結果。|
| _callback_|送信が受信されると起動するコールバック関数。|
