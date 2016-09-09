
# Labs.Components.InputComponentInstance

 _**適用対象:** Office 用アプリ | Office アドイン | Office Mix | PowerPoint_

入力コンポーネントのインスタンスを表す。

```
class InputComponentInstance extends Labs.ComponentInstance<Components.InputComponentAttempt>
```


## プロパティ


|プロパティ|説明|
|:-----|:-----|
| `public var component: Components.IInputComponentInstance`|The underlying [Labs.Components.IInputComponentInstance](../../reference/office-mix/labs.components.iinputcomponentinstance.md) object represented by this class.|

## メソッド




### コンストラクター

 `function constructor(component: Components.IInputComponentInstance)`

Creates a new [Labs.Components.IInputComponentInstance](../../reference/office-mix/labs.components.iinputcomponentinstance.md) instance.

 **パラメーター**


|パラメーター|説明|
|:-----|:-----|
| _コンポーネント_|The [Labs.Components.IInputComponentInstance](../../reference/office-mix/labs.components.iinputcomponentinstance.md) from which to create this class.|

### buildAttempt

 `public function buildAttempt(createAttemptAction: Labs.Core.IAction): Components.InputComponentAttempt`

新しい [Labs.Components.InputComponentAttempt](../../reference/office-mix/labs.components.inputcomponentattempt.md) を構築します。基本クラスで定義された抽象メソッドを実装します。

 **パラメーター**


|パラメーター|説明|
|:-----|:-----|
| _createAttemptResult_|The result of a create attempt action.|
