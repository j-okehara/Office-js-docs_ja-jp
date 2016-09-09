
# Labs.Components.ChoiceComponentInstance

 _**適用対象:** Office 用アプリ | Office アドイン | Office Mix | PowerPoint_

選択コンポーネントのインスタンスを表す。

```
class ChoiceComponentInstance extends Labs.ComponentInstance<Components.ChoiceComponentAttempt>
```


## プロパティ


|プロパティ|説明|
|:-----|:-----|
| `public var component: Components.IChoiceComponentInstance`|The underlying [Labs.Components.IChoiceComponentInstance](../../reference/office-mix/labs.components.ichoicecomponentinstance.md) which this class represents.|

## メソッド




### コンストラクター

 `function constructor(component: Components.IChoiceComponentInstance)`

Creates a new instance of the  **ChoiceComponentInstance** class.

 **パラメーター**


|パラメーター|説明|
|:-----|:-----|
| _コンポーネント_|The [Labs.Components.IChoiceComponentInstance](../../reference/office-mix/labs.components.ichoicecomponentinstance.md) object from which to create this class.|

### buildAttempt

 `public function buildAttempt(createAttemptAction: Labs.Core.IAction): Components.ChoiceComponentAttempt`

Builds a new  **ChoiceComponentAttempt** instance and implements the abstract method defined on the base class.

 **パラメーター**


|パラメーター|説明|
|:-----|:-----|
| _createAttemptResult_|The result from the create attempt action.|
