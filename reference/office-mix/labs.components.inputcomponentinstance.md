
# <a name="labs.components.inputcomponentinstance"></a>Labs.Components.InputComponentInstance

 _**適用対象:** Office 用アプリ | Office アドイン | Office Mix | PowerPoint_

入力コンポーネントのインスタンスを表す。

```
class InputComponentInstance extends Labs.ComponentInstance<Components.InputComponentAttempt>
```


## <a name="properties"></a>プロパティ


|プロパティ|説明|
|:-----|:-----|
| `public var component: Components.IInputComponentInstance`|The underlying [Labs.Components.IInputComponentInstance](../../reference/office-mix/labs.components.iinputcomponentinstance.md) object represented by this class.|

## <a name="methods"></a>メソッド




### <a name="constructor"></a>コンストラクター

 `function constructor(component: Components.IInputComponentInstance)`

新しい [Labs.Components.IInputComponentInstance](../../reference/office-mix/labs.components.iinputcomponentinstance.md) インスタンスを作成します。

 **パラメーター**


|パラメーター|説明|
|:-----|:-----|
| _component_|このクラスを作成する [Labs.Components.IInputComponentInstance](../../reference/office-mix/labs.components.iinputcomponentinstance.md)。|

### <a name="buildattempt"></a>buildAttempt

 `public function buildAttempt(createAttemptAction: Labs.Core.IAction): Components.InputComponentAttempt`

新しい [Labs.Components.InputComponentAttempt](../../reference/office-mix/labs.components.inputcomponentattempt.md) を構築します。基本クラスで定義された抽象メソッドを実装します。

 **パラメーター**


|パラメーター|説明|
|:-----|:-----|
| _createAttemptResult_|試行アクション作成の結果。|
