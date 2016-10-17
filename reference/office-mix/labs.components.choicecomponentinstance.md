
# <a name="labs.components.choicecomponentinstance"></a>Labs.Components.ChoiceComponentInstance

 _**適用対象:** Office 用アプリ | Office アドイン | Office Mix | PowerPoint_

選択コンポーネントのインスタンスを表す。

```
class ChoiceComponentInstance extends Labs.ComponentInstance<Components.ChoiceComponentAttempt>
```


## <a name="properties"></a>プロパティ


|プロパティ|説明|
|:-----|:-----|
| `public var component: Components.IChoiceComponentInstance`|The underlying [Labs.Components.IChoiceComponentInstance](../../reference/office-mix/labs.components.ichoicecomponentinstance.md) which this class represents.|

## <a name="methods"></a>メソッド




### <a name="constructor"></a>コンストラクター

 `function constructor(component: Components.IChoiceComponentInstance)`

**ChoiceComponentInstance** クラスの新しいインスタンスを作成します。

 **パラメーター**


|パラメーター|説明|
|:-----|:-----|
| _component_|このクラスの作成元になる [Labs.Components.IChoiceComponentInstance](../../reference/office-mix/labs.components.ichoicecomponentinstance.md) オブジェクト。|

### <a name="buildattempt"></a>buildAttempt

 `public function buildAttempt(createAttemptAction: Labs.Core.IAction): Components.ChoiceComponentAttempt`

新しい **ChoiceComponentAttempt** インスタンスを作成し、基底クラスに定義されている抽象メソッドを実装します。

 **パラメーター**


|パラメーター|説明|
|:-----|:-----|
| _createAttemptResult_|作成の試行アクションの結果。|
