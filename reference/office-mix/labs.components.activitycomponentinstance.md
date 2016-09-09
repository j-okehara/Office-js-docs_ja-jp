
# Labs.Components.ActivityComponentInstance

 _**適用対象:** Office 用アプリ | Office アドイン | Office Mix | PowerPoint_

アクティビティ コンポーネントの現在のインスタンスを表す。

```
class ActivityComponentInstance extends Labs.ComponentInstance<Components.ActivityComponentAttempt>
```


## プロパティ


|**名前**|**説明**|
|:-----|:-----|
| `public var component: Components.IActivityComponentInstance`|The underlying [Labs.Components.IActivityComponentInstance](../../reference/office-mix/labs.components.iactivitycomponentinstance.md) this class represents|

## メソッド




### コンストラクター

 `function constructor(component: Components.IActivityComponentInstance)`

Creates a new instance of the [Labs.Components.IActivityComponentInstance](../../reference/office-mix/labs.components.iactivitycomponentinstance.md) class.

 **パラメーター**


|**名前**|**説明**|
|:-----|:-----|
| _コンポーネント_|The  **IActivityComponentInstance** to create this class from this class.|

### buildAttempt

 `public function buildAttempt(createAttemptAction: Labs.Core.IAction): Components.ActivityComponentAttempt`

Builds a new  **ActivityComponentAttempt** instance and implements the abstract method defined on the base class

 **パラメーター**


|**名前**|**説明**|
|:-----|:-----|
| _createAttemptResult_|The result of a create attempt action.|
