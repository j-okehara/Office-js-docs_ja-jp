
# <a name="labs.components.activitycomponentinstance"></a>Labs.Components.ActivityComponentInstance

 _**適用対象:** Office 用アプリ | Office アドイン | Office Mix | PowerPoint_

アクティビティ コンポーネントの現在のインスタンスを表す。

```
class ActivityComponentInstance extends Labs.ComponentInstance<Components.ActivityComponentAttempt>
```


## <a name="properties"></a>プロパティ


|**名前**|**説明**|
|:-----|:-----|
| `public var component: Components.IActivityComponentInstance`|このクラスが表す、基になる [Labs.Components.IActivityComponentInstance](../../reference/office-mix/labs.components.iactivitycomponentinstance.md)|

## <a name="methods"></a>メソッド




### <a name="constructor"></a>コンストラクター

 `function constructor(component: Components.IActivityComponentInstance)`

[Labs.Components.IActivityComponentInstance](../../reference/office-mix/labs.components.iactivitycomponentinstance.md) クラスの新しいインスタンスを作成します。

 **パラメーター**


|**名前**|**説明**|
|:-----|:-----|
| _component_|このクラスからこのクラスを作成する **IActivityComponentInstance**|

### <a name="buildattempt"></a>buildAttempt

 `public function buildAttempt(createAttemptAction: Labs.Core.IAction): Components.ActivityComponentAttempt`

新しい **ActivityComponentAttempt** インスタンスを作成して、基底クラスに定義されている抽象メソッドを実装します。

 **パラメーター**


|**名前**|**説明**|
|:-----|:-----|
| _createAttemptResult_|試行アクション作成の結果。|
