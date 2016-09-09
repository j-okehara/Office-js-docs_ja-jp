
# Labs.Components.DynamicComponentInstance

 _**適用対象:** Office 用アプリ | Office アドイン | Office Mix | PowerPoint_

動的コンポーネントのインスタンスを表す。

```
class DynamicComponentInstance extends Labs.ComponentInstanceBase
```


## プロパティ


|プロパティ|説明|
|:-----|:-----|
| `public var component: Components.IDynamicComponentInstance`|コンポーネント インスタンス定義。|

## メソッド




### コンストラクター

 `function constructor(component: Components.IDynamicComponentInstance)`

Creates a new dynamic component instance using the [Labs.Components.IDynamicComponentInstance](../../reference/office-mix/labs.components.idynamiccomponentinstance.md) definition.


### getComponents

 `public function getComponents(callback: Labs.Core.ILabCallback<Labs.ComponentInstanceBase[]>): void`

Retrieves all of the components created by this dynamic component.

 **パラメーター**


|パラメーター|説明|
|:-----|:-----|
| _callback_|Callback function that fires once all of the components have been retrieved.|

### createComponent

 `public function createComponent(component: Labs.Core.IComponent, callback: Labs.Core.ILabCallback<Labs.ComponentInstanceBase>): void`

Creates a new component using the dynamic component as component base.

 **パラメーター**


|パラメーター|説明|
|:-----|:-----|
| _コンポーネント_|The component ([Labs.Core.IComponent](../../reference/office-mix/labs.core.icomponent.md)) from which to create the instance.|
| _callback_|Callback function that fires once the component is created.|

### 閉じる

 `public function close(callback: Labs.Core.ILabCallback<void>): void`

Indicates there will be no additional submissions associated with this component instance.

 **パラメーター**


|パラメーター|説明|
|:-----|:-----|
| _callback_|Callback function that fires once the instance is closed.|

### isClosed

 `public function isClosed(callback: Labs.Core.ILabCallback<boolean>): void`

動的コンポーネントが終了しているかどうかを返します。終了している場合は **true** を返します。

