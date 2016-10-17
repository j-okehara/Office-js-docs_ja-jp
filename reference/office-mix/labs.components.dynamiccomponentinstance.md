
# <a name="labs.components.dynamiccomponentinstance"></a>Labs.Components.DynamicComponentInstance

 _**適用対象:** Office 用アプリ | Office アドイン | Office Mix | PowerPoint_

動的コンポーネントのインスタンスを表す。

```
class DynamicComponentInstance extends Labs.ComponentInstanceBase
```


## <a name="properties"></a>プロパティ


|プロパティ|説明|
|:-----|:-----|
| `public var component: Components.IDynamicComponentInstance`|コンポーネント インスタンス定義。|

## <a name="methods"></a>メソッド




### <a name="constructor"></a>コンストラクター

 `function constructor(component: Components.IDynamicComponentInstance)`

Creates a new dynamic component instance using the [Labs.Components.IDynamicComponentInstance](../../reference/office-mix/labs.components.idynamiccomponentinstance.md) definition.


### <a name="getcomponents"></a>getComponents

 `public function getComponents(callback: Labs.Core.ILabCallback<Labs.ComponentInstanceBase[]>): void`

この動的コンポーネントで作成されたすべてのコンポーネントを取得します。

 **パラメーター**


|パラメーター|説明|
|:-----|:-----|
| _callback_|すべてのコンポーネントが取得されると起動するコールバック関数。|

### <a name="createcomponent"></a>createComponent

 `public function createComponent(component: Labs.Core.IComponent, callback: Labs.Core.ILabCallback<Labs.ComponentInstanceBase>): void`

動的コンポーネントをコンポーネント ベースとして新しいコンポーネントを作成します。

 **パラメーター**


|パラメーター|説明|
|:-----|:-----|
| _component_|インスタンスの作成元になるコンポーネント ([Labs.Core.IComponent](../../reference/office-mix/labs.core.icomponent.md))。|
| _callback_|コンポーネントが作成されると起動するコールバック関数。|

### <a name="close"></a>閉じる

 `public function close(callback: Labs.Core.ILabCallback<void>): void`

このコンポーネント インスタンスに関連する送信がこれ以上ないことを示します。

 **パラメーター**


|パラメーター|説明|
|:-----|:-----|
| _callback_|インスタンスが終了すると起動するコールバック関数。|

### <a name="isclosed"></a>isClosed

 `public function isClosed(callback: Labs.Core.ILabCallback<boolean>): void`

Returns whether the dynamic component is closed. Returns  **true** if closed.

