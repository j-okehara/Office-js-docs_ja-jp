
# Labs.LabInstance

 _**適用対象:** Office 用アプリ | Office アドイン | Office Mix | PowerPoint_

現在のユーザーに対して構成されているラボのインスタンスです。このオブジェクトを使用してユーザー用にラボのデータを記録し、取得します。

```
class LabInstance
```


## 変数


|||
|:-----|:-----|
| `public var data: any`|Container variable for holding user data.|
| `public var components: Labs.ComponentInstanceBase[]`|Components that make up the lab instance.|

## メソッド




### getState

 `public function getState(callback: Labs.Core.ILabCallback<any>): void`

Retrieves the current state of the lab for a given user.

 **パラメーター**


|||
|:-----|:-----|
| _callback_|The callback function that fires when the lab state is retrieved.|

### setState

 `public function setState(state: any, callback: Labs.Core.ILabCallback<void>): void`

Sets the state of the lab for a given user.

 **パラメーター**


|||
|:-----|:-----|
| _state_|State to set.|
| _callback_|Callback function that fires once the state is set.|

### Done

 `public function done(callback: Labs.Core.ILabCallback<void>): void`

Indicator function indicating that the user has finished taking the lab.

 **パラメーター**


|||
|:-----|:-----|
| _callback_|Callback function that fires once the lab has finished.|
