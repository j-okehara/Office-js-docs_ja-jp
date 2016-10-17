
# <a name="labs.labinstance"></a>Labs.LabInstance

 _**適用対象:** Office 用アプリ | Office アドイン | Office Mix | PowerPoint_

現在のユーザーに対して構成されているラボのインスタンスです。このオブジェクトを使用してユーザー用にラボのデータを記録し、取得します。

```
class LabInstance
```


## <a name="variables"></a>変数


|||
|:-----|:-----|
| `public var data: any`|Container variable for holding user data.|
| `public var components: Labs.ComponentInstanceBase[]`|Components that make up the lab instance.|

## <a name="methods"></a>メソッド




### <a name="getstate"></a>getState

 `public function getState(callback: Labs.Core.ILabCallback<any>): void`

特定のユーザー用のラボの現在の状態を取得します。

 **Parameters**


|||
|:-----|:-----|
| _callback_|ラボの状態を取得するときに起動するコールバック関数。|

### <a name="setstate"></a>setState

 `public function setState(state: any, callback: Labs.Core.ILabCallback<void>): void`

特定のユーザー用のラボの状態を設定します。

 **Parameters**


|||
|:-----|:-----|
| _state_|設定する状態。|
| _callback_|状態が設定されると起動するコールバック関数。|

### <a name="done"></a>Done

 `public function done(callback: Labs.Core.ILabCallback<void>): void`

ユーザーがラボの取得を完了したことを示すインジケーター関数。

 **Parameters**


|||
|:-----|:-----|
| _callback_|ラボが完了すると起動するコールバック関数。|
