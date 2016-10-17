
# <a name="labs.components.activitycomponentattempt"></a>Labs.Components.ActivityComponentAttempt

 _**適用対象:** Office 用アプリ | Office アドイン | Office Mix | PowerPoint_

アクティビティ コンポーネントの完了の試行を表す。

```
class Permissions
```


## <a name="methods"></a>メソッド




### <a name="constructor"></a>コンストラクター

 `function constructor(labs: Labs.LabsInternal, componentId: string, attemptId: string, values: {[type:string]: Labs.Core.IValueInstance[]})`

**ActivityComponentAttempt** クラスの新しいインスタンスを作成します。

 **パラメーター**


|**名前**|**説明**|
|:-----|:-----|
| _labs_|コンポーネントに関連付けられているラボのインスタンス ([Labs.LabsInternal](http://msdn.microsoft.com/library/599fb2c4-bb16-4422-84ad-10ed85a14018.aspx))。|
| _componentId_|試行に関連付けられているコンポーネントの ID。|
| _attemptId_|試行の ID。|
| _values_|コンポーネントに関連付けられている値 (ある場合)。|

### <a name="complete"></a>complete

 `public function complete(callback: Labs.Core.ILabCallback<void>): void`

アクティビティが完了したことを示すインジケーター。

 **パラメーター**


|**名前**|**説明**|
|:-----|:-----|
| _callback_|アクティビティが完了すると起動するコールバック関数。|

### <a name="processaction"></a>processAction

 `public function processAction(action: Labs.Core.IAction): void`

特定の試行について取得したアクションに実行し、ラボの状況を生成する関数。

 **パラメーター**


|**名前**|**説明**|
|:-----|:-----|
| _action_|アクション インスタンス ([Labs.Core.IAction](../../reference/office-mix/labs.core.iaction.md))。|
