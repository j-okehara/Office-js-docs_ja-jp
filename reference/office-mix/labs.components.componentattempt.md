
# Labs.Components.ComponentAttempt

 _**適用対象:** Office 用アプリ | Office アドイン | Office Mix | PowerPoint_

コンポーネントでの試行の基本クラス。

```
class ComponentAttempt
```


## プロパティ


|**名前**|**説明**|
|:-----|:-----|
| `public var _componentId: string`|指定されたコンポーネントの ID。|
| `public var _id: string`|ID of the associated lab.|
| `public var _labs: Labs.LabsInternal`|The lab ([Labs.LabsInternal](http://msdn.microsoft.com/library/599fb2c4-bb16-4422-84ad-10ed85a14018.aspx)) object that is used to interact with the underlying [Labs.Core.ILabHost](../../reference/office-mix/labs.core.ilabhost.md).|
| `public var _resumed: boolean`|**True** if the lab has resumed progress on a given attempt.|
| `public var _state: Labs.ProblemState`|Current state of the attempt as provided by the enum [Labs.ProblemState](../../reference/office-mix/labs.problemstate.md).|
| `public var _values: { [type:string]: Labs.ValueHolder<any>[]}`|Values associated with the attempt, if any, as contained in the [Labs.ValueHolder](../../reference/office-mix/labs.valueholder.md)object.|

## メソッド




### コンストラクター

 `(labs: Labs.LabsInternal, componentId: string, attemptId: string, values: {[type:string]: Labs.Core.IValueInstance[]})`

Creates a new instance of the ComponentAttempt class and provides input parameter values.

 **パラメーター**


|**名前**|**説明**|
|:-----|:-----|
| _labs_|The [Labs.LabsInternal](http://msdn.microsoft.com/library/599fb2c4-bb16-4422-84ad-10ed85a14018.aspx) instance to use with the attempt.|
| _attemptId_|The ID associated with the attempt.|
| _values_|Array of values ([Labs.Core.IValueInstance](../../reference/office-mix/labs.core.ivalueinstance.md)) associated with the attempt.|

### isResumed

 `public function isResumed(): boolean`

ラボが再開されたかどうかを示すブール型の関数です。  ラボが再開された場合は、**True** となります。

 **パラメーター**

なし。


### resume

 `public function resume(callback: Labs.Core.ILabCallback<void>): void`

ラボで特定の試行の進行が再開されたことを示し、このプロセスの一環として既存のデータを読み込みます。 それを使用するには試行を再開する必要があります。

 **パラメーター**


|**名前**|**説明**|
|:-----|:-----|
| _callback_|Callback function that is fired once the attempt has resumed.|

### getState

 `public function getState(): Labs.ProblemState`

Retrieves the state of the lab.

 **パラメーター**

なし。


### processAction

 `public function processAction(action: Labs.Core.IAction): void`

Executes the action associated with the attempt.

 **パラメーター**

なし。


### getValues

 `public function getValues(key: string): Labs.ValueHolder<any>[]`

Retrieves values associated with the attempt

 **パラメーター**


|**名前**|**説明**|
|:-----|:-----|
| _Key_|The key associated with the value in the value map.|
