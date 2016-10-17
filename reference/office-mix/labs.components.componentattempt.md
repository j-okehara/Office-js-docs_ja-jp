
# <a name="labs.components.componentattempt"></a>Labs.Components.ComponentAttempt

 _**適用対象:** Office 用アプリ | Office アドイン | Office Mix | PowerPoint_

コンポーネントでの試行の基底クラス。

```
class ComponentAttempt
```


## <a name="properties"></a>プロパティ


|**名前**|**説明**|
|:-----|:-----|
| `public var _componentId: string`|指定されたコンポーネントの ID。|
| `public var _id: string`|ID of the associated lab.|
| `public var _labs: Labs.LabsInternal`|The lab ([Labs.LabsInternal](http://msdn.microsoft.com/library/599fb2c4-bb16-4422-84ad-10ed85a14018.aspx)) object that is used to interact with the underlying [Labs.Core.ILabHost](../../reference/office-mix/labs.core.ilabhost.md).|
| `public var _resumed: boolean`|**True** if the lab has resumed progress on a given attempt.|
| `public var _state: Labs.ProblemState`|Current state of the attempt as provided by the enum [Labs.ProblemState](../../reference/office-mix/labs.problemstate.md).|
| `public var _values: { [type:string]: Labs.ValueHolder<any>[]}`|Values associated with the attempt, if any, as contained in the [Labs.ValueHolder](../../reference/office-mix/labs.valueholder.md)object.|

## <a name="methods"></a>メソッド




### <a name="constructor"></a>コンストラクター

 `(labs: Labs.LabsInternal, componentId: string, attemptId: string, values: {[type:string]: Labs.Core.IValueInstance[]})`

ComponentAttempt クラスの新しいインスタンスを作成し、入力パラメーター値を指定します。

 **パラメーター**


|**名前**|**説明**|
|:-----|:-----|
| _labs_|試行に使用する [Labs.LabsInternal](http://msdn.microsoft.com/library/599fb2c4-bb16-4422-84ad-10ed85a14018.aspx) インスタンス。|
| _attemptId_|試行に関連付けられている ID。|
| _values_|試行に関連付けられている値の配列 ([Labs.Core.IValueInstance](../../reference/office-mix/labs.core.ivalueinstance.md))。|

### <a name="isresumed"></a>isResumed

 `public function isResumed(): boolean`

ラボが再開されたかどうかを示すブール型の関数です。ラボが再開された場合は、**True** となります。

 **パラメーター**

なし。


### <a name="resume"></a>resume

 `public function resume(callback: Labs.Core.ILabCallback<void>): void`

ラボで特定の試行の進行が再開されたことを示し、このプロセスの一環として既存のデータを読み込みます。それを使用するには試行を再開する必要があります。

 **パラメーター**


|**名前**|**説明**|
|:-----|:-----|
| _callback_|試行が再開されると起動するコールバック関数。|

### <a name="getstate"></a>getState

 `public function getState(): Labs.ProblemState`

ラボの状態を取得します。

 **パラメーター**

なし。


### <a name="processaction"></a>processAction

 `public function processAction(action: Labs.Core.IAction): void`

試行に関連するアクションを実行します。

 **パラメーター**

なし。


### <a name="getvalues"></a>getValues

 `public function getValues(key: string): Labs.ValueHolder<any>[]`

試行に関連する値を取得します。

 **パラメーター**


|**名前**|**説明**|
|:-----|:-----|
| _key_|値マップにある値と関連付けられているキー。|
