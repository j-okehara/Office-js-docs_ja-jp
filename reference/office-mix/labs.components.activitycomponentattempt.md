﻿
# Labs.Components.ActivityComponentAttempt

 _**適用対象:** Office 用アプリ | Office アドイン | Office Mix | PowerPoint_

アクティビティ コンポーネントの完了の試行を表す。

```
class Permissions
```


## メソッド




### コンストラクター

 `function constructor(labs: Labs.LabsInternal, componentId: string, attemptId: string, values: {[type:string]: Labs.Core.IValueInstance[]})`

Creates a new instance of the  **ActivityComponentAttempt** class.

 **パラメーター**


|**名前**|**説明**|
|:-----|:-----|
| _labs_|Lab instances ([Labs.LabsInternal](http://msdn.microsoft.com/library/599fb2c4-bb16-4422-84ad-10ed85a14018.aspx)) associated with the component.|
| _componentId_|ID of the component associated with the attempt.|
| _attemptId_|ID of the attempt.|
| _values_|Values, if any, associated with the component.|

### complete

 `public function complete(callback: Labs.Core.ILabCallback<void>): void`

Indicator that the activity has been completed.

 **パラメーター**


|**名前**|**説明**|
|:-----|:-----|
| _callback_|Callback function that is invoked once the activity has completed.|

### processAction

 `public function processAction(action: Labs.Core.IAction): void`

Function that runs over the actions that are retrieved for a given attempt, then populates the state of the lab.

 **パラメーター**


|**名前**|**説明**|
|:-----|:-----|
| _action_|The action instance ([Labs.Core.IAction](../../reference/office-mix/labs.core.iaction.md)).|