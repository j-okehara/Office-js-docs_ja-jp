
# Labs.Timeline

 _**適用対象:** Office 用アプリ | Office アドイン | Office Mix | PowerPoint_

Labs.js タイムライン機能へのアクセスを提供します。

```
class Timeline
```


## メソッド




### method

 `function constructor(labsInternal: Labs.LabsInternal)`

Creates a new instance of the  **Timeline** class.


### 次へ

 `public function next(completionStatus: Labs.Core.ICompletionStatus, callback: Labs.Core.ILabCallback<void>): void`

Indicates that the timeline should advance to the next slide.

 **パラメーター**


|||
|:-----|:-----|
| _completionStatus_|Indicates the current status of the lab.|
| _callback_|Callback function that fires when the lab has moved to the next slide.|
