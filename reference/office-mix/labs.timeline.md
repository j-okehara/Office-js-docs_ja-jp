
# <a name="labs.timeline"></a>Labs.Timeline

 _**適用対象:** Office 用アプリ | Office アドイン | Office Mix | PowerPoint_

Labs.js タイムライン機能へのアクセスを提供します。

```
class Timeline
```


## <a name="methods"></a>メソッド




### <a name="method"></a>method

 `function constructor(labsInternal: Labs.LabsInternal)`

Creates a new instance of the  **Timeline** class.


### <a name="next"></a>次へ

 `public function next(completionStatus: Labs.Core.ICompletionStatus, callback: Labs.Core.ILabCallback<void>): void`

タイムラインが次のスライドに進むことを示します。

 **Parameters**


|||
|:-----|:-----|
| _completionStatus_|ラボの現在の状況を示します。|
| _callback_|ラボが次のスライドに移動させたときのコールバック関数。|
