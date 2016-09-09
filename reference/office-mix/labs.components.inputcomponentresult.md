
# Labs.Components.InputComponentResult

 _**適用対象:** Office 用アプリ | Office アドイン | Office Mix | PowerPoint_

入力コンポーネントの送信の結果。

```
class InputComponentResult
```


## プロパティ


|プロパティ|説明|
|:-----|:-----|
| `public var score: any`|送信と関連付けられたスコアです。|
| `public var complete: boolean`|送信された結果によって試行が完了したかどうかを示します。試行が完了している場合は、**True** です。|

## メソッド




### コンストラクター

 `function constructor(score: any, complete: boolean)`

Creates a new instance of the  **InputComponentResult** class.

 **パラメーター**


|パラメーター|説明|
|:-----|:-----|
| _score_|The score associated with the result.|
| _complete_|Boolean  **true** if the result completed the attempt.|
