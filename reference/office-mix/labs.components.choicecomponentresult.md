
# Labs.Components.ChoiceComponentResult

 _**適用対象:** Office 用アプリ | Office アドイン | Office Mix | PowerPoint_

選択コンポーネントの送信の結果。

```
class ChoiceComponentResult
```


## プロパティ


|プロパティ|説明|
|:-----|:-----|
| `public var score: any`|送信と関連付けられたスコアです。|
| `public var complete: boolean`|結果で試行が終了しているかどうかを示します。結果で試行が終了している場合は、**True** となります。|

## メソッド




### コンストラクター

 `function constructor(score: any, complete: boolean)`

Creates a new instance of the  **ChoiceComponentResult** class.

 **パラメーター**


|パラメーター|説明|
|:-----|:-----|
| _score_|The score of the result.|
| _complete_|Indicates whether the result completed the attempt.|
