
# <a name="labs.components.choicecomponentresult"></a>Labs.Components.ChoiceComponentResult

 _**適用対象:** Office 用アプリ | Office アドイン | Office Mix | PowerPoint_

選択コンポーネントの送信の結果。

```
class ChoiceComponentResult
```


## <a name="properties"></a>プロパティ


|プロパティ|説明|
|:-----|:-----|
| `public var score: any`|The score associated with the submission.|
| `public var complete: boolean`|Whether or not the result completed the attempt.  **True** if the result completed the attempt.|

## <a name="methods"></a>メソッド




### <a name="constructor"></a>コンストラクター

 `function constructor(score: any, complete: boolean)`

**ChoiceComponentResult** クラスの新しいインスタンスを作成します。

 **パラメーター**


|パラメーター|説明|
|:-----|:-----|
| _score_|結果のスコア。|
| _Complete_|結果で試行が終了しているかどうかを示します。|
