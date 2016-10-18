
# <a name="labs.components.inputcomponentresult"></a>Labs.Components.InputComponentResult

 _**適用対象:** Office 用アプリ | Office アドイン | Office Mix | PowerPoint_

入力コンポーネントの送信の結果。

```
class InputComponentResult
```


## <a name="properties"></a>プロパティ


|プロパティ|説明|
|:-----|:-----|
| `public var score: any`|送信に関連付けられているスコア。|
| `public var complete: boolean`|Indicates whether the result submitted resulted in the completion of the attempt.  **True** if the attempt is completed.|

## <a name="methods"></a>メソッド




### <a name="constructor"></a>コンストラクター

 `function constructor(score: any, complete: boolean)`

**InputComponentResult** クラスの新しいインスタンスを作成します。

 **パラメーター**


|パラメーター|説明|
|:-----|:-----|
| _score_|結果に関連付けられているスコア。|
| _Complete_|試行が完了している結果の場合、ブール値は **true** です。|
