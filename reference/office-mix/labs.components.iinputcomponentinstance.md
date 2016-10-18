
# <a name="labs.components.iinputcomponentinstance"></a>Labs.Components.IInputComponentInstance

 _**適用対象:** Office 用アプリ | Office アドイン | Office Mix | PowerPoint_

入力コンポーネントのインスタンス。

```
interface IInputComponentInstance extends Labs.Core.IComponentInstance
```


## <a name="properties"></a>プロパティ


|**名前**|**説明**|
|:-----|:-----|
| `maxScore: number`|入力コンポーネントの許容最大スコア。|
| `timeLimit: number`|入力の問題の時間制限。|
| `answer: any`|コンポーネントの問題に対する応答。|
