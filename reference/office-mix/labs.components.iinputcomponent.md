
# <a name="labs.components.iinputcomponent"></a>Labs.Components.IInputComponent

 _**適用対象:** Office 用アプリ | Office アドイン | Office Mix | PowerPoint_

入力コンポーネントとの対話を有効にする。

```
interface IInputComponent extends Labs.Core.IComponent
```


## <a name="properties"></a>プロパティ


|名前|説明|
|:-----|:-----|
| `maxScore: number`|入力コンポーネントの許容最大スコア。|
| `timeLimit: number`|入力の問題の時間制限。|
| `hasAnswer: boolean`|**True** if the component has an answer.|
| `answer: any`|コンポーネントの問題がある場合は、それに対する応答。|
| `secure: boolean`|**True** if the input component is secure.|
