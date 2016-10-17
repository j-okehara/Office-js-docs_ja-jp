
# <a name="labs.components.iinputcomponent"></a>Labs.Components.IInputComponent

 _**適用対象:** Office 用アプリ | Office アドイン | Office Mix | PowerPoint_

入力コンポーネントとの対話を有効にする。

```
interface IInputComponent extends Labs.Core.IComponent
```


## <a name="properties"></a>プロパティ


|名前|説明|
|:-----|:-----|
| `maxScore: number`|The maximum allowable score for the input component.|
| `timeLimit: number`|Time limit for the input problem.|
| `hasAnswer: boolean`|**True** if the component has an answer.|
| `answer: any`|The answer to the component problem, if any.|
| `secure: boolean`|**True** if the input component is secure.|
