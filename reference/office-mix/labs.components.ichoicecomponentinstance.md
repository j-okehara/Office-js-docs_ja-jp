
# Labs.Components.IChoiceComponentInstance

 _**適用対象:** Office 用アプリ | Office アドイン | Office Mix | PowerPoint_

選択コンポーネントのインスタンス。

```
interface IChoiceComponentInstance extends Labs.Core.IComponentInstance
```


## プロパティ


|名前|説明|
|:-----|:-----|
| `choices: Components.IChoice[]`|An array representing the list of choices associated with the problem.|
| `timeLimit: number`|Time limit for completing the problem.|
| `maxAttempts: number`|Maximum number of attempts allowed for the problem.|
| `maxScore: number`|The maximum score for the problem.|
| `hasAnswer: boolean`|問題に答えがある場合は、**True** です。|
| `answer: any`|問題の答えです。 複数の答えがサポートされている場合は配列、答えが 1 つしかサポートされていない場合は単一の ID です。|
| `secure: boolean`|テストがセキュリティで保護されているかどうかに限らず、セキュリティで保護されたフィールドはユーザーに公表されません。|
