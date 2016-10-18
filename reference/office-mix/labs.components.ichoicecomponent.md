
# <a name="labs.components.ichoicecomponent"></a>Labs.Components.IChoiceComponent

 _**適用対象:** Office 用アプリ | Office アドイン | Office Mix | PowerPoint_

選択コンポーネントとの対話を有効にする。

```
interface IChoiceComponent extends Labs.Core.IComponent
```


## <a name="properties"></a>プロパティ


|名前|説明|
|:-----|:-----|
| `choices: Components.IChoice[]`|問題に関連付けられている選択肢の一覧を表す配列。|
| `timeLimit: number`|問題を完了するための時間制限。|
| `maxAttempts: number`|問題に対して許可されている最大試行数。|
| `maxScore: number`|問題の最大スコア。|
| `hasAnswer: boolean`|問題に答えがある場合は、**True** です。|
| `answer: any`|問題の答えです。複数の答えがサポートされている場合は配列、答えが 1 つしかサポートされていない場合は単一の ID です。|
| `secure: boolean`|テストがセキュリティで保護されているかどうかに限らず、セキュリティで保護されたフィールドはユーザーに公表されません。|
