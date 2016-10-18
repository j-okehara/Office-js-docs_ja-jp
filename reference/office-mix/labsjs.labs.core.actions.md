
# <a name="labsjs.labs.core.actions"></a>LabsJS.Labs.Core.Actions
LabJS.Labs.Core.Actions JavaScript API の概要を示します。

 _**適用対象:** Office 用アプリ | Office アドイン | Office Mix | PowerPoint_

これらの API では、ラボの現在の動作を示す、ラボの操作を表現します。API は、新しいコンポーネントを作成する場合、または新しいドライバー (Office Mix 以外) との接続を開発する場合に便利です。

## <a name="labsjs.labs.core.actions-api-module"></a>LabsJS.Labs.Core.Actions API モジュール

Actions モジュールには次の種類が含まれます。


### <a name="interfaces"></a>インターフェイス


|||
|:-----|:-----|
|[Labs.Core.Actions.ICloseComponentOptions](../../reference/office-mix/labs.core.actions.iclosecomponentoptions.md)|終了するコンポーネント。|
|[Labs.Core.Actions.ICreateAttemptOptions](../../reference/office-mix/labs.core.actions.icreateattemptoptions.md)|試行に関連付けられているコンポーネント。|
|[Labs.Core.Actions.ICreateAttemptResult](../../reference/office-mix/labs.core.actions.icreateattemptresult.md)|指定のコンポーネントに対する試行の結果。|
|[Labs.Core.Actions.ICreateComponentOptions](../../reference/office-mix/labs.core.actions.icreatecomponentoptions.md)|新しいコンポーネントを作成します。|
|[Labs.Core.Actions.ICreateComponentResult](../../reference/office-mix/labs.core.actions.icreatecomponentresult.md)|[Labs.Core.IActionResult](../../reference/office-mix/labs.core.iactionresult.md) によって新しいコンポーネントが作成されます。|
|[Labs.Core.Actions.IGetValueResult](../../reference/office-mix/labs.core.actions.igetvalueresult.md)|get value アクションの結果。|
|[Labs.Core.Actions.ISubmitAnswerResult](../../reference/office-mix/labs.core.actions.isubmitanswerresult.md)|試行に対する応答の送信結果。|
|[Labs.Core.Actions.IAttemptTimeoutOptions](../../reference/office-mix/labs.core.actions.iattempttimeoutoptions.md)|現在の試行のタイムアウト アクションで使用できるオプション。|
|[Labs.Core.Actions.IGetValueOptions](../../reference/office-mix/labs.core.actions.igetvalueoptions.md)|get value 操作で使用できるオプション。|
|[Labs.Core.Actions.IResumeAttemptOptions](../../reference/office-mix/labs.core.actions.iresumeattemptoptions.md)|再開の試行に関連するオプション。|
|[Labs.Core.Actions.ISubmitAnswerOptions](../../reference/office-mix/labs.core.actions.isubmitansweroptions.md)|応答送信のアクションで使用できるオプション。|

### <a name="variables"></a>変数


|||
|:-----|:-----|
| `var CloseComponentAction: string`|コンポーネントを閉じ、今後それに対するアクションがないことを示します。|
| `var CreateAttemptAction: string`|新しい試行を作成するアクションです。|
| `var CreateComponentAction: string`|新しいコンポーネントを作成するアクションです。|
| `var AttemptTimeoutAction: string`|タイムアウト アクションを試行します。|
| `var GetValueAction: string`|試行に関連付けられた値を取得するアクションです。|
| `var ResumeAttemptAction: string`|試行アクションを再開します。ユーザーが指定の試行に対して作業を再開していることを示すために使用されます。|
| `var SubmitAnswerAction: string`|指定の試行で、応答を送信するためのアクションです。|
