
# LabsJS.Labs.Core.Actions
Provides an overview of the LabJS.Labs.Core.Actions JavaScript API.

 _**適用対象:** Office 用のアプリ | Office アドイン | Office Mix | PowerPoint_

これらの API では、ラボの現在の動作を示す、ラボの操作を表現します。 API は、新しいコンポーネントを作成する場合、または新しいドライバー (Office Mix 以外) との接続を開発する場合に便利です。

## LabsJS.Labs.Core.Actions API モジュール

The Actions module contains the following types:


### インターフェイス


|||
|:-----|:-----|
|[Labs.Core.Actions.ICloseComponentOptions](../../reference/office-mix/labs.core.actions.iclosecomponentoptions.md)|The component to close.|
|[Labs.Core.Actions.ICreateAttemptOptions](../../reference/office-mix/labs.core.actions.icreateattemptoptions.md)|The component associated with the attempt.|
|[Labs.Core.Actions.ICreateAttemptResult](../../reference/office-mix/labs.core.actions.icreateattemptresult.md)|The result of creating an attempt for the given component.|
|[Labs.Core.Actions.ICreateComponentOptions](../../reference/office-mix/labs.core.actions.icreatecomponentoptions.md)|Creates a new component.|
|[Labs.Core.Actions.ICreateComponentResult](../../reference/office-mix/labs.core.actions.icreatecomponentresult.md)|The [Labs.Core.IActionResult](../../reference/office-mix/labs.core.iactionresult.md) result of creating a new component.|
|[Labs.Core.Actions.IGetValueResult](../../reference/office-mix/labs.core.actions.igetvalueresult.md)|The result of a get value action.|
|[Labs.Core.Actions.ISubmitAnswerResult](../../reference/office-mix/labs.core.actions.isubmitanswerresult.md)|The result of submitting an answer for an attempt.|
|[Labs.Core.Actions.IAttemptTimeoutOptions](../../reference/office-mix/labs.core.actions.iattempttimeoutoptions.md)|Options available for the current attempt’s timeout action.|
|[Labs.Core.Actions.IGetValueOptions](../../reference/office-mix/labs.core.actions.igetvalueoptions.md)|Options available to the get value operation.|
|[Labs.Core.Actions.IResumeAttemptOptions](../../reference/office-mix/labs.core.actions.iresumeattemptoptions.md)|Options associated with a resume attempt.|
|[Labs.Core.Actions.ISubmitAnswerOptions](../../reference/office-mix/labs.core.actions.isubmitansweroptions.md)|Options available for the submit answer action.|

### 変数


|||
|:-----|:-----|
| `var CloseComponentAction: string`|Closes the component and indicates there will be no future actions against it.|
| `var CreateAttemptAction: string`|Action to create a new attempt.|
| `var CreateComponentAction: string`|Action to create a new component.|
| `var AttemptTimeoutAction: string`|Attempt a timeout action.|
| `var GetValueAction: string`|試行に関連付けられた値を取得するアクションです。|
| `var ResumeAttemptAction: string`|試行アクションを再開します。 ユーザーが指定の試行に対して作業を再開していることを示すために使用されます。|
| `var SubmitAnswerAction: string`|指定の試行で、応答を送信するためのアクションです。|
