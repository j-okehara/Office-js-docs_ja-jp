
# Labs.takeLab

 _**適用対象:** Office 用アプリ | Office アドイン | Office Mix | PowerPoint_

指定されたラボを実行し、サーバーへのラボの結果の送信を有効にします。編集中にラボを実行することはできないことに注意してください。

```
function takeLab(callback: Core.ILabCallback<LabInstance>): void
```


## パラメーター


|**名前**|**説明**|
|:-----|:-----|
| _callback_|[Labs.LabInstance](../../reference/office-mix/labs.labinstance.md) オブジェクトが作成されると起動されるコールバック メソッド。|
