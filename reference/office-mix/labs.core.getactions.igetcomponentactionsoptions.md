
# <a name="labs.core.getactions.igetcomponentactionsoptions"></a>Labs.Core.GetActions.IGetComponentActionsOptions

 _**適用対象:** Office 用アプリ | Office アドイン | Office Mix | PowerPoint_

取得コンポーネントの操作に関連付けられているオプションへのアクセスを提供する。

```
interface IGetComponentActionsOptions extends Core.IGetActionOptions
```


## <a name="properties"></a>プロパティ


|||
|:-----|:-----|
| `componentId: string`|検索対象となるコンポーネント。|
| `action: string`|検索対象となるアクションの種類。|
