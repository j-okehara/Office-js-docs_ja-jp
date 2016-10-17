
# <a name="labs.core.icomponentinstance"></a>Labs.Core.IComponentInstance

 _**適用対象:** Office 用アプリ | Office アドイン | Office Mix | PowerPoint_

ラボ コンポーネントのインスタンスの基底クラス。

```
interface IComponentInstance extends Core.ILabObject, Core.IUserData
```


## <a name="properties"></a>プロパティ


|||
|:-----|:-----|
| `componentId: string`|The ID of the component this instance is associated with.|
| `name: string`|Name of the component.|
| `values: {[type:string]: Core.IValueInstance[]}`|The value property map associated with the component.|

## <a name="remarks"></a>注釈

コンポーネントのインスタンスは、ユーザーのコンポーネントをインスタンス化したものです。これには、ラボの特定の実行におけるコンポーネントの変換ビューが含まれます。このビューでは、非表示の情報 (回答、ヒントなど) が除外される場合があります。また、このビューには、さまざまなインスタンスを識別する ID が含まれます。

