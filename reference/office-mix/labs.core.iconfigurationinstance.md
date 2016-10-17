
# <a name="labs.core.iconfigurationinstance"></a>Labs.Core.IConfigurationInstance

 _**適用対象:** Office 用アプリ | Office アドイン | Office Mix | PowerPoint_

ラボ構成のインスタンスの基底クラス。インスタンスは、指定のユーザーの構成をインスタンス化したものであり、ラボの特定の実行における構成の変換ビューが含まれています。このビューでは、非表示の情報 (ヒントや回答など) が除外される場合があります。また、このビューには、さまざまなインスタンスを識別する ID が含まれます。

```
interface IConfigurationInstance extends Core.IUserData
```


## <a name="properties"></a>プロパティ


|||
|:-----|:-----|
| `appVersion: Core.IVersion`|Version of the lab associated with this configuration.|
| `components: Core.IComponentInstance[]`|Components associated with the lab.|
| `name: string`|Name of the lab.|
| `timeline: Core.ITimelineConfiguration`|Timeline configuration for the lab.|
