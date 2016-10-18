
# <a name="labs.core.iconfigurationinstance"></a>Labs.Core.IConfigurationInstance

 _**適用対象:** Office 用アプリ | Office アドイン | Office Mix | PowerPoint_

ラボ構成のインスタンスの基底クラス。インスタンスは、指定のユーザーの構成をインスタンス化したものであり、ラボの特定の実行における構成の変換ビューが含まれています。このビューでは、非表示の情報 (ヒントや回答など) が除外される場合があります。また、このビューには、さまざまなインスタンスを識別する ID が含まれます。

```
interface IConfigurationInstance extends Core.IUserData
```


## <a name="properties"></a>プロパティ


|||
|:-----|:-----|
| `appVersion: Core.IVersion`|この構成に関連付けられているラボのバージョン。|
| `components: Core.IComponentInstance[]`|ラボに関連付けられているコンポーネント。|
| `name: string`|ラボの名前。|
| `timeline: Core.ITimelineConfiguration`|ラボのタイムライン構成。|
