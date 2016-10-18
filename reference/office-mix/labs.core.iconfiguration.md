
# <a name="labs.core.iconfiguration"></a>Labs.Core.IConfiguration

 _**適用対象:** Office 用アプリ | Office アドイン | Office Mix | PowerPoint_

ラボ構成のデータ構造。

```
interface IConfiguration extends Core.IUserData
```


## <a name="properties"></a>プロパティ


|||
|:-----|:-----|
| `appVersion: Core.IVersion`|この構成に関連付けられているアプリケーションのバージョン。|
| `components: Core.IComponent[]`|ラボに含まれるコンポーネント。|
| `name: string`|ラボの名前。|
| `timeline: Core.ITimelineConfiguration`|ラボのタイムライン構成。|
| `analytics: Core.IAnalyticsConfiguration`|ラボの分析構成。|
