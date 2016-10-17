
# <a name="labs.core.iconfiguration"></a>Labs.Core.IConfiguration

 _**適用対象:** Office 用アプリ | Office アドイン | Office Mix | PowerPoint_

ラボ構成のデータ構造。

```
interface IConfiguration extends Core.IUserData
```


## <a name="properties"></a>プロパティ


|||
|:-----|:-----|
| `appVersion: Core.IVersion`|Version of the application associated with this configuration.|
| `components: Core.IComponent[]`|Components included with the lab.|
| `name: string`|The name of the lab.|
| `timeline: Core.ITimelineConfiguration`|The timeline configuration for the lab.|
| `analytics: Core.IAnalyticsConfiguration`|The analytics configuration for the lab.|
