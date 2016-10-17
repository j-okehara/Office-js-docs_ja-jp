
# <a name="labs.core.itimelineconfiguration"></a>Labs.Core.ITimelineConfiguration

 _**適用対象:** Office 用アプリ | Office アドイン | Office Mix | PowerPoint_

[Labs.Timeline](../../reference/office-mix/labs.timeline.md) の構成オプション。一連のタイムライン構成オプションを指定するのを許可します。

```
interface ITimelineConfiguration
```


## <a name="properties"></a>プロパティ


|||
|:-----|:-----|
| `duration: number`|The duration of the lab, in seconds.|
| `capabilities: string[]`|An array list of timeline capabilities that the lab supports, for example, play, pause, seek, and so forth.|
