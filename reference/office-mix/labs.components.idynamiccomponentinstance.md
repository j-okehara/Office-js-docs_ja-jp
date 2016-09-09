﻿
# Labs.Components.IDynamicComponentInstance

 _**適用対象:** Office 用アプリ | Office アドイン | Office Mix | PowerPoint_

動的コンポーネントのインスタンス。

```
interface IDynamicComponentInstance extends Labs.Core.IComponentInstance
```


## プロパティ


|名前|説明|
|:-----|:-----|
| `generatedComponentTypes: string[]`|この動的コンポーネントが生成する場合があるコンポーネントの種類を含む配列です。|
| `maxComponents: number`|この動的コンポーネントによって生成されるコンポーネントの最大数です。または、上限がない場合は、**Labs.Components.Infinite** です。|