
# <a name="labs.core.iconnectionresponse"></a>Labs.Core.IConnectionResponse

 _**適用対象:** Office 用アプリ | Office アドイン | Office Mix | PowerPoint_

接続の呼び出しから返される応答情報。

```
interface IConnectionResponse
```


## <a name="properties"></a>プロパティ


|||
|:-----|:-----|
| `initializationInfo: Core.IConfigurationInfo`|Initialization configureation information, or  **null** if the app has not been initialized.|
| `mode: Core.LabMode`|The mode which the lab is currently running in.|
| `hostVersion: Core.IVersion`|Version information ([Labs.Core.IVersion](../../reference/office-mix/labs.core.iversion.md)) for the server.|
| `userInfo: Core.IUserInfo`|Information about the user ([Labs.Core.IUserInfo](../../reference/office-mix/labs.core.iuserinfo.md)).|
