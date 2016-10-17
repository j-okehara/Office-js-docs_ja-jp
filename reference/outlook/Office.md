 

# <a name="office"></a>Office

Office 名前空間は、すべての Office アプリケーションのアドインで使用される共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office 名前空間の完全な一覧については、「[共有 API](../shared/shared-api.md)」を参照してください。

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](./tutorial-api-requirement-sets.md)| 1.0|
|適用可能な Outlook のモード| 作成または読み取り|

### <a name="namespaces"></a>名前空間

[context](Office.context.md):Outlook アドイン API で使用するために、Office アドイン API のコンテキストの名前空間から共有インターフェイスを提供します。

[MailboxEnums](Office.MailboxEnums.md):ItemType、EntityType、AttachmentType、RecipientType、ResponseType、および ItemNotificationMessageType 列挙型が含まれます。

### <a name="members"></a>メンバー

####  <a name="asyncresultstatus-:string"></a>AsyncResultStatus :String

非同期呼び出しの結果を指定します。

##### <a name="type:"></a>型:

*   String

##### <a name="properties:"></a>プロパティ:

|名前| 型| 説明|
|---|---|---|
|`Succeeded`| String|呼び出しが成功しました。|
|`Failed`| String|呼び出しが失敗しました。|

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](./tutorial-api-requirement-sets.md)| 1.0|
|適用可能な Outlook のモード| 作成または読み取り|
####  <a name="coerciontype-:string"></a>CoercionType :String

呼び出されたメソッドによって返される、または設定されるデータを強制的に変換する方法を指定します。

##### <a name="type:"></a>型:

*   String

##### <a name="properties:"></a>プロパティ:

|名前| 型| 説明|
|---|---|---|
|`Html`| String|HTML 形式で返されるデータを要求します。|
|`Text`| String|テキスト形式で返されるデータを要求します。|

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](./tutorial-api-requirement-sets.md)| 1.0|
|適用可能な Outlook のモード| 作成または読み取り|
####  <a name="sourceproperty-:string"></a>SourceProperty :String

呼び出されたメソッドによって返されるデータのソースを指定します。

##### <a name="type:"></a>型:

*   String

##### <a name="properties:"></a>プロパティ:

|名前| 型| 説明|
|---|---|---|
|`Body`| String|データのソースは、メッセージの本文です。|
|`Subject`| String|データのソースは、メッセージの件名です。|

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](./tutorial-api-requirement-sets.md)| 1.0|
|適用可能な Outlook のモード| 作成または読み取り|
