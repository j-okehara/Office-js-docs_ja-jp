
# <a name="labs.valueholder"></a>Labs.ValueHolder

 _**適用対象:** Office 用アプリ | Office アドイン | Office Mix | PowerPoint_

指定したラボの値を保持および追跡するコンテナー オブジェクトです。値は、ローカルまたはサーバーのいずれにも格納される場合があります。

```
class ValueHolder<T>
```


## <a name="variables"></a>変数


|||
|:-----|:-----|
| `public var isHint: boolean`|**True** if the value is a hint.|
| `public var hasBeenRequested: boolean`|**True** if the value has been requested by the lab.|
| `public var hasValue: boolean`|**True** if the value container currently has the desired value.|
| `public var value: T`|The value that is held in the container.|
| `public var id: string`|The ID of the value.|

## <a name="methods"></a>メソッド




### <a name="getvalue"></a>getValue

 `public function getValue(callback: Labs.Core.ILabCallback<T>): void`

指定された値を取得します。

 **Parameters**


|||
|:-----|:-----|
| _callback_|指定した値を返すコールバック関数です。|

### <a name="providevalue"></a>provideValue

 `public function provideValue(value: T): void`

値コンテナーに値を提供する内部メソッド。

 **Parameters**


|||
|:-----|:-----|
| _value_|値コンテナーに提供できる値。|
