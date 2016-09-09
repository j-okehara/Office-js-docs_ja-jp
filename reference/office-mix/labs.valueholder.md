
# Labs.ValueHolder

 _**適用対象:** Office 用アプリ | Office アドイン | Office Mix | PowerPoint_

指定したラボの値を保持および追跡するコンテナー オブジェクトです。値は、ローカルまたはサーバーのいずれにも格納される場合があります。

```
class ValueHolder<T>
```


## 変数


|||
|:-----|:-----|
| `public var isHint: boolean`|**True** if the value is a hint.|
| `public var hasBeenRequested: boolean`|**True** if the value has been requested by the lab.|
| `public var hasValue: boolean`|**True** if the value container currently has the desired value.|
| `public var value: T`|The value that is held in the container.|
| `public var id: string`|The ID of the value.|

## メソッド




### getValue

 `public function getValue(callback: Labs.Core.ILabCallback<T>): void`

Retrieves the specified value.

 **パラメーター**


|||
|:-----|:-----|
| _callback_|Callback function that returns the specified value.|

### provideValue

 `public function provideValue(value: T): void`

Internal method that provides the value to the value container.

 **パラメーター**


|||
|:-----|:-----|
| _value_|The value to provide to the value container.|
