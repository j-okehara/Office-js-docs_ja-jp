
# Labs.ComponentInstance

 _**適用対象:** Office 用アプリ | Office アドイン | Office Mix | PowerPoint_

実行時にユーザーに指定されるコンポーネントのインスタンス化である、コンポーネントのインスタンスを表します。オブジェクトには、ラボの特定の実行のためのコンポーネントの翻訳されたビューが含まれます。

```
class ComponentInstance<T> extends Labs.ComponentInstanceBase
```


## プロパティ

なし。


## Methods




### コンストラクター

 `function constructor()`

Initializes a new instance of the  **ComponentInstance** class.


### createAttempt

 `public function createAttempt(callback: Labs.Core.ILabCallback<T>): void`

Creates a new attempt in the context of a component.

 **パラメーター**


|**名前**|**説明**|
|:-----|:-----|
| _callback_|Callback fired when the attempt has been created.|

### getAttempts

 `public function getAttempts(callback: Labs.Core.ILabCallback<T[]>): void`

Retrieves all attempts associated with the given component.

 **パラメーター**


|**名前**|**説明**|
|:-----|:-----|
| _callback_|Callback fired when the attempts have been retrieved.|

### getCreateAttemptOptions

 `public function getCreateAttemptOptions(): Labs.Core.Actions.ICreateAttemptOptions`

既定の作成試行オプションを取得します。 派生クラスによってオーバーライドできます。


### buildAttempt

 `public function buildAttempt(createAttemptResult: Labs.Core.IAction): T`

特定のアクションからの試行を構築します。 派生クラスで実装される必要があります。

 **パラメーター**


|**名前**|**説明**|
|:-----|:-----|
| _createAttemptResult_|The create attempt action for the specified attempt.|
