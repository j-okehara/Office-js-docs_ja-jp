
# <a name="labs.componentinstance"></a>Labs.ComponentInstance

 _**適用対象:** Office 用アプリ | Office アドイン | Office Mix | PowerPoint_

実行時にユーザーに指定されるコンポーネントのインスタンス化である、コンポーネントのインスタンスを表します。オブジェクトには、ラボの特定の実行のためのコンポーネントの翻訳されたビューが含まれます。

```
class ComponentInstance<T> extends Labs.ComponentInstanceBase
```


## <a name="properties"></a>プロパティ

なし。


## <a name="methods"></a>Methods




### <a name="constructor"></a>コンストラクター

 `function constructor()`

Initializes a new instance of the  **ComponentInstance** class.


### <a name="createattempt"></a>createAttempt

 `public function createAttempt(callback: Labs.Core.ILabCallback<T>): void`

コンポーネントのテキスト内で新しい試行を作成します。

 **パラメーター**


|**名前**|**説明**|
|:-----|:-----|
| _callback_|試行が作成されたときに起動するコールバック。|

### <a name="getattempts"></a>getAttempts

 `public function getAttempts(callback: Labs.Core.ILabCallback<T[]>): void`

指定のコンポーネントに関連するすべての試行を取得します。

 **パラメーター**


|**名前**|**説明**|
|:-----|:-----|
| _callback_|試行が取得されると起動するコールバック。|

### <a name="getcreateattemptoptions"></a>getCreateAttemptOptions

 `public function getCreateAttemptOptions(): Labs.Core.Actions.ICreateAttemptOptions`

既定の作成試行オプションを取得します。派生クラスによってオーバーライドできます。


### <a name="buildattempt"></a>buildAttempt

 `public function buildAttempt(createAttemptResult: Labs.Core.IAction): T`

特定のアクションからの試行を構築します。派生クラスで実装される必要があります。

 **パラメーター**


|**名前**|**説明**|
|:-----|:-----|
| _createAttemptResult_|指定した試行の試行作成アクション。|
