# <a name="action-element"></a>Action 要素
ユーザーが[ボタン](./control.md#button-control)または[メニュー](./control.md#menu-dropdown-button-controls) コントロールを選択したときに実行する操作を指定します。
 
## <a name="attributes"></a>属性

|  属性  |  必須  |  説明  |
|:-----|:-----|:-----|
|  [xsi:type](#xsitype)  |  はい  | 実行する操作の種類|


## <a name="child-elements"></a>子要素

|  要素 |  説明  |
|:-----|:-----|
|  [FunctionName](#functionname) |    実行する関数の名前を指定します。 |
|  [SourceLocation](#sourcelocation) |    この操作のソース ファイルの場所を指定します。 |
|  [TaskpaneId](#taskpaneid) | 作業ウィンドウ コンテナーの ID を指定します。|
|  [SupportsPinning](#supportspinning) | 作業ウィンドウがピン留めをサポートすることを指定します。これにより、ユーザーが選択を変更したときも作業ウィンドウが開いたままになります。|
  

## <a name="xsitype"></a>xsi:type
この属性は、ユーザーがボタンをクリックしたときに実行される操作の種類を指定します。次のいずれかを指定できます。

- `ExecuteFunction`
- `ShowTaskpane`

## <a name="functionname"></a>FunctionName

**xsi:type** が "ExecuteFunction" のときに必ず指定する要素です。実行する関数の名前を指定します。関数は、[FunctionFile](./functionfile.md) 要素に指定されたファイルに含まれています。

```xml
<Action xsi:type="ExecuteFunction">
  <FunctionName>getSubject</FunctionName>
</Action>
```

## <a name="sourcelocation"></a>SourceLocation
**xsi:type** が "ShowTaskpane" のときに必ず指定する要素です。このアクションのソース ファイルの場所を指定します。 **resid** 属性は、 **Resources** 要素の **Urls** 要素にある [Url](./resources.md#urls) 要素の [id](./resources.md) 属性の値を指定します。

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
</Action>
```  

## <a name="taskpaneid"></a>TaskpaneId
**xsi:type** が "ShowTaskpane" の場合に省略可能な要素。作業ウィンドウ コンテナーの ID を指定します。複数の "ShowTaskpane" の操作があり、それぞれに対して独立したウィンドウを開く場合は、異なる **TaskpaneId** を使用します。同じウィンドウを共有する異なる操作に対しては、同じ **TaskpaneId** を使用します。ユーザーが同じ **TaskpaneId** を共有するコマンドを選択した場合、ウィンドウ コンテナーは開いたままですが、ウィンドウのコンテンツは対応する操作の "SourceLocation" に置き換えられます。 

>**注:**この要素は、Outlook ではサポートされていません。

次の例では、同じ **TaskpaneId** を共有する 2 つの操作を示します。 


```xml
<Action xsi:type="ShowTaskpane">
  <TaskpaneId>MyPane</TaskpaneId>
  <SourceLocation resid="aTaskPaneUrl" />
</Action>

<Action xsi:type="ShowTaskpane">
  <TaskpaneId>MyPane</TaskpaneId>
  <SourceLocation resid="anotherTaskPaneUrl" />
</Action>
```  

## <a name="supportspinning"></a>SupportsPinning

**xsi:type** が "ShowTaskpane" の場合に省略可能な要素。これを収容している [VersionOverrides](./versionoverrides.md) 要素は、`xsi:type` 属性の値が `VersionOverridesV1_1` になっている必要があります。ピン留めをサポートする場合は、この要素に `true` の値を含めます。ユーザーは、作業ウィンドウをピン留めできるようになります。ピン留めすると、選択を変更したときも作業ウィンドウが開いたままになります。詳細については、「[Outlook にピン留め可能な作業ウィンドウを実装する](../../docs/outlook/manifests/pinnable-taskpane)」を参照してください。

> **注**:現時点で、SupportsPinning は Outlook 2016 for Windows (ビルド 7628.1000 以降) でのみサポートされます。

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
  <SupportsPinning>true</SupportsPinning>
</Action>
```