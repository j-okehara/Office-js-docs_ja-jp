#<a name="ui.dialog-object"></a>UI.Dialog オブジェクト
[displayDialogAsync](officeui.displaydialogasync.md) メソッドが呼び出されたときに返されるオブジェクト。

## <a name="members"></a>メンバー
| メンバー	       | 型	   |説明|
|:---------------|:--------|:----------|
|close|機能|アドインでダイアログ ボックスを閉じることができます。|
|addEventHandler|機能|イベント ハンドラーを登録します。サポートされているイベントは次の 2 つです。 <ul><li>DialogMessageReceived。ダイアログ ボックスがメッセージを親に送信すると発生します。</li><li>DialogEventReceived。ダイアログ ボックスが閉じられたとき、またはアンロードされたときに発生します。</li></ul> |


### <a name="close()"></a>close()
対応するダイアログ ボックスを閉じるために親ページから呼び出されます。     
```js    
[dialogObject].close();    
``` 

#### <a name="parameters"></a>パラメーター    
なし。 

#### <a name="returns"></a>戻り値    
void  


#### <a name="examples"></a>例
例については、「[DisplayDialogAsync メソッド](officeui.displaydialogasync.md)」トピックを参照してください。
