# <a name="ui.messageparent-method"></a>UI.messageParent メソッド

メッセージをダイアログ ボックスからその親/オープナー ページに配信します。この API を呼び出すページは、親と同じドメインにある必要があります。 

## <a name="syntax"></a>構文

```js
Office.context.ui.messageParent("Message from Dialog box");
```

## <a name="parameters"></a>パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|messageObject|String または Boolean|ダイアログ ボックスからメッセージを受け付け、アドインに配信します。|

## <a name="returns"></a>戻り値
void

## <a name="examples"></a>例
例については、「[DisplayDialogAsync メソッド](officeui.displaydialogasync.md)」トピックを参照してください。

