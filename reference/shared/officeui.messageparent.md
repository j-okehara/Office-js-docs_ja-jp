# UI.messageParent メソッド

メッセージをダイアログ ボックスからその親/オープナー ページに配信します。 この API を呼び出すページは、親と同じドメインにある必要があります。 

## 構文

```js
Office.context.ui.messageParent("Message from Dialog box");
```

## パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|messageObject|String または Boolean|ダイアログ ボックスからメッセージを受け付け、アドインに配信します。|

## 戻り値
void

## 例
例については、「[DisplayDialogAsync メソッド](officeui.displaydialogasync.md)」トピックを参照してください。

