# <a name="functionfile-element"></a>FunctionFile 要素

UI を表示する代わりに JavaScript 関数を実行するアドイン コマンドによってアドインが公開する操作のソース コード ファイルを指定します。**FunctionFile** 要素は、[FormFactor](./formfactor.md) の子要素です。**FunctionFile** 要素の **resid** 属性は、HTML ファイルの URL を含む **Resources** 要素内の **Url** 要素の **id** 属性値に設定されます。この HTML ファイルには、[Control 要素](control.md)の定義に従い、UI なしのアドイン コマンド ボタンに使用されるすべての JavaScript 関数が含まれるか、読み込まれます。

**FunctionFile** 要素の例を次に示します。


```XML
<DesktopFormFactor>
          <FunctionFile resid="residDesktopFuncUrl" />
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <!-- information about this extension point -->
          </ExtensionPoint>

          <!-- You can define more than one ExtensionPoint element as needed -->

        </DesktopFormFactor>
```

**FunctionFile** 要素で示される HTML ファイルの JavaScript は `Office.initialize` を呼び出して、1 つのパラメーター (`event`) を取る名前付きの関数を定義する必要があります。この関数は、[item.notificationMessages](../../reference/outlook/Office.context.mailbox.item.md) API を使用して、ユーザーに進行状況、成功、失敗を示す必要があります。また、実行が終了したときに [event.completed](../../reference/shared/event.completed.md) を呼び出す必要もあります。関数の名前は、UI なしボタンの **FunctionName** 要素で使用されます。

**trackMessage** 関数を定義する HTML ファイルの例を次に示します。

```js
Office.intialize = function () {
    doAuth();
}

function trackMessage (event) {
    var buttonId = event.source.id;    
    var itemId = Office.context.mailbox.item.id;
    // save this message
    event.completed();
}
```

次のコードは、**FunctionName** で使用される関数の実装方法を示しています。




```js
        // The initialize function must be run each time a new page is loaded.
        (function () {
            Office.initialize = function (reason) {
               // If you need to initialize something you can do so here.
            };
        })();

            // Your function must be in the global namespace.
        function writeText(event) {

            // Implement your custom code here. The following code is a simple example.

            Office.context.document.setSelectedDataAsync("ExecuteFunction works. Button ID=" + event.source.id,
                function (asyncResult) {
                    var error = asyncResult.error;
                    if (asyncResult.status === "failed") {
                        // Show error message.
                    }
                    else {
                        // Show success message.
                    }
                });
           // Calling event.completed is required. event.completed lets the platform know that processing has completed.
       event.completed();
        }
```


 >**重要**  **event.completed** に対する呼び出しにより、イベントが正常に処理されたことが通知されます。同一のアドイン コマンドを複数回クリックするなど、関数を複数回呼び出すと、すべてのイベントが自動的にキューに入れられます。最初のイベントが自動的に実行され、その他のイベントはキューに残ります。関数により **event.completed** が呼び出されると、キューに入れられている、その関数に対する次の呼び出しが実行されます。**event.completed** を実装する必要があります。実装しない場合、関数は実行されません。
