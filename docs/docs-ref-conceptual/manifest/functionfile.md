# <a name="functionfile-element"></a>FunctionFile 要素

UI を表示する代わりに JavaScript 関数を実行するアドイン コマンドによってアドインが公開する操作の、ソース コード ファイルを指定します。**FunctionFile** 要素は、[DesktopFormFactor](desktopformfactor.md) または [MobileFormFactor](mobileformfactor.md) の子要素です。**FunctionFile** 要素の **resid** 属性は、HTML ファイルの URL を含む **Resources** 要素内の **Url** 要素の **id** 属性値に設定されます。この HTML ファイルには、[Control 要素](control.md)の定義に従い、UI なしのアドイン コマンド ボタンに使用されるすべての JavaScript 関数が含まれるか、読み込まれます。

次は、 **FunctionFile**要素の例です。

```XML
<DesktopFormFactor>
  <FunctionFile resid="residDesktopFuncUrl" />
  <ExtensionPoint xsi:type="PrimaryCommandSurface">
    <!-- information about this extension point -->
  </ExtensionPoint>

  <!-- You can define more than one ExtensionPoint element as needed -->

</DesktopFormFactor>
```

**FunctionFile**要素で指定された HTML ファイル内の JavaScript を呼び出す必要があります`Office.initialize`を 1 つのパラメーターを受け取るという名前の関数を定義して: `event`。 関数を使用する必要があります、`item.notificationMessages`の進行状況、成功、または失敗をユーザーに示すために API です。 呼び出す必要もあります`event.completed`の終了時に実行します。 関数の名前は、省略ボタンの場合、**関数名**の要素で使用されます。

**trackMessage** 関数を定義する HTML ファイルの例を次に示します。

```js
Office.initialize = function () {
    doAuth();
}

function trackMessage (event) {
    var buttonId = event.source.id;    
    var itemId = Office.context.mailbox.item.id;
    // save this message
    event.completed();
}
```

次のコードは、**関数名**で使用する関数を実装する方法を示します。

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

> [!IMPORTANT]
> **Event.completed**への呼び出しは、イベントが正常に処理されたことを通知します。 同一のアドイン コマンドを複数回クリックするなどの方法で関数を複数回呼び出すと、すべてのイベントは自動的にキューに入れられます。 最初のイベントが自動的に実行され、その他のイベントはキューに残ります。 関数により **event.completed** が呼び出されると、キューに入れられている、その関数に対する次の呼び出しが実行されます。 **Event.completed**; を呼び出す必要があります。それ以外の場合、関数は実行されません。