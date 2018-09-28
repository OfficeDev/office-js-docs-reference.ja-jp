# <a name="functionfile-element"></a><span data-ttu-id="89e69-101">FunctionFile 要素</span><span class="sxs-lookup"><span data-stu-id="89e69-101">FunctionFile element</span></span>

<span data-ttu-id="89e69-p101">UI を表示する代わりに JavaScript 関数を実行するアドイン コマンドによってアドインが公開する操作の、ソース コード ファイルを指定します。**FunctionFile** 要素は、[DesktopFormFactor](desktopformfactor.md) または [MobileFormFactor](mobileformfactor.md) の子要素です。**FunctionFile** 要素の **resid** 属性は、HTML ファイルの URL を含む **Resources** 要素内の **Url** 要素の **id** 属性値に設定されます。この HTML ファイルには、[Control 要素](control.md)の定義に従い、UI なしのアドイン コマンド ボタンに使用されるすべての JavaScript 関数が含まれるか、読み込まれます。</span><span class="sxs-lookup"><span data-stu-id="89e69-p101">Specifies the source code file for operations that an add-in exposes through add-in commands that execute a JavaScript function instead of displaying UI. The  **FunctionFile** element is a child element of [DesktopFormFactor](desktopformfactor.md) or [MobileFormFactor](mobileformfactor.md). The **resid** attribute of the **FunctionFile** element is set to the value of the **id** attribute of a **Url** element in the **Resources** element that contains the URL to an HTML file that contains or loads all  the JavaScript functions used by UI-less add-in command buttons, as defined by the [Control element](control.md).</span></span>

<span data-ttu-id="89e69-105">次は、 **FunctionFile**要素の例です。</span><span class="sxs-lookup"><span data-stu-id="89e69-105">The following is an example of the  **FunctionFile** element.</span></span>

```XML
<DesktopFormFactor>
  <FunctionFile resid="residDesktopFuncUrl" />
  <ExtensionPoint xsi:type="PrimaryCommandSurface">
    <!-- information about this extension point -->
  </ExtensionPoint>

  <!-- You can define more than one ExtensionPoint element as needed -->

</DesktopFormFactor>
```

<span data-ttu-id="89e69-106">**FunctionFile**要素で指定された HTML ファイル内の JavaScript を呼び出す必要があります`Office.initialize`を 1 つのパラメーターを受け取るという名前の関数を定義して: `event`。</span><span class="sxs-lookup"><span data-stu-id="89e69-106">The JavaScript in the HTML file indicated by the  **FunctionFile** element must call `Office.initialize` and define named functions that take a single parameter: `event`.</span></span> <span data-ttu-id="89e69-107">関数を使用する必要があります、`item.notificationMessages`の進行状況、成功、または失敗をユーザーに示すために API です。</span><span class="sxs-lookup"><span data-stu-id="89e69-107">The functions should use the `item.notificationMessages` API to indicate progress, success, or failure to the user.</span></span> <span data-ttu-id="89e69-108">呼び出す必要もあります`event.completed`の終了時に実行します。</span><span class="sxs-lookup"><span data-stu-id="89e69-108">It should also call `event.completed` when it has finished execution.</span></span> <span data-ttu-id="89e69-109">関数の名前は、省略ボタンの場合、**関数名**の要素で使用されます。</span><span class="sxs-lookup"><span data-stu-id="89e69-109">The name of the functions are used in the **FunctionName** element for UI-less buttons.</span></span>

<span data-ttu-id="89e69-110">**trackMessage** 関数を定義する HTML ファイルの例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="89e69-110">The following is an example of an HTML file defining a **trackMessage** function.</span></span>

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

<span data-ttu-id="89e69-111">次のコードは、**関数名**で使用する関数を実装する方法を示します。</span><span class="sxs-lookup"><span data-stu-id="89e69-111">The following code shows how to implement the function used by  **FunctionName**.</span></span>

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
> <span data-ttu-id="89e69-112">**Event.completed**への呼び出しは、イベントが正常に処理されたことを通知します。</span><span class="sxs-lookup"><span data-stu-id="89e69-112">The call to  **event.completed** signals that you have successfully handled the event.</span></span> <span data-ttu-id="89e69-113">同一のアドイン コマンドを複数回クリックするなどの方法で関数を複数回呼び出すと、すべてのイベントは自動的にキューに入れられます。</span><span class="sxs-lookup"><span data-stu-id="89e69-113">When a function is called multiple times, such as multiple clicks on the same add-in command, all events are automatically queued.</span></span> <span data-ttu-id="89e69-114">最初のイベントが自動的に実行され、その他のイベントはキューに残ります。</span><span class="sxs-lookup"><span data-stu-id="89e69-114">The first event runs automatically, while the other events remain on the queue.</span></span> <span data-ttu-id="89e69-115">関数により **event.completed** が呼び出されると、キューに入れられている、その関数に対する次の呼び出しが実行されます。</span><span class="sxs-lookup"><span data-stu-id="89e69-115">When your function calls **event.completed**, the next queued call to that function runs.</span></span> <span data-ttu-id="89e69-116">**Event.completed**; を呼び出す必要があります。それ以外の場合、関数は実行されません。</span><span class="sxs-lookup"><span data-stu-id="89e69-116">You must call **event.completed**; otherwise your function will not run.</span></span>