# <a name="word-javascript-api-overview"></a><span data-ttu-id="b9b40-101">単語の JavaScript API の概要</span><span class="sxs-lookup"><span data-stu-id="b9b40-101">Word JavaScript API overview</span></span>

<span data-ttu-id="b9b40-p101">Word には、ドキュメント コンテンツおよびメタデータとデータをやり取りするアドインを作成するために使用できる豊富な API のセットが用意されています。これらの API を使用して、Word を統合および拡張する魅力的なエクスペリエンスを作成します。コンテンツのインポートとエクスポート、別のデータ ソースから新しいドキュメントのアセンブル、カスタムのドキュメント ソリューションを作成するドキュメント ワークフローとの統合を行えます。</span><span class="sxs-lookup"><span data-stu-id="b9b40-p101">Word provides a rich set of APIs that you can use to create add-ins that interact with document content and metadata. Use these APIs to create compelling experiences that integrate with and extend Word. You can import and export content, assemble new documents from different data sources, and integrate with document workflows to create custom document solutions.</span></span>

<span data-ttu-id="b9b40-105">2 つの JavaScript API を使用して、Word 文書のオブジェクトおよびメタデータと対話できます。</span><span class="sxs-lookup"><span data-stu-id="b9b40-105">You can use two JavaScript APIs to interact with the objects and metadata in a Word document:</span></span>

- <span data-ttu-id="b9b40-106">Word JavaScript API - Office 2016 で導入。</span><span class="sxs-lookup"><span data-stu-id="b9b40-106">Word JavaScript API - Introduced in Office 2016.</span></span>
- <span data-ttu-id="b9b40-107">[JavaScript API for Office](../javascript-api-for-office.md) (Office.js) - Office 2013 で導入。</span><span class="sxs-lookup"><span data-stu-id="b9b40-107">[JavaScript API for Office](../javascript-api-for-office.md) (Office.js) - Introduced in Office 2013.</span></span>

## <a name="word-javascript-api"></a><span data-ttu-id="b9b40-108">Word JavaScript API</span><span class="sxs-lookup"><span data-stu-id="b9b40-108">Word JavaScript API</span></span>

<span data-ttu-id="b9b40-p102">Word JavaScript API は Office.js によって読み込まれます。Word JavaScript API では、ドキュメントや段落などのオブジェクトとの対話方法が変わります。Word JavaScript API は、これらのそれぞれのオブジェクトの取得や更新をする個々の非同期の API を提供するのではなく、Word で実行されている実際のオブジェクトに対応する JavaScript の “プロキシ” オブジェクトを提供します。プロキシ オブジェクトのプロパティの読み取りと書き込みを同期的に行い、プロキシ オブジェクトに操作を実行する同期メソッドを呼び出すことによって、それらのプロキシ オブジェクトを操作することができます。プロキシ オブジェクトに対するこうした操作は実行中のスクリプトですぐには認識されません。**context.sync** メソッドは、キューに入れられた命令を実行し、また読み込まれた Word オブジェクトのプロパティをスクリプトで使用するために取得することで、実行中の JavaScript オブジェクトと Office の実際のオブジェクトとの間で状態を同期します。</span><span class="sxs-lookup"><span data-stu-id="b9b40-p102">The Word JavaScript API is loaded by Office.js. The Word JavaScript API changes the way that you can interact with objects like documents and paragraphs. Rather than providing individual asynchronous APIs for retrieving and updating each of these objects, the Word JavaScript API provides “proxy” JavaScript objects that correspond to the real objects running in Word. You can interact with these proxy objects by synchronously reading and writing their properties and calling synchronous methods to perform operations on them. These interactions with proxy objects aren’t immediately realized in the running script. The **context.sync** method synchronizes the state between your running JavaScript and the real objects in Office by executing queued instructions and retrieving properties of loaded Word objects for use in your script.</span></span>

## <a name="javascript-api-for-office"></a><span data-ttu-id="b9b40-115">JavaScript API for Office</span><span class="sxs-lookup"><span data-stu-id="b9b40-115">JavaScript API for Office</span></span>

<span data-ttu-id="b9b40-116">Office.js は、次の場所から参照できます。</span><span class="sxs-lookup"><span data-stu-id="b9b40-116">You can reference Office.js from the following locations:</span></span>

* <span data-ttu-id="b9b40-117">https://appsforoffice.microsoft.com/lib/1/hosted/office.js-生産のアドインの場合、このリソースを使用します。</span><span class="sxs-lookup"><span data-stu-id="b9b40-117">https://appsforoffice.microsoft.com/lib/1/hosted/office.js - use this resource for production add-ins.</span></span>
* <span data-ttu-id="b9b40-118">https://appsforoffice.microsoft.com/lib/beta/hosted/office.js-プレビュー機能を開こうとしているときにこのリソースを使用します。</span><span class="sxs-lookup"><span data-stu-id="b9b40-118">https://appsforoffice.microsoft.com/lib/beta/hosted/office.js - use this resource when you're trying out preview features.</span></span>

<span data-ttu-id="b9b40-p103">[Visual Studio](https://www.visualstudio.com/products/free-developer-offers-vs) を使用している場合、[Office Developer Tools](https://www.visualstudio.com/features/office-tools-vs.aspx) をダウンロードして、Office.js を含むプロジェクト テンプレートを取得できます。[nuget から Office.js を取得する](https://www.nuget.org/packages/Microsoft.Office.js/)こともできます。</span><span class="sxs-lookup"><span data-stu-id="b9b40-p103">If you're using [Visual Studio](https://www.visualstudio.com/products/free-developer-offers-vs), you can download the [Office Developer Tools](https://www.visualstudio.com/features/office-tools-vs.aspx) to get project templates that include Office.js.  You can also use [nuget to get Office.js](https://www.nuget.org/packages/Microsoft.Office.js/).</span></span>

<span data-ttu-id="b9b40-121">TypeScript を使用していて npm がある場合、コマンド ライン インターフェイスにこれを入力すると、TypeScript の定義を取得できます: `typings install office-js --ambient`。</span><span class="sxs-lookup"><span data-stu-id="b9b40-121">If you use TypeScript and have npm, you can get the the TypeScript definitions by typing this in your command line interface: `typings install office-js --ambient`.</span></span>

## <a name="running-word-add-ins"></a><span data-ttu-id="b9b40-122">Word アドインを実行します</span><span class="sxs-lookup"><span data-stu-id="b9b40-122">Running Word add-ins</span></span>

<span data-ttu-id="b9b40-p104">アドインを実行するには、Office.initialize イベント ハンドラーを使用します。アドインの初期化の詳細については、「[API について](https://docs.microsoft.com/office/dev/add-ins/develop/understanding-the-javascript-api-for-office)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="b9b40-p104">To run your add-in, use an Office.initialize event handler. For more information about add-in initialization, see [Understanding the API](https://docs.microsoft.com/office/dev/add-ins/develop/understanding-the-javascript-api-for-office) .</span></span>

<span data-ttu-id="b9b40-125">アドインに Word 2016 のターゲットまたは後で、 **Word.run()** メソッドに関数を渡すことにより実行します。</span><span class="sxs-lookup"><span data-stu-id="b9b40-125">Add-ins that target Word 2016 or later execute by passing a function into the **Word.run()** method.</span></span> <span data-ttu-id="b9b40-126">**Run**メソッドに渡される関数は、コンテキストの引数が必要です。</span><span class="sxs-lookup"><span data-stu-id="b9b40-126">The function passed into the **run** method must have a context argument.</span></span> <span data-ttu-id="b9b40-127">この[コンテキスト オブジェクト](/javascript/api/word/word.requestcontext)は、Office のオブジェクトから取得するコンテキスト オブジェクトとは異なるが、Word の実行時環境との対話にも使用します。</span><span class="sxs-lookup"><span data-stu-id="b9b40-127">This [context object](/javascript/api/word/word.requestcontext) is different than the context object you get from the Office object, but it is also used to interact with the Word runtime environment.</span></span> <span data-ttu-id="b9b40-128">コンテキスト オブジェクトでは、JavaScript API の Word オブジェクト モデルへのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="b9b40-128">The context object provides access to the Word JavaScript API object model.</span></span> <span data-ttu-id="b9b40-129">次の例では、 **Word.run()** メソッドを使用して、アドインを初期化し、Word を実行する方法を示します。</span><span class="sxs-lookup"><span data-stu-id="b9b40-129">The following example shows how to initialize and execute a Word add-in by using the **Word.run()** method.</span></span>

```js
(function () {
    "use strict";

    // The initialize event handler must be run on each page to initialize Office JS.
    // You can add optional custom initialization code that will run after OfficeJS
    // has initialized.
    Office.initialize = function (reason) {
        // The reason object tells how the add-in was initialized. The values can be:
        // inserted - the add-in was inserted to an open document.
        // documentOpened - the add-in was already inserted in to the document and the document was opened.

        // Checks for the DOM to load using the jQuery ready function.
        $(document).ready(function () {
            // Set your optional initialization code.
            // You can also load saved settings from the Office object.
        });
    };

    // Run a batch operation against the Word JavaScript API object model.
    // Use the context argument to get access to the Word document.
    Word.run(function (context) {

        // Create a proxy object for the document.
        var thisDocument = context.document;
        // ...
    })
})();
```

### <a name="synchronizing-word-documents-with-word-javascript-api-proxy-objects"></a><span data-ttu-id="b9b40-130">Word 文書を Word JavaScript API のプロキシ オブジェクトと同期します</span><span class="sxs-lookup"><span data-stu-id="b9b40-130">Synchronizing Word documents with Word JavaScript API proxy objects</span></span>

<span data-ttu-id="b9b40-p106">Word JavaScript API オブジェクト モデルは、Word 内のオブジェクトと緩く結合されています。Word JavaScript API のオブジェクトは、Word 文書内のオブジェクトのプロキシです。プロキシ オブジェクトで実行されたアクションは、ドキュメントの状態が同期されるまで、Word では認識されません。逆に、Word 文書の状態は、ドキュメントの状態が同期されるまでプロキシ オブジェクトでは認識されません。ドキュメントの状態を同期するには、**context.sync()** メソッドを実行します。次の例では、本文のプロキシ オブジェクトと、その本文プロキシ オブジェクトにテキスト プロパティを読み込むためのキューに登録済みのコマンドを作成し、さらに **context.sync()** メソッドを使用してWord 文書内の本文と本文プロキシ オブジェクトとを同期します。</span><span class="sxs-lookup"><span data-stu-id="b9b40-p106">The Word JavaScript API object model is loosely coupled with the objects in Word. Word JavaScript API objects are proxies for objects in a Word document. Actions taken on proxy objects are not realized in Word until the document state has been synchronized. Conversely, the state of the Word document is not realized in the proxy objects until the document state has been synchronized. To synchronize the document state, you run the **context.sync()** method. The following example creates a proxy body object and a queued command to load the text property on the proxy body object, and uses the **context.sync()** method to synchronize the body of the Word document with the body proxy object.</span></span>

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    // The body object hasn't been set with any property values.
    var body = context.document.body;

    // Queue a command to load the text property for the proxy document body object.
    context.load(body, 'text');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log("Body contents: " + body.text);
    });
})
```

### <a name="executing-a-batch-of-commands"></a><span data-ttu-id="b9b40-137">コマンドのバッチを実行する</span><span class="sxs-lookup"><span data-stu-id="b9b40-137">Executing a batch of commands</span></span>

<span data-ttu-id="b9b40-p107">Word のプロキシ オブジェクトには、オブジェクト モデルにアクセスして更新するためのメソッドが用意されています。これらのメソッドは、バッチでキューに入れられた順序で順番に実行されます。context.sync() 呼び出しが行われると、キューに入れられたすべてのコマンドが実行されます。</span><span class="sxs-lookup"><span data-stu-id="b9b40-p107">The Word proxy objects have methods for accessing and updating the object model. These methods are executed sequentially in the order in which they were queued in the batch. All of the commands that are queued in the batch are executed when context.sync() is called.</span></span>

<span data-ttu-id="b9b40-p108">次の例では、コマンドのキューが機能する仕組みを示しています。**context.sync()** が呼び出されると、本文テキストを読み込むコマンドが Word で実行されます。次に、Word の本文にテキストを挿入するコマンドが生成されます。その結果は本文のプロキシ オブジェクトに返されます。Word JavaScript API の **body.text** プロパティの値は、テキストが Word 文書に挿入される<u>前</u>の Word 文書本文の値です。</span><span class="sxs-lookup"><span data-stu-id="b9b40-p108">The following example shows how the command queue works. When **context.sync()** is called, the command to load the body text is executed in Word. Then, the command to insert text into the body in Word occurs. The results are then returned to the body proxy object. The value of the **body.text** property in the Word JavaScript API is the value of the Word document body <u>before</u> the text was inserted into Word document.</span></span>


```js
// Run a batch operation against the Word JavaScript API.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a command to load the text property of the proxy body object.
    context.load(body, 'text');

    // Queue a command to insert text into the end of the Word document body.
    body.insertText('This is text inserted after loading the body.text property',
                    Word.InsertLocation.end);

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log("Body contents: " + body.text);
    });
})
```

## <a name="word-javascript-api-open-specifications"></a><span data-ttu-id="b9b40-146">単語の JavaScript API 仕様を開く</span><span class="sxs-lookup"><span data-stu-id="b9b40-146">Word JavaScript API open specifications</span></span>

<span data-ttu-id="b9b40-p109">新しい Word アドイン用の API の設計と開発にあたり、[Open API の仕様](../openspec.md) ページでこれらに対するフィードバックの提供が可能になります。Word JavaScript API 用のパイプラインの新機能をご確認いただき、設計の仕様に関する情報をお寄せください。</span><span class="sxs-lookup"><span data-stu-id="b9b40-p109">As we design and develop new APIs for Word add-ins, we'll make them available for your feedback on our [Open API specifications](../openspec.md) page. Find out what new features are in the pipeline for the Word JavaScript APIs, and provide your input on our design specifications.</span></span>

## <a name="word-javascript-api-reference"></a><span data-ttu-id="b9b40-149">Word JavaScript API リファレンス</span><span class="sxs-lookup"><span data-stu-id="b9b40-149">Word JavaScript API reference</span></span>

<span data-ttu-id="b9b40-150">単語の JavaScript API の詳細については、 [Word の JavaScript API リファレンス ドキュメント](/javascript/api/word)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="b9b40-150">For detailed information about the Word JavaScript API, see the [Word JavaScript API reference documentation](/javascript/api/word).</span></span>

## <a name="see-also"></a><span data-ttu-id="b9b40-151">関連項目</span><span class="sxs-lookup"><span data-stu-id="b9b40-151">See also</span></span>

* [<span data-ttu-id="b9b40-152">Word アドインの概要</span><span class="sxs-lookup"><span data-stu-id="b9b40-152">Word add-ins overview</span></span>](https://docs.microsoft.com/office/dev/add-ins/word/word-add-ins-programming-overview)
* [<span data-ttu-id="b9b40-153">Office アドイン プラットフォームの概要</span><span class="sxs-lookup"><span data-stu-id="b9b40-153">Office Add-ins platform overview</span></span>](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-ins)
* [<span data-ttu-id="b9b40-154">GitHub の Word アドインのサンプル</span><span class="sxs-lookup"><span data-stu-id="b9b40-154">Word add-in samples on GitHub</span></span>](https://github.com/OfficeDev?utf8=%E2%9C%93&q=Word)
