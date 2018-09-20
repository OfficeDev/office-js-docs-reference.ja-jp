# <a name="action-element"></a><span data-ttu-id="b12aa-101">Action 要素</span><span class="sxs-lookup"><span data-stu-id="b12aa-101">Action element</span></span>

<span data-ttu-id="b12aa-102">ユーザーが[ボタン](control.md#button-control)または[メニュー](control.md#menu-dropdown-button-controls) コントロールを選択したときに実行する操作を指定します。</span><span class="sxs-lookup"><span data-stu-id="b12aa-102">Specifies the action to perform when the user selects a  [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls) controls.</span></span>
 
## <a name="attributes"></a><span data-ttu-id="b12aa-103">属性</span><span class="sxs-lookup"><span data-stu-id="b12aa-103">Attributes</span></span>

|  <span data-ttu-id="b12aa-104">属性</span><span class="sxs-lookup"><span data-stu-id="b12aa-104">Attribute</span></span>  |  <span data-ttu-id="b12aa-105">必須</span><span class="sxs-lookup"><span data-stu-id="b12aa-105">Required</span></span>  |  <span data-ttu-id="b12aa-106">説明</span><span class="sxs-lookup"><span data-stu-id="b12aa-106">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="b12aa-107">xsi:type</span><span class="sxs-lookup"><span data-stu-id="b12aa-107">xsi:type</span></span>](#xsitype)  |  <span data-ttu-id="b12aa-108">はい</span><span class="sxs-lookup"><span data-stu-id="b12aa-108">Yes</span></span>  | <span data-ttu-id="b12aa-109">実行する操作の種類</span><span class="sxs-lookup"><span data-stu-id="b12aa-109">Action type to take</span></span>|

## <a name="child-elements"></a><span data-ttu-id="b12aa-110">子要素</span><span class="sxs-lookup"><span data-stu-id="b12aa-110">Child elements</span></span>

|  <span data-ttu-id="b12aa-111">要素</span><span class="sxs-lookup"><span data-stu-id="b12aa-111">Element</span></span> |  <span data-ttu-id="b12aa-112">説明</span><span class="sxs-lookup"><span data-stu-id="b12aa-112">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="b12aa-113">FunctionName</span><span class="sxs-lookup"><span data-stu-id="b12aa-113">FunctionName</span></span>](#functionname) |    <span data-ttu-id="b12aa-114">実行する関数の名前を指定します。</span><span class="sxs-lookup"><span data-stu-id="b12aa-114">Specifies the name of the function to execute.</span></span> |
|  [<span data-ttu-id="b12aa-115">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="b12aa-115">SourceLocation</span></span>](#sourcelocation) |    <span data-ttu-id="b12aa-116">この操作のソース ファイルの場所を指定します。</span><span class="sxs-lookup"><span data-stu-id="b12aa-116">Specifies the source file location for this action.</span></span> |
|  [<span data-ttu-id="b12aa-117">TaskpaneId</span><span class="sxs-lookup"><span data-stu-id="b12aa-117">TaskpaneId</span></span>](#taskpaneid) | <span data-ttu-id="b12aa-118">作業ウィンドウ コンテナーの ID を指定します。</span><span class="sxs-lookup"><span data-stu-id="b12aa-118">Specifies the ID of the task pane container.</span></span>|
|  [<span data-ttu-id="b12aa-119">Title</span><span class="sxs-lookup"><span data-stu-id="b12aa-119">Title</span></span>](#title) | <span data-ttu-id="b12aa-120">作業ウィンドウのカスタム タイトルを指定します。</span><span class="sxs-lookup"><span data-stu-id="b12aa-120">Specifies the custom title for the task pane.</span></span>|
|  [<span data-ttu-id="b12aa-121">SupportsPinning</span><span class="sxs-lookup"><span data-stu-id="b12aa-121">SupportsPinning</span></span>](#supportspinning) | <span data-ttu-id="b12aa-122">作業ウィンドウがピン留めをサポートすることを指定します。これにより、ユーザーが選択を変更したときも作業ウィンドウが開いたままになります。</span><span class="sxs-lookup"><span data-stu-id="b12aa-122">Specifies that a task pane supports pinning, which keeps the task pane open when the user changes the selection.</span></span>|
  

## <a name="xsitype"></a><span data-ttu-id="b12aa-123">xsi:type</span><span class="sxs-lookup"><span data-stu-id="b12aa-123">xsi:type</span></span>

<span data-ttu-id="b12aa-p101">この属性は、ユーザーがボタンをクリックしたときに実行される操作の種類を指定します。次のいずれかを指定できます。</span><span class="sxs-lookup"><span data-stu-id="b12aa-p101">This attribute specifies the kind of action performed when the user selects the button. It can be one of the following:</span></span>

- `ExecuteFunction`
- `ShowTaskpane`

## <a name="functionname"></a><span data-ttu-id="b12aa-126">FunctionName</span><span class="sxs-lookup"><span data-stu-id="b12aa-126">FunctionName</span></span>

<span data-ttu-id="b12aa-p102">**xsi:type** が "ExecuteFunction" のときに必ず指定する要素です。実行する関数の名前を指定します。関数は、[FunctionFile](functionfile.md) 要素に指定されたファイルに含まれています。</span><span class="sxs-lookup"><span data-stu-id="b12aa-p102">Required element when **xsi:type** is "ExecuteFunction". Specifies the name of the function to execute. The function is contained in the file specified in the [FunctionFile](functionfile.md) element.</span></span>

```xml
<Action xsi:type="ExecuteFunction">
  <FunctionName>getSubject</FunctionName>
</Action>
```

## <a name="sourcelocation"></a><span data-ttu-id="b12aa-130">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="b12aa-130">SourceLocation</span></span>

<span data-ttu-id="b12aa-p103">**xsi:type** が "ShowTaskpane" のときに必ず指定する要素です。このアクションのソース ファイルの場所を指定します。 **resid** 属性は、 **Resources** 要素の **Urls** 要素にある **Url** 要素の [id](resources.md) 属性の値を指定します。</span><span class="sxs-lookup"><span data-stu-id="b12aa-p103">Required element when  **xsi:type** is "ShowTaskpane". Specifies the source file location for this action. The **resid** attribute must be set to the value of the **id** attribute of a **Url** element in the **Urls** element in the [Resources](resources.md) element.</span></span>

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
</Action>
```  

## <a name="taskpaneid"></a><span data-ttu-id="b12aa-134">TaskpaneId</span><span class="sxs-lookup"><span data-stu-id="b12aa-134">TaskpaneId</span></span>

<span data-ttu-id="b12aa-p104">**xsi:type** が "ShowTaskpane" の場合に省略可能な要素。作業ウィンドウ コンテナーの ID を指定します。複数の "ShowTaskpane" の操作があり、それぞれに対して独立したウィンドウを開く場合は、異なる **TaskpaneId** を使用します。同じウィンドウを共有する異なる操作に対しては、同じ **TaskpaneId** を使用します。ユーザーが同じ **TaskpaneId** を共有するコマンドを選択した場合、ウィンドウ コンテナーは開いたままですが、ウィンドウのコンテンツは対応する操作の "SourceLocation" に置き換えられます。</span><span class="sxs-lookup"><span data-stu-id="b12aa-p104">Optional element when  **xsi:type** is "ShowTaskpane". Specifies the ID of the task pane container. When you have multiple "ShowTaskpane" actions, use a different **TaskpaneId** if you want an independent pane for each. Use the same **TaskpaneId** for  different actions that share the same pane. When users choose commands that share the same **TaskpaneId**, the pane container will remain open but the contents of the pane will be replaced with the corresponding Action "SourceLocation".</span></span> 

> [!NOTE]
> <span data-ttu-id="b12aa-140">Outlook では、この要素はサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="b12aa-140">This element is not supported in Outlook.</span></span>

<span data-ttu-id="b12aa-141">次の例では、同じ **TaskpaneId** を共有する 2 つのアクションを示します。</span><span class="sxs-lookup"><span data-stu-id="b12aa-141">The following example shows two actions that share the same **TaskpaneId**.</span></span> 

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

<span data-ttu-id="b12aa-p105">次の例では、異なる **TaskpaneId** を使用する 2 つの操作を示します。これらの例を全体的な流れで確認する場合は、「[Simple Add-in Commands Sample](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/blob/master/Simple/Manifest/SimpleAddin.xml)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="b12aa-p105">The following examples show two actions that use a different **TaskpaneId**. To see these examples in context, see [Simple Add-in Commands Sample](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/blob/master/Simple/Manifest/SimpleAddin.xml).</span></span>

```xml
<Action xsi:type="ShowTaskpane">
   <TaskpaneId>MyTaskPaneID1</TaskpaneId>
   <SourceLocation resid="Contoso.Taskpane1.Url" />
</Action>

<Action xsi:type="ShowTaskpane">
   <TaskpaneId>MyTaskPaneID2</TaskpaneId>
   <SourceLocation resid="Contoso.Taskpane2.Url" />
</Action>
```  

```xml
<bt:Urls>
   <bt:Url id="Contoso.Taskpane1.Url" DefaultValue="https://commandsimple.azurewebsites.net/Taskpane.html" />
   <bt:Url id="Contoso.Taskpane2.Url" DefaultValue="https://commandsimple.azurewebsites.net/Taskpane2.html" />
</bt:Urls>
```  

## <a name="title"></a><span data-ttu-id="b12aa-144">Title</span><span class="sxs-lookup"><span data-stu-id="b12aa-144">Title</span></span>
<span data-ttu-id="b12aa-p106">**xsi:type** が "ShowTaskpane" の場合に省略可能な要素。この操作に関する、作業ウィンドウのカスタム タイトルを指定します。</span><span class="sxs-lookup"><span data-stu-id="b12aa-p106">Optional element when  **xsi:type** is "ShowTaskpane". Specifies the custom title for the task pane for this action.</span></span> 

<span data-ttu-id="b12aa-147">以下の例は、**Title** 要素を使用する 2 つの異なるアクションを示します。</span><span class="sxs-lookup"><span data-stu-id="b12aa-147">The following examples show two different actions that use the **Title** element.</span></span>

```xml
<Action xsi:type="ShowTaskpane">
<TaskpaneId>Office.AutoShowTaskpaneWithDocument</TaskpaneId>
<SourceLocation resid="PG.Code.Url" />
<Title resid="PG.CodeCommand.Title" />
</Action>
``` 

```xml
<Action xsi:type="ShowTaskpane">
<SourceLocation resid="PG.Run.Url" />
<Title resid="PG.RunCommand.Title" />
</Action>
``` 

```xml
<bt:Urls>
<bt:Url id="PG.Code.Url" DefaultValue="https://localhost:3000?commands=1" />
<bt:Url id="PG.Run.Url" DefaultValue="https://localhost:3000/run.html" />
</bt:Urls>
``` 

```xml
<bt:ShortStrings>
<bt:String id="PG.CodeCommand.Title" DefaultValue="Code" />
<bt:String id="PG.RunCommand.Title" DefaultValue="Run" />
</bt:ShortStrings>
``` 

## <a name="supportspinning"></a><span data-ttu-id="b12aa-148">SupportsPinning</span><span class="sxs-lookup"><span data-stu-id="b12aa-148">SupportsPinning</span></span>

<span data-ttu-id="b12aa-p107">**xsi:type** が "ShowTaskpane" の場合に省略可能な要素。これを収容している [VersionOverrides](versionoverrides.md) 要素は、`xsi:type` 属性の値が `VersionOverridesV1_1` になっている必要があります。ピン留めをサポートする場合は、この要素に `true` の値を含めます。ユーザーは、作業ウィンドウをピン留めできるようになります。ピン留めすると、選択を変更したときも作業ウィンドウが開いたままになります。詳細については、「[Outlook にピン留め可能な作業ウィンドウを実装する](https://docs.microsoft.com/outlook/add-ins/pinnable-taskpane)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="b12aa-p107">Optional element when **xsi:type** is "ShowTaskpane". The containing [VersionOverrides](versionoverrides.md) elements must have an `xsi:type` attribute value of `VersionOverridesV1_1`. Include this element with a value of `true` to support taskpane pinning. The user will be able to "pin" the taskpane, causing it to stay open when changing the selection. For more information, see [Implement a pinnable taskpane in Outlook](https://docs.microsoft.com/outlook/add-ins/pinnable-taskpane).</span></span>

> [!NOTE]
> <span data-ttu-id="b12aa-154">SupportsPinning だけで現在サポートされて Outlook 2016 (ビルド 7628.1000 以降) のウィンドウです。</span><span class="sxs-lookup"><span data-stu-id="b12aa-154">SupportsPinning currently only supported by Outlook 2016 for Windows (build 7628.1000 or later).</span></span>

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
  <SupportsPinning>true</SupportsPinning>
</Action>
```


