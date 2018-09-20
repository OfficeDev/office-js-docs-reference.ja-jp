# <a name="permissions-element"></a><span data-ttu-id="21a47-101">Permissions 要素</span><span class="sxs-lookup"><span data-stu-id="21a47-101">Permissions element</span></span>

<span data-ttu-id="21a47-102">Office アドインの API アクセスのレベルを指定します。最小特権の原則に基づいてアクセス許可を要求する必要があります。</span><span class="sxs-lookup"><span data-stu-id="21a47-102">Specifies the level of API access for your Office Add-in; you should request permissions based on the principle of least privilege.</span></span>

<span data-ttu-id="21a47-103">**アドインの種類:** コンテンツ、作業ウィンドウ、メール</span><span class="sxs-lookup"><span data-stu-id="21a47-103">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="21a47-104">構文</span><span class="sxs-lookup"><span data-stu-id="21a47-104">Syntax</span></span>

<span data-ttu-id="21a47-105">コンテンツ アドインおよび作業ウィンドウ アドインの場合:</span><span class="sxs-lookup"><span data-stu-id="21a47-105">For content and task pane add-ins:</span></span>

```XML
 <Permissions> [Restricted | ReadDocument | ReadAllDocument | WriteDocument | ReadWriteDocument]</Permissions>
```

<span data-ttu-id="21a47-106">メールのアドイン</span><span class="sxs-lookup"><span data-stu-id="21a47-106">For mail add-ins</span></span>

```XML
 <Permissions>[Restricted | ReadItem | ReadWriteItem | ReadWriteMailbox]</Permissions>
```

## <a name="contained-in"></a><span data-ttu-id="21a47-107">含まれています。</span><span class="sxs-lookup"><span data-stu-id="21a47-107">Contained in</span></span>

[<span data-ttu-id="21a47-108">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="21a47-108">OfficeApp</span></span>](officeapp.md)

## <a name="remarks"></a><span data-ttu-id="21a47-109">備考</span><span class="sxs-lookup"><span data-stu-id="21a47-109">Remarks</span></span>

<span data-ttu-id="21a47-110">詳細については、「[コンテンツ アドインおよび作業ウィンドウ アドインでの API 使用のアクセス許可を要求する](https://docs.microsoft.com/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins)」と「[Outlook アドインのアクセス許可について](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="21a47-110">For more detail, see [Requesting permissions for API use in content and task pane add-ins](https://docs.microsoft.com/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins) and [Understanding Outlook add-in permissions](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>
