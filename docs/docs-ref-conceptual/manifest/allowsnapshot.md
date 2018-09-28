# <a name="allowsnapshot-element"></a><span data-ttu-id="7a927-101">AllowSnapshot 要素</span><span class="sxs-lookup"><span data-stu-id="7a927-101">AllowSnapshot element</span></span>

<span data-ttu-id="7a927-102">ホスト ドキュメントと共にコンテンツ アドインのスナップショット イメージを保存するかどうかを指定します。</span><span class="sxs-lookup"><span data-stu-id="7a927-102">Specifies whether a snapshot image of your content add-in is saved with the host document.</span></span>

<span data-ttu-id="7a927-103">**アドインの種類:** コンテンツ</span><span class="sxs-lookup"><span data-stu-id="7a927-103">**Add-in type:** Content</span></span>

## <a name="syntax"></a><span data-ttu-id="7a927-104">構文</span><span class="sxs-lookup"><span data-stu-id="7a927-104">Syntax</span></span>

```XML
<AllowSnapshot> [true | false]</AllowSnapshot>
```

## <a name="contained-in"></a><span data-ttu-id="7a927-105">含まれています。</span><span class="sxs-lookup"><span data-stu-id="7a927-105">Contained in</span></span>

[<span data-ttu-id="7a927-106">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="7a927-106">OfficeApp</span></span>](officeapp.md)

## <a name="remarks"></a><span data-ttu-id="7a927-107">備考</span><span class="sxs-lookup"><span data-stu-id="7a927-107">Remarks</span></span>

 > [!IMPORTANT]
 > <span data-ttu-id="7a927-108">**AllowSnapshot**は、`true`既定します。</span><span class="sxs-lookup"><span data-stu-id="7a927-108">**AllowSnapshot** is `true` by default.</span></span> <span data-ttu-id="7a927-109">この場合、Office アドインをサポートしていないバージョンのホスト アプリケーションでドキュメントを開くユーザーがアドインのイメージを表示できるようになったり、ホスト アプリケーションがアドインをホストするサーバーに接続できない場合にアドインの静的イメージが提供されたりします。</span><span class="sxs-lookup"><span data-stu-id="7a927-109">This makes an image of the add-in visible for users that open the document in a version of the host application that doesn't support Office Add-ins, or provides a static image of the add-in if the host application can't connect to the server hosting the add-in.</span></span> <span data-ttu-id="7a927-110">しかしこれは、アドインがホストされるドキュメントから、アドインに表示される機密性の高い情報に直接アクセスできるということでもあります。</span><span class="sxs-lookup"><span data-stu-id="7a927-110">However, this also means that potentially sensitive information displayed in the add-in can be accessed directly from the document hosting the add-in.</span></span>

