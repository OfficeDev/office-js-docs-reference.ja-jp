# <a name="method-element"></a><span data-ttu-id="bfa40-101">Method 要素</span><span class="sxs-lookup"><span data-stu-id="bfa40-101">Method element</span></span>

<span data-ttu-id="bfa40-102">Office アドインをアクティブにするために必要な JavaScript API for Office の個別のメソッドを指定します。</span><span class="sxs-lookup"><span data-stu-id="bfa40-102">Specifies an individual method from the JavaScript API for Office that your Office Add-in requires in order to activate.</span></span>

<span data-ttu-id="bfa40-103">**アドインの種類:** コンテンツ、作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="bfa40-103">**Add-in type:** Content, Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="bfa40-104">構文</span><span class="sxs-lookup"><span data-stu-id="bfa40-104">Syntax</span></span>

```XML
<Method Name="string"/>
```

## <a name="contained-in"></a><span data-ttu-id="bfa40-105">含まれています。</span><span class="sxs-lookup"><span data-stu-id="bfa40-105">Contained in</span></span>

[<span data-ttu-id="bfa40-106">メソッド</span><span class="sxs-lookup"><span data-stu-id="bfa40-106">Methods</span></span>](methods.md)

## <a name="attributes"></a><span data-ttu-id="bfa40-107">属性</span><span class="sxs-lookup"><span data-stu-id="bfa40-107">Attributes</span></span>

|<span data-ttu-id="bfa40-108">**属性**</span><span class="sxs-lookup"><span data-stu-id="bfa40-108">**Attribute**</span></span>|<span data-ttu-id="bfa40-109">**型**</span><span class="sxs-lookup"><span data-stu-id="bfa40-109">**Type**</span></span>|<span data-ttu-id="bfa40-110">**必須**</span><span class="sxs-lookup"><span data-stu-id="bfa40-110">**Required**</span></span>|<span data-ttu-id="bfa40-111">**説明**</span><span class="sxs-lookup"><span data-stu-id="bfa40-111">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="bfa40-112">名前</span><span class="sxs-lookup"><span data-stu-id="bfa40-112">Name</span></span>|<span data-ttu-id="bfa40-113">文字列</span><span class="sxs-lookup"><span data-stu-id="bfa40-113">string</span></span>|<span data-ttu-id="bfa40-114">必須</span><span class="sxs-lookup"><span data-stu-id="bfa40-114">required</span></span>|<span data-ttu-id="bfa40-p101">必要なメソッドの名前をその親オブジェクトで修飾して指定します。たとえば、**getSelectedDataAsync** メソッドを指定するには、`"Document.getSelectedDataAsync"` と指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="bfa40-p101">Specifies the name of the required method qualified with its parent object. For example, to specify the  **getSelectedDataAsync** method, you must specify `"Document.getSelectedDataAsync"`.</span></span>|

## <a name="remarks"></a><span data-ttu-id="bfa40-117">備考</span><span class="sxs-lookup"><span data-stu-id="bfa40-117">Remarks</span></span>

<span data-ttu-id="bfa40-118">メールのアドインでは、**メソッド**および**メソッド**の要素はサポートされていません。要求セットの詳細については、 [Office のバージョンおよび要件の設定](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="bfa40-118">The  **Methods** and **Method** elements aren't supported by mail add-ins. For more information about requirement sets, see [Office versions and requirement sets](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="bfa40-119">個々 のメソッドの最小バージョン要件を指定する方法がないため、メソッドが、実行時に使用可能であることを確認する必要がありますもを使用する**if**ステートメントの追加のスクリプトでそのメソッドを呼び出すときにします。</span><span class="sxs-lookup"><span data-stu-id="bfa40-119">Because there is no way to specify the minimum version requirement for individual methods, to make sure that a method is available at runtime, you should also use an **if** statement when calling that method in the script of your add-in.</span></span> <span data-ttu-id="bfa40-120">これを行う方法の詳細については、 [Office 用の JavaScript API を理解する](https://docs.microsoft.com/office/dev/add-ins/develop/understanding-the-javascript-api-for-office)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="bfa40-120">For more information about how to do this, see [Understanding the JavaScript API for Office](https://docs.microsoft.com/office/dev/add-ins/develop/understanding-the-javascript-api-for-office).</span></span>

