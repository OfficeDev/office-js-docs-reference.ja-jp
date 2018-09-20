# <a name="scopes-element"></a><span data-ttu-id="6301f-101">Scopes 要素</span><span class="sxs-lookup"><span data-stu-id="6301f-101">Scopes element</span></span>

<span data-ttu-id="6301f-102">アドインで必要な Microsoft Graph に対するアクセス許可が含まれます。</span><span class="sxs-lookup"><span data-stu-id="6301f-102">Contains permissions to Microsoft Graph that the add-in needs.</span></span> <span data-ttu-id="6301f-103">Office ストアは、Scopes 要素を使用して同意ダイアログ ボックスを作成します。</span><span class="sxs-lookup"><span data-stu-id="6301f-103">The Office Store uses the Scopes element to create a consent dialog box.</span></span> <span data-ttu-id="6301f-104">ユーザーがストアからアドインをインストールすると、ユーザーの Microsoft Graph のデータに対する指定されたアクセス許可をアドインに付与するように要求されます。</span><span class="sxs-lookup"><span data-stu-id="6301f-104">When users install the add-in from the Store, they are prompted to grant the add-in the specified permissions to the user's Microsoft Graph data.</span></span>

## <a name="child-elements"></a><span data-ttu-id="6301f-105">子要素</span><span class="sxs-lookup"><span data-stu-id="6301f-105">Child elements</span></span>

|  <span data-ttu-id="6301f-106">要素</span><span class="sxs-lookup"><span data-stu-id="6301f-106">Element</span></span> |  <span data-ttu-id="6301f-107">支払期日</span><span class="sxs-lookup"><span data-stu-id="6301f-107">Type</span></span>  |  <span data-ttu-id="6301f-108">説明</span><span class="sxs-lookup"><span data-stu-id="6301f-108">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="6301f-109">**Scope**</span><span class="sxs-lookup"><span data-stu-id="6301f-109">**Scope**</span></span>                |  <span data-ttu-id="6301f-110">string</span><span class="sxs-lookup"><span data-stu-id="6301f-110">string</span></span>     |   <span data-ttu-id="6301f-111">Microsoft Graph に対するアクセス許可の名前。たとえば、Files.Read.All です。</span><span class="sxs-lookup"><span data-stu-id="6301f-111">The name of a permission to Microsoft Graph; for example, Files.Read.All.</span></span> |

## <a name="example"></a><span data-ttu-id="6301f-112">例</span><span class="sxs-lookup"><span data-stu-id="6301f-112">Example</span></span>

```xml
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    ...
    <WebApplicationInfo>
      <Id>12345678-abcd-1234-efab-123456789abc</Id>
      <Resource>api://myDomain.com/12345678-abcd-1234-efab-123456789abc<Resource>
      <Scopes>
        <Scope>Files.Read.All</Scope>
        <Scope>offline_access</Scope>
        <Scope>openid</Scope>
        <Scope>profile</Scope>
      </Scopes>
    </WebApplicationInfo>
  </VersionOverrides>
...
</OfficeApp>
```
