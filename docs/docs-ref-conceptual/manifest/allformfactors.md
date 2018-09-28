# <a name="allformfactors-element"></a><span data-ttu-id="0b106-101">AllFormFactors 要素</span><span class="sxs-lookup"><span data-stu-id="0b106-101">AllFormFactors element</span></span>

<span data-ttu-id="0b106-102">すべてのフォーム ファクターについてアドインの設定を指定します。</span><span class="sxs-lookup"><span data-stu-id="0b106-102">Specifies the settings for an add-in for all form factors.</span></span> <span data-ttu-id="0b106-103">現在、 **AllFormFactors**を使用する唯一の機能は、ユーザー定義関数です。</span><span class="sxs-lookup"><span data-stu-id="0b106-103">Currently, the only feature using **AllFormFactors** is custom functions.</span></span> <span data-ttu-id="0b106-104">**AllFormFactors**は、ユーザー定義関数を使用する場合に必要な要素です。</span><span class="sxs-lookup"><span data-stu-id="0b106-104">**AllFormFactors** is a required element when using custom functions.</span></span>

## <a name="child-elements"></a><span data-ttu-id="0b106-105">子要素</span><span class="sxs-lookup"><span data-stu-id="0b106-105">Child elements</span></span>

|  <span data-ttu-id="0b106-106">要素</span><span class="sxs-lookup"><span data-stu-id="0b106-106">Element</span></span> |  <span data-ttu-id="0b106-107">必須</span><span class="sxs-lookup"><span data-stu-id="0b106-107">Required</span></span>  |  <span data-ttu-id="0b106-108">説明</span><span class="sxs-lookup"><span data-stu-id="0b106-108">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="0b106-109">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="0b106-109">ExtensionPoint</span></span>](extensionpoint.md) |  <span data-ttu-id="0b106-110">はい</span><span class="sxs-lookup"><span data-stu-id="0b106-110">Yes</span></span> |  <span data-ttu-id="0b106-111">アドインが機能を公開する場所を定義します。</span><span class="sxs-lookup"><span data-stu-id="0b106-111">Defines where an add-in exposes functionality.</span></span> |

## <a name="allformfactors-example"></a><span data-ttu-id="0b106-112">AllFormFactors の例</span><span class="sxs-lookup"><span data-stu-id="0b106-112">AllFormFactors example</span></span>

```xml
<Hosts>
    <Host xsi:type="Workbook">
        <AllFormFactors>
            <ExtensionPoint xsi:type="CustomFunctions">
                    <!-- Information on this extension point -->
            </ExtensionPoint>
        </AllFormFactors>
    </Host>
</Hosts>
```
