# <a name="extensionpoint-element"></a><span data-ttu-id="07a9c-101">ExtensionPoint 要素</span><span class="sxs-lookup"><span data-stu-id="07a9c-101">ExtensionPoint element</span></span>

 <span data-ttu-id="07a9c-102">Office UI でアドインが機能を公開する場所を定義します。</span><span class="sxs-lookup"><span data-stu-id="07a9c-102">Defines where an add-in exposes functionality in the Office UI.</span></span> <span data-ttu-id="07a9c-103">**ExtensionPoint** 要素は、[AllFormFactors](allformfactors.md)、[DesktopFormFactor](desktopformfactor.md)、[MobileFormFactor](mobileformfactor.md) の子要素です。</span><span class="sxs-lookup"><span data-stu-id="07a9c-103">The **ExtensionPoint** element is a child element of [AllFormFactors](allformfactors.md), [DesktopFormFactor](desktopformfactor.md) or [MobileFormFactor](mobileformfactor.md).</span></span> 

## <a name="attributes"></a><span data-ttu-id="07a9c-104">属性</span><span class="sxs-lookup"><span data-stu-id="07a9c-104">Attributes</span></span>

|  <span data-ttu-id="07a9c-105">属性</span><span class="sxs-lookup"><span data-stu-id="07a9c-105">Attribute</span></span>  |  <span data-ttu-id="07a9c-106">必須</span><span class="sxs-lookup"><span data-stu-id="07a9c-106">Required</span></span>  |  <span data-ttu-id="07a9c-107">説明</span><span class="sxs-lookup"><span data-stu-id="07a9c-107">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="07a9c-108">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="07a9c-108">**xsi:type**</span></span>  |  <span data-ttu-id="07a9c-109">はい</span><span class="sxs-lookup"><span data-stu-id="07a9c-109">Yes</span></span>  | <span data-ttu-id="07a9c-110">定義される拡張点の種類。</span><span class="sxs-lookup"><span data-stu-id="07a9c-110">The type of extension point being defined.</span></span>|

## <a name="extension-points-for-excel-only"></a><span data-ttu-id="07a9c-111">Excel のみの拡張点</span><span class="sxs-lookup"><span data-stu-id="07a9c-111">Extension points for Excel only</span></span>

- <span data-ttu-id="07a9c-112">**CustomFunctions** - Excel の JavaScript で記述されたユーザー定義関数です。</span><span class="sxs-lookup"><span data-stu-id="07a9c-112">**CustomFunctions** - A custom function written in JavaScript for Excel.</span></span>

<span data-ttu-id="07a9c-113">[この XML コードのサンプル](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/customfunctions.xml)では、 **CustomFunctions**属性の値と子要素を使用する**ExtensionPoint**要素を使用する方法を示します。</span><span class="sxs-lookup"><span data-stu-id="07a9c-113">[This XML code sample](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/customfunctions.xml) shows how to use the **ExtensionPoint** element with the **CustomFunctions** attribute value, and the child elements to be used.</span></span>

## <a name="extension-points-for-word-excel-powerpoint-and-onenote-add-in-commands"></a><span data-ttu-id="07a9c-114">Word、Excel、PowerPoint、OneNote アドイン コマンドの拡張点</span><span class="sxs-lookup"><span data-stu-id="07a9c-114">Extension points for Word, Excel, PowerPoint, and OneNote add-in commands</span></span>

- <span data-ttu-id="07a9c-115">**PrimaryCommandSurface** - Office のリボン。</span><span class="sxs-lookup"><span data-stu-id="07a9c-115">**PrimaryCommandSurface** - The ribbon in Office.</span></span>
- <span data-ttu-id="07a9c-116">**ContextMenu**Office UI で右クリックしたときに表示されるショートカット メニュー。</span><span class="sxs-lookup"><span data-stu-id="07a9c-116">**ContextMenu** - The shortcut menu that appears when you right-click in the Office UI.</span></span>

<span data-ttu-id="07a9c-117">次の例は、 **PrimaryCommandSurface** と **ContextMenu** の属性値を持つ **ExtensionPoint** 要素を使用する方法と、各要素と併用する必要がある子要素を示しています。</span><span class="sxs-lookup"><span data-stu-id="07a9c-117">The following examples show how to use the  **ExtensionPoint** element with **PrimaryCommandSurface** and **ContextMenu** attribute values, and the child elements that should be used with each.</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="07a9c-118">ID 属性が含まれている要素の一意の ID を提供するを確認します</span><span class="sxs-lookup"><span data-stu-id="07a9c-118">For elements that contain an ID attribute, make sure you provide a unique ID.</span></span> <span data-ttu-id="07a9c-119">あなたの ID と共に、会社の名前を使用することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="07a9c-119">We recommend that you use your company's name along with your ID.</span></span> <span data-ttu-id="07a9c-120">たとえば、次の形式にします。</span><span class="sxs-lookup"><span data-stu-id="07a9c-120">For example, use the following format.</span></span> <CustomTab id="mycompanyname.mygroupname">

```XML
<ExtensionPoint xsi:type="PrimaryCommandSurface">
          <CustomTab id="Contoso Tab">
          <!-- If you want to use a default tab that comes with Office, remove the above CustomTab element, and then uncomment the following OfficeTab element -->
            <!-- <OfficeTab id="TabData"> -->
            <Label resid="residLabel4" />
            <Group id="Group1Id12">
              <Label resid="residLabel4" />
              <Icon>
                <bt:Image size="16" resid="icon1_32x32" />
                <bt:Image size="32" resid="icon1_32x32" />
                <bt:Image size="80" resid="icon1_32x32" />
              </Icon>
              <Tooltip resid="residToolTip" />
              <Control xsi:type="Button" id="Button1Id1">

                  <!-- information about the control -->
              </Control>
              <!-- other controls, as needed -->
            </Group>
          </CustomTab>
        </ExtensionPoint>

      <ExtensionPoint xsi:type="ContextMenu">
        <OfficeMenu id="ContextMenuCell">
          <Control xsi:type="Menu" id="ContextMenu2">
                  <!-- information about the control -->
          </Control>
          <!-- other controls, as needed -->
        </OfficeMenu>
        </ExtensionPoint>
```

#### <a name="child-elements"></a><span data-ttu-id="07a9c-121">子要素</span><span class="sxs-lookup"><span data-stu-id="07a9c-121">Child elements</span></span>
 
|<span data-ttu-id="07a9c-122">**要素**</span><span class="sxs-lookup"><span data-stu-id="07a9c-122">**Element**</span></span>|<span data-ttu-id="07a9c-123">**説明**</span><span class="sxs-lookup"><span data-stu-id="07a9c-123">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="07a9c-124">**CustomTab**</span><span class="sxs-lookup"><span data-stu-id="07a9c-124">**CustomTab**</span></span>|<span data-ttu-id="07a9c-p103">カスタム タブをリボンに追加する場合は必須です (  **PrimaryCommandSurface** を使用)。 **CustomTab** 要素を使用する場合、 **OfficeTab** 要素は使用できません。 **id** 属性が必要です。</span><span class="sxs-lookup"><span data-stu-id="07a9c-p103">Required if you want to add a custom tab to the ribbon (using  **PrimaryCommandSurface**). If you use the  **CustomTab** element, you can't use the **OfficeTab** element. The **id** attribute is required.</span></span>|
|<span data-ttu-id="07a9c-128">**OfficeTab**</span><span class="sxs-lookup"><span data-stu-id="07a9c-128">**OfficeTab**</span></span>|<span data-ttu-id="07a9c-p104">既定の Office リボン タブを拡張する場合は必須です (**PrimaryCommandSurface** を使用)。**OfficeTab** 要素を使用する場合、**CustomTab** 要素は使用できません。詳細については、「[OfficeTab](officetab.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="07a9c-p104">Required if you want to extend a default Office ribbon tab (using **PrimaryCommandSurface**). If you use the  **OfficeTab** element, you can't use the **CustomTab** element. For details, see [OfficeTab](officetab.md).</span></span>|
|<span data-ttu-id="07a9c-132">**OfficeMenu**</span><span class="sxs-lookup"><span data-stu-id="07a9c-132">**OfficeMenu**</span></span>|<span data-ttu-id="07a9c-p105">既定のコンテキスト メニューにアドイン コマンドを追加する場合は必須です (**ContextMenu** を使用)。**id** 属性は以下に設定する必要があります。 </span><span class="sxs-lookup"><span data-stu-id="07a9c-p105">Required if you're adding add-in commands to a default context menu (using  **ContextMenu**). The  **id** attribute must be set to: </span></span><br/> <span data-ttu-id="07a9c-p106">Excel または Word の場合は - **ContextMenuText**。テキストが選択され、ユーザーが選択されたテキストを右クリックしたときに、コンテキスト メニューに項目が表示されます。 </span><span class="sxs-lookup"><span data-stu-id="07a9c-p106">- **ContextMenuText** for Excel or Word. Displays the item on the context menu when text is selected and then the user right-clicks on the selected text. </span></span><br/> <span data-ttu-id="07a9c-p107">Excel の場合は - **ContextMenuCell**。ユーザーがスプレッドシートのセルを右クリックすると、コンテキスト メニューに項目が表示されます。</span><span class="sxs-lookup"><span data-stu-id="07a9c-p107">- **ContextMenuCell** for Excel. Displays the  item on the context menu when the user right-clicks on a cell on the spreadsheet.</span></span>|
|<span data-ttu-id="07a9c-139">**Group**</span><span class="sxs-lookup"><span data-stu-id="07a9c-139">**Group**</span></span>|<span data-ttu-id="07a9c-p108">タブのユーザー インターフェイスの拡張点のグループ。グループには、最大 6 個のコントロールを指定できます。 **id** 属性が必要です。id は最大 125 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="07a9c-p108">A group of user interface extension points on a tab. A group can have up to six controls. The  **id** attribute is required. It's a string with a maximum of 125 characters.</span></span>|
|<span data-ttu-id="07a9c-143">**Label**</span><span class="sxs-lookup"><span data-stu-id="07a9c-143">**Label**</span></span>|<span data-ttu-id="07a9c-p109">必須。グループのラベル。 **resid** 属性は、 **String** 要素の **id** 属性の値に設定する必要があります。 **String** 要素は、 **Resources** 要素の子要素である **ShortStrings** 要素の子要素です。</span><span class="sxs-lookup"><span data-stu-id="07a9c-p109">Required. The label of the group. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **ShortStrings** element, which is a child element of the **Resources** element.</span></span>|
|<span data-ttu-id="07a9c-148">**Icon**</span><span class="sxs-lookup"><span data-stu-id="07a9c-148">**Icon**</span></span>|<span data-ttu-id="07a9c-p110">必須。小さいフォーム ファクターのデバイス、または表示されるボタンが多すぎるときに使用されるグループのアイコンを指定します。 **resid** 属性は、 **Image** 要素の **id** 属性の値に設定する必要があります。 **Image** 要素は、 **Resources** 要素の子要素である **Images** 要素の子要素です。 **size** 属性は、イメージのサイズをピクセル単位で指定します。3 つのイメージのサイズ (16、32、80) が必要です。5 つのオプションのサイズ (20、24、40、48、64) もサポートされています。</span><span class="sxs-lookup"><span data-stu-id="07a9c-p110">Required. Specifies the group's icon to be used on small form factor devices, or when too many buttons are displayed. The  **resid** attribute must be set to the value of the **id** attribute of an **Image** element. The **Image** element is a child element of the **Images** element, which is a child element of the **Resources** element. The **size** attribute gives the size, in pixels, of the image. Three image sizes are required: 16, 32, and 80. Five optional sizes are also supported: 20, 24, 40, 48, and 64.</span></span>|
|<span data-ttu-id="07a9c-156">**Tooltip**</span><span class="sxs-lookup"><span data-stu-id="07a9c-156">**Tooltip**</span></span>|<span data-ttu-id="07a9c-p111">省略可能。グループのツールヒント。 **resid** 属性は、 **String** 要素の **id** 属性の値に設定する必要があります。 **String** 要素は、 **Resources** 要素の子要素である **LongStrings** 要素の子要素です。</span><span class="sxs-lookup"><span data-stu-id="07a9c-p111">Optional. The tooltip of the group. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **LongStrings** element, which is a child element of the **Resources** element.</span></span>|
|<span data-ttu-id="07a9c-161">**Control**</span><span class="sxs-lookup"><span data-stu-id="07a9c-161">**Control**</span></span>|<span data-ttu-id="07a9c-162">各グループには、少なくとも 1 つのコントロールが必要です。</span><span class="sxs-lookup"><span data-stu-id="07a9c-162">Each group requires at least one control.</span></span> <span data-ttu-id="07a9c-163">**Control** 要素は、**Button** または **Menu** のどちらかになります。</span><span class="sxs-lookup"><span data-stu-id="07a9c-163">A  **Control** element can be either a **Button** or a **Menu**.</span></span> <span data-ttu-id="07a9c-164">ボタン コントロールのドロップダウン リストを指定する場合は、**Menu** を使用します。</span><span class="sxs-lookup"><span data-stu-id="07a9c-164">Use  **Menu** to specify a drop-down list of button controls.</span></span> <span data-ttu-id="07a9c-165">現在は、ボタンとメニューのみがサポートされています。</span><span class="sxs-lookup"><span data-stu-id="07a9c-165">Currently, only buttons and menus are supported.</span></span> <span data-ttu-id="07a9c-166">詳細については、「[Button コントロール](control.md#button-control)」および「[Menu コントロール](control.md#menu-dropdown-button-controls)」のセクションを参照してください。</span><span class="sxs-lookup"><span data-stu-id="07a9c-166">See the [Button controls](control.md#button-control) and [Menu controls](control.md#menu-dropdown-button-controls) sections for more information.</span></span><br/><span data-ttu-id="07a9c-167">**注:** トラブルシューティングを容易にするためには、**コントロール**の要素および関連する**リソース**の子要素が、一度に 1 つに追加することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="07a9c-167">**Note:**  To make troubleshooting easier, we recommend that a  **Control** element and the related **Resources** child elements be added one at a time.</span></span>|
|<span data-ttu-id="07a9c-168">**Script**</span><span class="sxs-lookup"><span data-stu-id="07a9c-168">**Script**</span></span>|<span data-ttu-id="07a9c-169">カスタム関数の定義と登録コードを含む JavaScript ファイルにリンクします。</span><span class="sxs-lookup"><span data-stu-id="07a9c-169">Links to the JavaScript file with the custom function definition and registration code.</span></span> <span data-ttu-id="07a9c-170">Developer Preview では、この要素は使用しません。</span><span class="sxs-lookup"><span data-stu-id="07a9c-170">This element is not used in the Developer Preview.</span></span> <span data-ttu-id="07a9c-171">代わりに、HTML ページはすべての JavaScript ファイルを読み込みます。</span><span class="sxs-lookup"><span data-stu-id="07a9c-171">Instead, the HTML page is responsible for loading all JavaScript files.</span></span>|
|<span data-ttu-id="07a9c-172">**Page**</span><span class="sxs-lookup"><span data-stu-id="07a9c-172">**Page**</span></span>|<span data-ttu-id="07a9c-173">カスタム関数についての HTML ページにリンクします。</span><span class="sxs-lookup"><span data-stu-id="07a9c-173">Links to the HTML page for your custom functions.</span></span>|

## <a name="extension-points-for-outlook"></a><span data-ttu-id="07a9c-174">Outlook の拡張ポイント</span><span class="sxs-lookup"><span data-stu-id="07a9c-174">Extension points for Outlook</span></span>

- [<span data-ttu-id="07a9c-175">MessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="07a9c-175">MessageReadCommandSurface</span></span>](#messagereadcommandsurface) 
- [<span data-ttu-id="07a9c-176">MessageComposeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="07a9c-176">MessageComposeCommandSurface</span></span>](#messagecomposecommandsurface) 
- [<span data-ttu-id="07a9c-177">AppointmentOrganizerCommandSurface</span><span class="sxs-lookup"><span data-stu-id="07a9c-177">AppointmentOrganizerCommandSurface</span></span>](#appointmentorganizercommandsurface) 
- [<span data-ttu-id="07a9c-178">AppointmentAttendeeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="07a9c-178">AppointmentAttendeeCommandSurface</span></span>](#appointmentattendeecommandsurface)
- <span data-ttu-id="07a9c-179">[Module](#module) ([DesktopFormFactor](desktopformfactor.md) でのみ使用できます。)</span><span class="sxs-lookup"><span data-stu-id="07a9c-179">[Module](#module) (Can only be used in the [DesktopFormFactor](desktopformfactor.md).)</span></span>
- [<span data-ttu-id="07a9c-180">MobileMessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="07a9c-180">MobileMessageReadCommandSurface</span></span>](#mobilemessagereadcommandsurface)
- [<span data-ttu-id="07a9c-181">Events</span><span class="sxs-lookup"><span data-stu-id="07a9c-181">Events</span></span>](#events)
- [<span data-ttu-id="07a9c-182">DetectedEntity</span><span class="sxs-lookup"><span data-stu-id="07a9c-182">DetectedEntity</span></span>](#detectedentity)

### <a name="messagereadcommandsurface"></a><span data-ttu-id="07a9c-183">MessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="07a9c-183">MessageReadCommandSurface</span></span>
<span data-ttu-id="07a9c-p114">この拡張点により、メールの閲覧ビューのコマンド サーフェスにボタンが配置されます。Outlook デスクトップでは、これはリボンに表示されます。</span><span class="sxs-lookup"><span data-stu-id="07a9c-p114">This extension point puts buttons in the command surface for the mail read view. In Outlook desktop, this appears in the ribbon.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="07a9c-186">子要素</span><span class="sxs-lookup"><span data-stu-id="07a9c-186">Child elements</span></span>

|  <span data-ttu-id="07a9c-187">要素</span><span class="sxs-lookup"><span data-stu-id="07a9c-187">Element</span></span> |  <span data-ttu-id="07a9c-188">説明</span><span class="sxs-lookup"><span data-stu-id="07a9c-188">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="07a9c-189">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="07a9c-189">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="07a9c-190">コマンドを既定のリボン タブに追加します。</span><span class="sxs-lookup"><span data-stu-id="07a9c-190">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="07a9c-191">CustomTab</span><span class="sxs-lookup"><span data-stu-id="07a9c-191">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="07a9c-192">コマンドをカスタム リボン タブに追加します。</span><span class="sxs-lookup"><span data-stu-id="07a9c-192">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="07a9c-193">OfficeTab の例</span><span class="sxs-lookup"><span data-stu-id="07a9c-193">OfficeTab example</span></span>
```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="07a9c-194">CustomTab の例</span><span class="sxs-lookup"><span data-stu-id="07a9c-194">CustomTab example</span></span>
```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="messagecomposecommandsurface"></a><span data-ttu-id="07a9c-195">MessageComposeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="07a9c-195">MessageComposeCommandSurface</span></span>
<span data-ttu-id="07a9c-196">この拡張点は、メールの新規作成フォームを使用してアドイン用のリボンにボタンを配置します。</span><span class="sxs-lookup"><span data-stu-id="07a9c-196">This extension point puts buttons on the ribbon for add-ins using mail compose form.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="07a9c-197">子要素</span><span class="sxs-lookup"><span data-stu-id="07a9c-197">Child elements</span></span>

|  <span data-ttu-id="07a9c-198">要素</span><span class="sxs-lookup"><span data-stu-id="07a9c-198">Element</span></span> |  <span data-ttu-id="07a9c-199">説明</span><span class="sxs-lookup"><span data-stu-id="07a9c-199">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="07a9c-200">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="07a9c-200">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="07a9c-201">コマンドを既定のリボン タブに追加します。</span><span class="sxs-lookup"><span data-stu-id="07a9c-201">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="07a9c-202">CustomTab</span><span class="sxs-lookup"><span data-stu-id="07a9c-202">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="07a9c-203">コマンドをカスタム リボン タブに追加します。</span><span class="sxs-lookup"><span data-stu-id="07a9c-203">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="07a9c-204">OfficeTab の例</span><span class="sxs-lookup"><span data-stu-id="07a9c-204">OfficeTab example</span></span>
```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="07a9c-205">CustomTab の例</span><span class="sxs-lookup"><span data-stu-id="07a9c-205">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="appointmentorganizercommandsurface"></a><span data-ttu-id="07a9c-206">AppointmentOrganizerCommandSurface</span><span class="sxs-lookup"><span data-stu-id="07a9c-206">AppointmentOrganizerCommandSurface</span></span>

<span data-ttu-id="07a9c-207">この拡張点は、会議の開催者に表示されるフォームのリボンにボタンを配置します。</span><span class="sxs-lookup"><span data-stu-id="07a9c-207">This extension point puts buttons on the ribbon for the form that's displayed to the organizer of the meeting.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="07a9c-208">子要素</span><span class="sxs-lookup"><span data-stu-id="07a9c-208">Child elements</span></span>

|  <span data-ttu-id="07a9c-209">要素</span><span class="sxs-lookup"><span data-stu-id="07a9c-209">Element</span></span> |  <span data-ttu-id="07a9c-210">説明</span><span class="sxs-lookup"><span data-stu-id="07a9c-210">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="07a9c-211">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="07a9c-211">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="07a9c-212">コマンドを既定のリボン タブに追加します。</span><span class="sxs-lookup"><span data-stu-id="07a9c-212">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="07a9c-213">CustomTab</span><span class="sxs-lookup"><span data-stu-id="07a9c-213">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="07a9c-214">コマンドをカスタム リボン タブに追加します。</span><span class="sxs-lookup"><span data-stu-id="07a9c-214">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="07a9c-215">OfficeTab の例</span><span class="sxs-lookup"><span data-stu-id="07a9c-215">OfficeTab example</span></span>
```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="07a9c-216">CustomTab の例</span><span class="sxs-lookup"><span data-stu-id="07a9c-216">CustomTab example</span></span>
```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="appointmentattendeecommandsurface"></a><span data-ttu-id="07a9c-217">AppointmentAttendeeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="07a9c-217">AppointmentAttendeeCommandSurface</span></span>

<span data-ttu-id="07a9c-218">この拡張点は、会議の出席者に表示されるフォームのリボンにボタンを配置します。</span><span class="sxs-lookup"><span data-stu-id="07a9c-218">This extension point puts buttons on the ribbon for the form that's displayed to the attendee of the meeting.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="07a9c-219">子要素</span><span class="sxs-lookup"><span data-stu-id="07a9c-219">Child elements</span></span>

|  <span data-ttu-id="07a9c-220">要素</span><span class="sxs-lookup"><span data-stu-id="07a9c-220">Element</span></span> |  <span data-ttu-id="07a9c-221">説明</span><span class="sxs-lookup"><span data-stu-id="07a9c-221">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="07a9c-222">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="07a9c-222">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="07a9c-223">コマンドを既定のリボン タブに追加します。</span><span class="sxs-lookup"><span data-stu-id="07a9c-223">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="07a9c-224">CustomTab</span><span class="sxs-lookup"><span data-stu-id="07a9c-224">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="07a9c-225">コマンドをカスタム リボン タブに追加します。</span><span class="sxs-lookup"><span data-stu-id="07a9c-225">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="07a9c-226">OfficeTab の例</span><span class="sxs-lookup"><span data-stu-id="07a9c-226">OfficeTab example</span></span>
```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="07a9c-227">CustomTab の例</span><span class="sxs-lookup"><span data-stu-id="07a9c-227">CustomTab example</span></span>
```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="module"></a><span data-ttu-id="07a9c-228">Module</span><span class="sxs-lookup"><span data-stu-id="07a9c-228">Module</span></span>

<span data-ttu-id="07a9c-229">この拡張点は、モジュール拡張機能用のリボンにボタンを配置します。</span><span class="sxs-lookup"><span data-stu-id="07a9c-229">This extension point puts buttons on the ribbon for the module extension.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="07a9c-230">子要素</span><span class="sxs-lookup"><span data-stu-id="07a9c-230">Child elements</span></span>

|  <span data-ttu-id="07a9c-231">要素</span><span class="sxs-lookup"><span data-stu-id="07a9c-231">Element</span></span> |  <span data-ttu-id="07a9c-232">説明</span><span class="sxs-lookup"><span data-stu-id="07a9c-232">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="07a9c-233">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="07a9c-233">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="07a9c-234">コマンドを既定のリボン タブに追加します。</span><span class="sxs-lookup"><span data-stu-id="07a9c-234">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="07a9c-235">CustomTab</span><span class="sxs-lookup"><span data-stu-id="07a9c-235">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="07a9c-236">コマンドをカスタム リボン タブに追加します。</span><span class="sxs-lookup"><span data-stu-id="07a9c-236">Adds the command(s) to the custom ribbon tab.</span></span>  |

### <a name="mobilemessagereadcommandsurface"></a><span data-ttu-id="07a9c-237">MobileMessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="07a9c-237">MobileMessageReadCommandSurface</span></span>
<span data-ttu-id="07a9c-238">この拡張点により、モバイル フォーム ファクターのメールの閲覧ビューのコマンド領域にボタンが配置されます。</span><span class="sxs-lookup"><span data-stu-id="07a9c-238">This extension point puts buttons in the command surface for the mail read view in the mobile form factor.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="07a9c-239">子要素</span><span class="sxs-lookup"><span data-stu-id="07a9c-239">Child elements</span></span>

|  <span data-ttu-id="07a9c-240">要素</span><span class="sxs-lookup"><span data-stu-id="07a9c-240">Element</span></span> |  <span data-ttu-id="07a9c-241">説明</span><span class="sxs-lookup"><span data-stu-id="07a9c-241">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="07a9c-242">Group</span><span class="sxs-lookup"><span data-stu-id="07a9c-242">Group</span></span>](group.md) |  <span data-ttu-id="07a9c-243">コマンド領域にボタンのグループを追加します。</span><span class="sxs-lookup"><span data-stu-id="07a9c-243">Adds a group of buttons to the command surface.</span></span>  |

<span data-ttu-id="07a9c-244">この種類の **ExtensionPoint** 要素には子要素を 1 つだけ含めることができます (**Group** 要素)。</span><span class="sxs-lookup"><span data-stu-id="07a9c-244">**ExtensionPoint** elements of this type can only have one child element: a **Group** element.</span></span>

<span data-ttu-id="07a9c-245">この拡張点に含まれる **Control** 要素の **xsi:type** 属性を `MobileButton` に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="07a9c-245">**Control** elements contained in this extension point must have the **xsi:type** attribute set to `MobileButton`.</span></span>

#### <a name="example"></a><span data-ttu-id="07a9c-246">例</span><span class="sxs-lookup"><span data-stu-id="07a9c-246">Example</span></span>
```xml
<ExtensionPoint xsi:type="MobileMessageReadCommandSurface">
  <Group id="mobileGroupID">
    <Label resid="residAppName"/>
      <Control id="mobileButton1" xsi:type="MobileButton">
        <!-- Control definition -->
      </Control>
  </Group>
</ExtensionPoint>
```

### <a name="events"></a><span data-ttu-id="07a9c-247">Events</span><span class="sxs-lookup"><span data-stu-id="07a9c-247">Events</span></span>

<span data-ttu-id="07a9c-248">この拡張点は、指定したイベントのイベント ハンドラーを追加します。</span><span class="sxs-lookup"><span data-stu-id="07a9c-248">This extension point adds an event handler for a specified event.</span></span>

> [!NOTE]
> <span data-ttu-id="07a9c-249">この要素の型は、Outlook を Office 365 に web 上でのみサポートされます。</span><span class="sxs-lookup"><span data-stu-id="07a9c-249">This element type is only supported by Outlook on the web in Office 365.</span></span>

| <span data-ttu-id="07a9c-250">要素</span><span class="sxs-lookup"><span data-stu-id="07a9c-250">Element</span></span> | <span data-ttu-id="07a9c-251">説明</span><span class="sxs-lookup"><span data-stu-id="07a9c-251">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="07a9c-252">Event</span><span class="sxs-lookup"><span data-stu-id="07a9c-252">Event</span></span>](event.md) |  <span data-ttu-id="07a9c-253">イベントとイベント ハンドラーの関数を指定します。</span><span class="sxs-lookup"><span data-stu-id="07a9c-253">Specifies the event and event handler function.</span></span>  |

#### <a name="itemsend-event-example"></a><span data-ttu-id="07a9c-254">ItemSend イベントの例</span><span class="sxs-lookup"><span data-stu-id="07a9c-254">ItemSend event example</span></span>

```xml
<ExtensionPoint xsi:type="Events"> 
  <Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="itemSendHandler" /> 
</ExtensionPoint> 
```

### <a name="detectedentity"></a><span data-ttu-id="07a9c-255">DetectedEntity</span><span class="sxs-lookup"><span data-stu-id="07a9c-255">DetectedEntity</span></span>

<span data-ttu-id="07a9c-256">この拡張点は、指定したエンティティの種類に対するコンテキスト アドインのアクティブ化を追加します。</span><span class="sxs-lookup"><span data-stu-id="07a9c-256">This extension point adds a contextual add-in activation on a specified entity type.</span></span>

<span data-ttu-id="07a9c-257">これを収容している [VersionOverrides](versionoverrides.md) 要素は、`xsi:type` 属性の値が `VersionOverridesV1_1` になっている必要があります。</span><span class="sxs-lookup"><span data-stu-id="07a9c-257">The containing [VersionOverrides](versionoverrides.md) element must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span>

> [!NOTE]
> <span data-ttu-id="07a9c-258">この要素の型は、Outlook を Office 365 に web 上でのみサポートされます。</span><span class="sxs-lookup"><span data-stu-id="07a9c-258">This element type is only supported by Outlook on the web in Office 365.</span></span>

|  <span data-ttu-id="07a9c-259">要素</span><span class="sxs-lookup"><span data-stu-id="07a9c-259">Element</span></span> |  <span data-ttu-id="07a9c-260">説明</span><span class="sxs-lookup"><span data-stu-id="07a9c-260">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="07a9c-261">Label</span><span class="sxs-lookup"><span data-stu-id="07a9c-261">Label</span></span>](#label) |  <span data-ttu-id="07a9c-262">アドインのコンテキスト ウィンドウのラベルを指定します。</span><span class="sxs-lookup"><span data-stu-id="07a9c-262">Specifies the label for the add-in in the contextual window.</span></span>  |
|  [<span data-ttu-id="07a9c-263">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="07a9c-263">SourceLocation</span></span>](sourcelocation.md) |  <span data-ttu-id="07a9c-264">コンテキスト ウィンドウの URL を指定します。</span><span class="sxs-lookup"><span data-stu-id="07a9c-264">Specifies the URL for the contextual window.</span></span>  |
|  [<span data-ttu-id="07a9c-265">Rule</span><span class="sxs-lookup"><span data-stu-id="07a9c-265">Rule</span></span>](rule.md) |  <span data-ttu-id="07a9c-266">アドインをアクティブ化するタイミングを決定する 1 つ以上のルールを指定します。</span><span class="sxs-lookup"><span data-stu-id="07a9c-266">Specifies the rule or rules that determine when an add-in activates.</span></span>  |

#### <a name="label"></a><span data-ttu-id="07a9c-267">Label</span><span class="sxs-lookup"><span data-stu-id="07a9c-267">Label</span></span>

<span data-ttu-id="07a9c-p115">必ず指定します。グループのラベルです。 **resid** 属性には、 **Resources** 要素の **ShortStrings** 要素にある **String** 要素の [id](resources.md) 属性の値を設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="07a9c-p115">Required. The label of the group. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

#### <a name="highlight-requirements"></a><span data-ttu-id="07a9c-271">強調表示の要件</span><span class="sxs-lookup"><span data-stu-id="07a9c-271">Highlight requirements</span></span>

<span data-ttu-id="07a9c-p116">ユーザーは、強調表示されたエンティティに対話型の操作を実行する方法でのみコンテキスト アドインを有効化できます。開発者は、`ItemHasKnownEntity` および `ItemHasRegularExpressionMatch` のルールの種類に対応する `Rule` 要素の `Highlight` 属性を使用して、強調表示にするエンティティを制御します。</span><span class="sxs-lookup"><span data-stu-id="07a9c-p116">The only way a user can activate a contextual add-in is to interact with a highlighted entity. Developers can control which entities are highlighted by using the `Highlight` attribute of the `Rule` element for `ItemHasKnownEntity` and `ItemHasRegularExpressionMatch` rule types.</span></span>

<span data-ttu-id="07a9c-p117">ただし、注意する必要のある制限があります。これらの制限は、ユーザーにアドインをアクティブ化する方法を提供するために、適用可能なメッセージや予定で強調表示されたエンティティが常に存在するようにするために実施されます。</span><span class="sxs-lookup"><span data-stu-id="07a9c-p117">However, there are some limitations to be aware of. These limitations are in place to ensure that there will always be a highlighted entity in applicable messages or appointments to give the user a way to activate the add-in.</span></span>

- <span data-ttu-id="07a9c-276">`EmailAddress` および `Url` のエンティティの種類は、強調表示できません。そのため、アドインをアクティブ化するためには使用できません。</span><span class="sxs-lookup"><span data-stu-id="07a9c-276">The `EmailAddress` and `Url` entity types cannot be highlighted, and therefore cannot be used to activate an add-in.</span></span>
- <span data-ttu-id="07a9c-277">単一のルールを使用する場合、`Highlight` は `all` に設定されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="07a9c-277">If using a single rule, `Highlight` MUST be set to `all`.</span></span>
- <span data-ttu-id="07a9c-278">複数のルールを組み合わせるために `Mode="AND"` で `RuleCollection` のルールの種類を使用する場合は、少なくとも 1 つのルールの `Highlight` が `all` に設定されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="07a9c-278">If using a `RuleCollection` rule type with `Mode="AND"` to combine multiple rules, at least one of the rules MUST have `Highlight` set to `all`.</span></span>
- <span data-ttu-id="07a9c-279">複数のルールを組み合わせるために `Mode="OR"` で `RuleCollection` のルールの種類を使用する場合は、すべてのルールの `Highlight` が `all` に設定されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="07a9c-279">If using a `RuleCollection` rule type with `Mode="OR"` to combine multiple rules, all of the rules MUST have `Highlight` set to `all`.</span></span>

#### <a name="detectedentity-event-example"></a><span data-ttu-id="07a9c-280">DetectedEntity イベントの例</span><span class="sxs-lookup"><span data-stu-id="07a9c-280">DetectedEntity event example</span></span>

```xml
<ExtensionPoint xsi:type="DetectedEntity">
  <Label resid="residLabelName"/>
  <SourceLocation resid="residDetectedEntityURL" />
  <Rule xsi:type="RuleCollection" Mode="And">
    <Rule xsi:type="ItemIs" ItemType="Message" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="MeetingSuggestion" Highlight="all" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" Highlight="none" />
  </Rule>
</ExtensionPoint> 
```