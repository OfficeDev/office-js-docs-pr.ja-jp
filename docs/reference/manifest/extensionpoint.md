---
title: マニフェスト ファイルの ExtensionPoint 要素
description: Office UI でアドインが機能を公開する場所を定義します。
ms.date: 05/18/2020
localization_priority: Normal
ms.openlocfilehash: 44824e0c74b35105833f1f05cdda87bc873a4427
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/10/2020
ms.locfileid: "45094457"
---
# <a name="extensionpoint-element"></a><span data-ttu-id="a9acc-103">ExtensionPoint 要素</span><span class="sxs-lookup"><span data-stu-id="a9acc-103">ExtensionPoint element</span></span>

 <span data-ttu-id="a9acc-104">Office UI でアドインが機能を公開する場所を定義します。</span><span class="sxs-lookup"><span data-stu-id="a9acc-104">Defines where an add-in exposes functionality in the Office UI.</span></span> <span data-ttu-id="a9acc-105">**ExtensionPoint** 要素は、[AllFormFactors](allformfactors.md)、[DesktopFormFactor](desktopformfactor.md)、[MobileFormFactor](mobileformfactor.md) の子要素です。</span><span class="sxs-lookup"><span data-stu-id="a9acc-105">The **ExtensionPoint** element is a child element of [AllFormFactors](allformfactors.md), [DesktopFormFactor](desktopformfactor.md) or [MobileFormFactor](mobileformfactor.md).</span></span>

## <a name="attributes"></a><span data-ttu-id="a9acc-106">属性</span><span class="sxs-lookup"><span data-stu-id="a9acc-106">Attributes</span></span>

|  <span data-ttu-id="a9acc-107">属性</span><span class="sxs-lookup"><span data-stu-id="a9acc-107">Attribute</span></span>  |  <span data-ttu-id="a9acc-108">必須</span><span class="sxs-lookup"><span data-stu-id="a9acc-108">Required</span></span>  |  <span data-ttu-id="a9acc-109">説明</span><span class="sxs-lookup"><span data-stu-id="a9acc-109">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="a9acc-110">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="a9acc-110">**xsi:type**</span></span>  |  <span data-ttu-id="a9acc-111">はい</span><span class="sxs-lookup"><span data-stu-id="a9acc-111">Yes</span></span>  | <span data-ttu-id="a9acc-112">定義される拡張点の種類。</span><span class="sxs-lookup"><span data-stu-id="a9acc-112">The type of extension point being defined.</span></span>|

## <a name="extension-points-for-excel-only"></a><span data-ttu-id="a9acc-113">Excel のみの拡張点</span><span class="sxs-lookup"><span data-stu-id="a9acc-113">Extension points for Excel only</span></span>

- <span data-ttu-id="a9acc-114">**CustomFunctions** - Excel 向けの JavaScript で記述されたカスタム関数。</span><span class="sxs-lookup"><span data-stu-id="a9acc-114">**CustomFunctions** - A custom function written in JavaScript for Excel.</span></span>

<span data-ttu-id="a9acc-115">[この XML コード サンプル](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/manifest.xml)は、**CustomFunctions** 属性の値を持つ **ExtensionPoint** 要素を使用する方法と、使用する子要素を示しています。</span><span class="sxs-lookup"><span data-stu-id="a9acc-115">[This XML code sample](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/manifest.xml) shows how to use the **ExtensionPoint** element with the **CustomFunctions** attribute value, and the child elements to be used.</span></span>

## <a name="extension-points-for-word-excel-powerpoint-and-onenote-add-in-commands"></a><span data-ttu-id="a9acc-116">Word、Excel、PowerPoint、OneNote アドイン コマンドの拡張点</span><span class="sxs-lookup"><span data-stu-id="a9acc-116">Extension points for Word, Excel, PowerPoint, and OneNote add-in commands</span></span>

- <span data-ttu-id="a9acc-117">**PrimaryCommandSurface** - Office のリボン。</span><span class="sxs-lookup"><span data-stu-id="a9acc-117">**PrimaryCommandSurface** - The ribbon in Office.</span></span>
- <span data-ttu-id="a9acc-118">**ContextMenu**Office UI で右クリックしたときに表示されるショートカット メニュー。</span><span class="sxs-lookup"><span data-stu-id="a9acc-118">**ContextMenu** - The shortcut menu that appears when you right-click in the Office UI.</span></span>

<span data-ttu-id="a9acc-119">次の例は、**PrimaryCommandSurface** と **ContextMenu** の属性値を持つ **ExtensionPoint** 要素を使用する方法と、各要素と併用する必要がある子要素を示しています。</span><span class="sxs-lookup"><span data-stu-id="a9acc-119">The following examples show how to use the **ExtensionPoint** element with **PrimaryCommandSurface** and **ContextMenu** attribute values, and the child elements that should be used with each.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="a9acc-p102">ID 属性を含む要素では、一意の ID を指定してください。会社の名前と ID を使用することをお勧めします。たとえば、次の形式にします。<CustomTab id="mycompanyname.mygroupname"></span><span class="sxs-lookup"><span data-stu-id="a9acc-p102">For elements that contain an ID attribute, make sure you provide a unique ID. We recommend that you use your company's name along with your ID. For example, use the following format. <CustomTab id="mycompanyname.mygroupname"></span></span>

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

#### <a name="child-elements"></a><span data-ttu-id="a9acc-123">子要素</span><span class="sxs-lookup"><span data-stu-id="a9acc-123">Child elements</span></span>
 
|<span data-ttu-id="a9acc-124">**Element**</span><span class="sxs-lookup"><span data-stu-id="a9acc-124">**Element**</span></span>|<span data-ttu-id="a9acc-125">**説明**</span><span class="sxs-lookup"><span data-stu-id="a9acc-125">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="a9acc-126">**CustomTab**</span><span class="sxs-lookup"><span data-stu-id="a9acc-126">**CustomTab**</span></span>|<span data-ttu-id="a9acc-p103">カスタム タブをリボンに追加する必要がある場合は必須 (**PrimaryCommandSurface** を使用)。**CustomTab** 要素を使用する場合、**OfficeTab** 要素は使用できません。**id** 属性が必要です。 </span><span class="sxs-lookup"><span data-stu-id="a9acc-p103">Required if you want to add a custom tab to the ribbon (using **PrimaryCommandSurface**). If you use the **CustomTab** element, you can't use the **OfficeTab** element. The **id** attribute is required.</span></span>|
|<span data-ttu-id="a9acc-130">**OfficeTab**</span><span class="sxs-lookup"><span data-stu-id="a9acc-130">**OfficeTab**</span></span>|<span data-ttu-id="a9acc-131">既定の Office アプリリボンタブ ( **Primarycommandsurface**を使用) を拡張する場合に必要です。</span><span class="sxs-lookup"><span data-stu-id="a9acc-131">Required if you want to extend a default Office app ribbon tab (using **PrimaryCommandSurface**).</span></span> <span data-ttu-id="a9acc-132">**Officetab**要素を使用する場合、 **customtab**要素は使用できません。</span><span class="sxs-lookup"><span data-stu-id="a9acc-132">If you use the **OfficeTab** element, you can't use the **CustomTab** element.</span></span> <span data-ttu-id="a9acc-133">詳細については、「[OfficeTab](officetab.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="a9acc-133">For details, see [OfficeTab](officetab.md).</span></span>|
|<span data-ttu-id="a9acc-134">**OfficeMenu**</span><span class="sxs-lookup"><span data-stu-id="a9acc-134">**OfficeMenu**</span></span>|<span data-ttu-id="a9acc-p105">既定のコンテキスト メニューにアドイン コマンドを追加する場合は必須 (**ContextMenu** を使用)。**id** 属性は以下に設定する必要があります。 </span><span class="sxs-lookup"><span data-stu-id="a9acc-p105">Required if you're adding add-in commands to a default context menu (using **ContextMenu**). The **id** attribute must be set to: </span></span><br/> <span data-ttu-id="a9acc-p106">Excel または Word の場合は - **ContextMenuText**。テキストが選択され、ユーザーが選択されたテキストを右クリックしたときに、コンテキスト メニューに項目が表示されます。 </span><span class="sxs-lookup"><span data-stu-id="a9acc-p106">- **ContextMenuText** for Excel or Word. Displays the item on the context menu when text is selected and then the user right-clicks on the selected text. </span></span><br/> <span data-ttu-id="a9acc-p107">Excel の場合は - **ContextMenuCell**。ユーザーがスプレッドシートのセルを右クリックすると、コンテキスト メニューに項目が表示されます。</span><span class="sxs-lookup"><span data-stu-id="a9acc-p107">- **ContextMenuCell** for Excel. Displays the  item on the context menu when the user right-clicks on a cell on the spreadsheet.</span></span>|
|<span data-ttu-id="a9acc-141">**グループ**</span><span class="sxs-lookup"><span data-stu-id="a9acc-141">**Group**</span></span>|<span data-ttu-id="a9acc-p108">タブのユーザー インターフェイスの拡張点のグループ。1 つのグループに、最大 6 個のコントロールを指定できます。**id** 属性が必要です。最大 125 文字の文字列です。 </span><span class="sxs-lookup"><span data-stu-id="a9acc-p108">A group of user interface extension points on a tab. A group can have up to six controls. The **id** attribute is required. It's a string with a maximum of 125 characters.</span></span>|
|<span data-ttu-id="a9acc-145">**Label**</span><span class="sxs-lookup"><span data-stu-id="a9acc-145">**Label**</span></span>|<span data-ttu-id="a9acc-p109">必須。グループのラベル。**resid** 属性は、**String** 要素の **id** 属性の値に設定する必要があります。**String** 要素は、**ShortStrings** 要素 (**Resources** 要素の子要素) の子要素です。 </span><span class="sxs-lookup"><span data-stu-id="a9acc-p109">Required. The label of the group. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **ShortStrings** element, which is a child element of the **Resources** element.</span></span>|
|<span data-ttu-id="a9acc-150">**Icon**</span><span class="sxs-lookup"><span data-stu-id="a9acc-150">**Icon**</span></span>|<span data-ttu-id="a9acc-p110">必須。小さいフォーム ファクターのデバイス、または多くのボタンが表示されるときに使用されるグループのアイコンを指定します。**resid** 属性は、**Image** 要素の **id** 属性の値に設定する必要があります。**Image** 要素は、**Images** 要素 (**Resources** 要素の子要素) の子要素です。**size** 属性は、イメージのサイズをピクセル単位で指定します。次の 3 つのイメージのサイズが必要です。16、32、および 80。次の 5 つのオプションのサイズもサポートされています。20、24、40、48、および 64。 </span><span class="sxs-lookup"><span data-stu-id="a9acc-p110">Required. Specifies the group's icon to be used on small form factor devices, or when too many buttons are displayed. The **resid** attribute must be set to the value of the **id** attribute of an **Image** element. The **Image** element is a child element of the **Images** element, which is a child element of the **Resources** element. The **size** attribute gives the size, in pixels, of the image. Three image sizes are required: 16, 32, and 80. Five optional sizes are also supported: 20, 24, 40, 48, and 64.</span></span>|
|<span data-ttu-id="a9acc-158">**Tooltip**</span><span class="sxs-lookup"><span data-stu-id="a9acc-158">**Tooltip**</span></span>|<span data-ttu-id="a9acc-p111">省略可能。グループのヒント。**resid** 属性は、**String** 要素の **id** 属性の値に設定する必要があります。**String** 要素は、**LongStrings** 要素 (**Resources** 要素の子要素) の子要素です。 </span><span class="sxs-lookup"><span data-stu-id="a9acc-p111">Optional. The tooltip of the group. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **LongStrings** element, which is a child element of the **Resources** element.</span></span>|
|<span data-ttu-id="a9acc-163">**Control**</span><span class="sxs-lookup"><span data-stu-id="a9acc-163">**Control**</span></span>|<span data-ttu-id="a9acc-164">各グループには、少なくとも 1 つのコントロールが必要です。</span><span class="sxs-lookup"><span data-stu-id="a9acc-164">Each group requires at least one control.</span></span> <span data-ttu-id="a9acc-165">**Control**要素には、**ボタン**または**メニュー**のいずれかを指定できます。</span><span class="sxs-lookup"><span data-stu-id="a9acc-165">A **Control** element can be either a **Button** or a **Menu**.</span></span> <span data-ttu-id="a9acc-166">**メニュー**を使用して、ボタンコントロールのドロップダウンリストを指定します。</span><span class="sxs-lookup"><span data-stu-id="a9acc-166">Use **Menu** to specify a drop-down list of button controls.</span></span> <span data-ttu-id="a9acc-167">現在は、ボタンとメニューのみがサポートされています。</span><span class="sxs-lookup"><span data-stu-id="a9acc-167">Currently, only buttons and menus are supported.</span></span> <span data-ttu-id="a9acc-168">詳細については、「[Button コントロール](control.md#button-control)」および「[Menu コントロール](control.md#menu-dropdown-button-controls)」のセクションを参照してください。</span><span class="sxs-lookup"><span data-stu-id="a9acc-168">See the [Button controls](control.md#button-control) and [Menu controls](control.md#menu-dropdown-button-controls) sections for more information.</span></span><br/><span data-ttu-id="a9acc-169">**注:** トラブルシューティングを簡単にするために、 **Control**要素と関連する**Resources**子要素を一度に1つずつ追加することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="a9acc-169">**Note:**  To make troubleshooting easier, we recommend that a **Control** element and the related **Resources** child elements be added one at a time.</span></span>|
|<span data-ttu-id="a9acc-170">**スクリプト**</span><span class="sxs-lookup"><span data-stu-id="a9acc-170">**Script**</span></span>|<span data-ttu-id="a9acc-171">カスタム関数の定義と登録コードを含む JavaScript ファイルにリンクします。</span><span class="sxs-lookup"><span data-stu-id="a9acc-171">Links to the JavaScript file with the custom function definition and registration code.</span></span> <span data-ttu-id="a9acc-172">Developer Preview では、この要素は使用しません。</span><span class="sxs-lookup"><span data-stu-id="a9acc-172">This element is not used in the Developer Preview.</span></span> <span data-ttu-id="a9acc-173">代わりに、HTML ページはすべての JavaScript ファイルを読み込みます。</span><span class="sxs-lookup"><span data-stu-id="a9acc-173">Instead, the HTML page is responsible for loading all JavaScript files.</span></span>|
|<span data-ttu-id="a9acc-174">**Page**</span><span class="sxs-lookup"><span data-stu-id="a9acc-174">**Page**</span></span>|<span data-ttu-id="a9acc-175">カスタム関数についての HTML ページにリンクします。</span><span class="sxs-lookup"><span data-stu-id="a9acc-175">Links to the HTML page for your custom functions.</span></span>|

## <a name="extension-points-for-outlook"></a><span data-ttu-id="a9acc-176">Outlook のみの拡張点</span><span class="sxs-lookup"><span data-stu-id="a9acc-176">Extension points for Outlook</span></span>

- [<span data-ttu-id="a9acc-177">MessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="a9acc-177">MessageReadCommandSurface</span></span>](#messagereadcommandsurface)
- [<span data-ttu-id="a9acc-178">MessageComposeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="a9acc-178">MessageComposeCommandSurface</span></span>](#messagecomposecommandsurface)
- [<span data-ttu-id="a9acc-179">AppointmentOrganizerCommandSurface</span><span class="sxs-lookup"><span data-stu-id="a9acc-179">AppointmentOrganizerCommandSurface</span></span>](#appointmentorganizercommandsurface)
- [<span data-ttu-id="a9acc-180">AppointmentAttendeeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="a9acc-180">AppointmentAttendeeCommandSurface</span></span>](#appointmentattendeecommandsurface)
- <span data-ttu-id="a9acc-181">[Module](#module) ([DesktopFormFactor](desktopformfactor.md) でのみ使用できます。)</span><span class="sxs-lookup"><span data-stu-id="a9acc-181">[Module](#module) (Can only be used in the [DesktopFormFactor](desktopformfactor.md).)</span></span>
- [<span data-ttu-id="a9acc-182">MobileMessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="a9acc-182">MobileMessageReadCommandSurface</span></span>](#mobilemessagereadcommandsurface)
- [<span data-ttu-id="a9acc-183">MobileOnlineMeetingCommandSurface</span><span class="sxs-lookup"><span data-stu-id="a9acc-183">MobileOnlineMeetingCommandSurface</span></span>](#mobileonlinemeetingcommandsurface-preview)
- [<span data-ttu-id="a9acc-184">LaunchEvent</span><span class="sxs-lookup"><span data-stu-id="a9acc-184">LaunchEvent</span></span>](#launchevent-preview)
- [<span data-ttu-id="a9acc-185">Events</span><span class="sxs-lookup"><span data-stu-id="a9acc-185">Events</span></span>](#events)
- [<span data-ttu-id="a9acc-186">DetectedEntity</span><span class="sxs-lookup"><span data-stu-id="a9acc-186">DetectedEntity</span></span>](#detectedentity)

### <a name="messagereadcommandsurface"></a><span data-ttu-id="a9acc-187">MessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="a9acc-187">MessageReadCommandSurface</span></span>

<span data-ttu-id="a9acc-p114">この拡張点により、メールの閲覧ビューのコマンド サーフェスにボタンが配置されます。Outlook デスクトップでは、これはリボンに表示されます。</span><span class="sxs-lookup"><span data-stu-id="a9acc-p114">This extension point puts buttons in the command surface for the mail read view. In Outlook desktop, this appears in the ribbon.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="a9acc-190">子要素</span><span class="sxs-lookup"><span data-stu-id="a9acc-190">Child elements</span></span>

|  <span data-ttu-id="a9acc-191">要素</span><span class="sxs-lookup"><span data-stu-id="a9acc-191">Element</span></span> |  <span data-ttu-id="a9acc-192">説明</span><span class="sxs-lookup"><span data-stu-id="a9acc-192">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="a9acc-193">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="a9acc-193">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="a9acc-194">コマンドを既定のリボン タブに追加します。</span><span class="sxs-lookup"><span data-stu-id="a9acc-194">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="a9acc-195">CustomTab</span><span class="sxs-lookup"><span data-stu-id="a9acc-195">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="a9acc-196">コマンドをカスタム リボン タブに追加します。</span><span class="sxs-lookup"><span data-stu-id="a9acc-196">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="a9acc-197">OfficeTab の例</span><span class="sxs-lookup"><span data-stu-id="a9acc-197">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="a9acc-198">CustomTab の例</span><span class="sxs-lookup"><span data-stu-id="a9acc-198">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="messagecomposecommandsurface"></a><span data-ttu-id="a9acc-199">MessageComposeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="a9acc-199">MessageComposeCommandSurface</span></span>

<span data-ttu-id="a9acc-200">この拡張点は、メールの新規作成フォームを使用してアドイン用のリボンにボタンを配置します。</span><span class="sxs-lookup"><span data-stu-id="a9acc-200">This extension point puts buttons on the ribbon for add-ins using mail compose form.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="a9acc-201">子要素</span><span class="sxs-lookup"><span data-stu-id="a9acc-201">Child elements</span></span>

|  <span data-ttu-id="a9acc-202">要素</span><span class="sxs-lookup"><span data-stu-id="a9acc-202">Element</span></span> |  <span data-ttu-id="a9acc-203">説明</span><span class="sxs-lookup"><span data-stu-id="a9acc-203">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="a9acc-204">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="a9acc-204">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="a9acc-205">コマンドを既定のリボン タブに追加します。</span><span class="sxs-lookup"><span data-stu-id="a9acc-205">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="a9acc-206">CustomTab</span><span class="sxs-lookup"><span data-stu-id="a9acc-206">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="a9acc-207">コマンドをカスタム リボン タブに追加します。</span><span class="sxs-lookup"><span data-stu-id="a9acc-207">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="a9acc-208">OfficeTab の例</span><span class="sxs-lookup"><span data-stu-id="a9acc-208">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="a9acc-209">CustomTab の例</span><span class="sxs-lookup"><span data-stu-id="a9acc-209">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="appointmentorganizercommandsurface"></a><span data-ttu-id="a9acc-210">AppointmentOrganizerCommandSurface</span><span class="sxs-lookup"><span data-stu-id="a9acc-210">AppointmentOrganizerCommandSurface</span></span>

<span data-ttu-id="a9acc-211">この拡張点は、会議の開催者に表示されるフォームのリボンにボタンを配置します。</span><span class="sxs-lookup"><span data-stu-id="a9acc-211">This extension point puts buttons on the ribbon for the form that's displayed to the organizer of the meeting.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="a9acc-212">子要素</span><span class="sxs-lookup"><span data-stu-id="a9acc-212">Child elements</span></span>

|  <span data-ttu-id="a9acc-213">要素</span><span class="sxs-lookup"><span data-stu-id="a9acc-213">Element</span></span> |  <span data-ttu-id="a9acc-214">説明</span><span class="sxs-lookup"><span data-stu-id="a9acc-214">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="a9acc-215">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="a9acc-215">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="a9acc-216">コマンドを既定のリボン タブに追加します。</span><span class="sxs-lookup"><span data-stu-id="a9acc-216">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="a9acc-217">CustomTab</span><span class="sxs-lookup"><span data-stu-id="a9acc-217">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="a9acc-218">コマンドをカスタム リボン タブに追加します。</span><span class="sxs-lookup"><span data-stu-id="a9acc-218">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="a9acc-219">OfficeTab の例</span><span class="sxs-lookup"><span data-stu-id="a9acc-219">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="a9acc-220">CustomTab の例</span><span class="sxs-lookup"><span data-stu-id="a9acc-220">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="appointmentattendeecommandsurface"></a><span data-ttu-id="a9acc-221">AppointmentAttendeeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="a9acc-221">AppointmentAttendeeCommandSurface</span></span>

<span data-ttu-id="a9acc-222">この拡張点は、会議の出席者に表示されるフォームのリボンにボタンを配置します。</span><span class="sxs-lookup"><span data-stu-id="a9acc-222">This extension point puts buttons on the ribbon for the form that's displayed to the attendee of the meeting.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="a9acc-223">子要素</span><span class="sxs-lookup"><span data-stu-id="a9acc-223">Child elements</span></span>

|  <span data-ttu-id="a9acc-224">要素</span><span class="sxs-lookup"><span data-stu-id="a9acc-224">Element</span></span> |  <span data-ttu-id="a9acc-225">説明</span><span class="sxs-lookup"><span data-stu-id="a9acc-225">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="a9acc-226">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="a9acc-226">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="a9acc-227">コマンドを既定のリボン タブに追加します。</span><span class="sxs-lookup"><span data-stu-id="a9acc-227">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="a9acc-228">CustomTab</span><span class="sxs-lookup"><span data-stu-id="a9acc-228">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="a9acc-229">コマンドをカスタム リボン タブに追加します。</span><span class="sxs-lookup"><span data-stu-id="a9acc-229">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="a9acc-230">OfficeTab の例</span><span class="sxs-lookup"><span data-stu-id="a9acc-230">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="a9acc-231">CustomTab の例</span><span class="sxs-lookup"><span data-stu-id="a9acc-231">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="module"></a><span data-ttu-id="a9acc-232">Module</span><span class="sxs-lookup"><span data-stu-id="a9acc-232">Module</span></span>

<span data-ttu-id="a9acc-233">この拡張点は、モジュール拡張機能用のリボンにボタンを配置します。</span><span class="sxs-lookup"><span data-stu-id="a9acc-233">This extension point puts buttons on the ribbon for the module extension.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="a9acc-234">子要素</span><span class="sxs-lookup"><span data-stu-id="a9acc-234">Child elements</span></span>

|  <span data-ttu-id="a9acc-235">要素</span><span class="sxs-lookup"><span data-stu-id="a9acc-235">Element</span></span> |  <span data-ttu-id="a9acc-236">説明</span><span class="sxs-lookup"><span data-stu-id="a9acc-236">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="a9acc-237">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="a9acc-237">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="a9acc-238">コマンドを既定のリボン タブに追加します。</span><span class="sxs-lookup"><span data-stu-id="a9acc-238">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="a9acc-239">CustomTab</span><span class="sxs-lookup"><span data-stu-id="a9acc-239">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="a9acc-240">コマンドをカスタム リボン タブに追加します。</span><span class="sxs-lookup"><span data-stu-id="a9acc-240">Adds the command(s) to the custom ribbon tab.</span></span>  |

### <a name="mobilemessagereadcommandsurface"></a><span data-ttu-id="a9acc-241">MobileMessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="a9acc-241">MobileMessageReadCommandSurface</span></span>

<span data-ttu-id="a9acc-242">この拡張点により、モバイル フォーム ファクターのメールの閲覧ビューのコマンド領域にボタンが配置されます。</span><span class="sxs-lookup"><span data-stu-id="a9acc-242">This extension point puts buttons in the command surface for the mail read view in the mobile form factor.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="a9acc-243">子要素</span><span class="sxs-lookup"><span data-stu-id="a9acc-243">Child elements</span></span>

|  <span data-ttu-id="a9acc-244">要素</span><span class="sxs-lookup"><span data-stu-id="a9acc-244">Element</span></span> |  <span data-ttu-id="a9acc-245">説明</span><span class="sxs-lookup"><span data-stu-id="a9acc-245">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="a9acc-246">Group</span><span class="sxs-lookup"><span data-stu-id="a9acc-246">Group</span></span>](group.md) |  <span data-ttu-id="a9acc-247">コマンド領域にボタンのグループを追加します。</span><span class="sxs-lookup"><span data-stu-id="a9acc-247">Adds a group of buttons to the command surface.</span></span>  |

<span data-ttu-id="a9acc-248">この種類の **ExtensionPoint** 要素には子要素を 1 つだけ含めることができます (**Group** 要素)。</span><span class="sxs-lookup"><span data-stu-id="a9acc-248">**ExtensionPoint** elements of this type can only have one child element: a **Group** element.</span></span>

<span data-ttu-id="a9acc-249">この拡張点に含まれる **Control** 要素の **xsi:type** 属性を `MobileButton` に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="a9acc-249">**Control** elements contained in this extension point must have the **xsi:type** attribute set to `MobileButton`.</span></span>

#### <a name="example"></a><span data-ttu-id="a9acc-250">例</span><span class="sxs-lookup"><span data-stu-id="a9acc-250">Example</span></span>

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

### <a name="mobileonlinemeetingcommandsurface-preview"></a><span data-ttu-id="a9acc-251">MobileOnlineMeetingCommandSurface (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="a9acc-251">MobileOnlineMeetingCommandSurface (preview)</span></span>

> [!NOTE]
> <span data-ttu-id="a9acc-252">この拡張点は、Microsoft 365 サブスクリプションを使用した Android の[プレビュー](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md)でのみサポートされています。</span><span class="sxs-lookup"><span data-stu-id="a9acc-252">This extension point is only supported in [preview](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) on Android with a Microsoft 365 subscription.</span></span>

<span data-ttu-id="a9acc-253">この拡張点は、モバイルフォームファクターの予定に対して、適切なモードのトグルをコマンド画面に配置します。</span><span class="sxs-lookup"><span data-stu-id="a9acc-253">This extension point puts a mode-appropriate toggle in the command surface for an appointment in the mobile form factor.</span></span> <span data-ttu-id="a9acc-254">会議の開催者は、オンライン会議を作成できます。</span><span class="sxs-lookup"><span data-stu-id="a9acc-254">A meeting organizer can create an online meeting.</span></span> <span data-ttu-id="a9acc-255">出席者は、後でオンライン会議に参加できます。</span><span class="sxs-lookup"><span data-stu-id="a9acc-255">An attendee can subsequently join the online meeting.</span></span> <span data-ttu-id="a9acc-256">このシナリオの詳細については、「[オンライン会議プロバイダー用の Outlook モバイルアドインを作成](../../outlook/online-meeting.md)する」の記事を参照してください。</span><span class="sxs-lookup"><span data-stu-id="a9acc-256">To learn more about this scenario, see the [Create an Outlook mobile add-in for an online-meeting provider](../../outlook/online-meeting.md) article.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="a9acc-257">子要素</span><span class="sxs-lookup"><span data-stu-id="a9acc-257">Child elements</span></span>

|  <span data-ttu-id="a9acc-258">要素</span><span class="sxs-lookup"><span data-stu-id="a9acc-258">Element</span></span> |  <span data-ttu-id="a9acc-259">説明</span><span class="sxs-lookup"><span data-stu-id="a9acc-259">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="a9acc-260">Control</span><span class="sxs-lookup"><span data-stu-id="a9acc-260">Control</span></span>](control.md) |  <span data-ttu-id="a9acc-261">コマンド画面にボタンを追加します。</span><span class="sxs-lookup"><span data-stu-id="a9acc-261">Adds a button to the command surface.</span></span>  |

<span data-ttu-id="a9acc-262">`ExtensionPoint`この型の要素には、1つの子要素 (要素) のみを含めることができます `Control` 。</span><span class="sxs-lookup"><span data-stu-id="a9acc-262">`ExtensionPoint` elements of this type can only have one child element: a `Control` element.</span></span>

<span data-ttu-id="a9acc-263">`Control`この拡張点に含まれる要素の `xsi:type` 属性をに設定する必要があり `MobileButton` ます。</span><span class="sxs-lookup"><span data-stu-id="a9acc-263">The `Control` element contained in this extension point must have the `xsi:type` attribute set to `MobileButton`.</span></span>

<span data-ttu-id="a9acc-264">画像は、 `Icon` 16 進コード `#919191` または[他のカラー形式](https://convertingcolors.com/hex-color-919191.html)の同等の画像を使用したグレースケールである必要があります。</span><span class="sxs-lookup"><span data-stu-id="a9acc-264">The `Icon` images should be in grayscale using hex code `#919191` or its equivalent in [other color formats](https://convertingcolors.com/hex-color-919191.html).</span></span>

#### <a name="example"></a><span data-ttu-id="a9acc-265">例</span><span class="sxs-lookup"><span data-stu-id="a9acc-265">Example</span></span>

```xml
<ExtensionPoint xsi:type="MobileOnlineMeetingCommandSurface">
  <Control xsi:type="MobileButton" id="onlineMeetingFunctionButton">
    <Label resid="residUILessButton0Name" />
    <Icon>
      <bt:Image resid="UiLessIcon" size="25" scale="1" />
      <bt:Image resid="UiLessIcon" size="25" scale="2" />
      <bt:Image resid="UiLessIcon" size="25" scale="3" />
      <bt:Image resid="UiLessIcon" size="32" scale="1" />
      <bt:Image resid="UiLessIcon" size="32" scale="2" />
      <bt:Image resid="UiLessIcon" size="32" scale="3" />
      <bt:Image resid="UiLessIcon" size="48" scale="1" />
      <bt:Image resid="UiLessIcon" size="48" scale="2" />
      <bt:Image resid="UiLessIcon" size="48" scale="3" />
    </Icon>
    <Action xsi:type="ExecuteFunction">
      <FunctionName>insertContosoMeeting</FunctionName>
    </Action>
  </Control>
</ExtensionPoint>
```

### <a name="launchevent-preview"></a><span data-ttu-id="a9acc-266">LaunchEvent (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="a9acc-266">LaunchEvent (preview)</span></span>

> [!NOTE]
> <span data-ttu-id="a9acc-267">この拡張点は、Microsoft 365 サブスクリプションを使用する web 上の Outlook の[プレビュー](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md)でのみサポートされています。</span><span class="sxs-lookup"><span data-stu-id="a9acc-267">This extension point is only supported in [preview](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) in Outlook on the web with a Microsoft 365 subscription.</span></span>

<span data-ttu-id="a9acc-268">この拡張点により、デスクトップフォームファクターでサポートされているイベントに基づいてアドインをアクティブにすることができます。</span><span class="sxs-lookup"><span data-stu-id="a9acc-268">This extension point enables an add-in to activate based on supported events in the desktop form factor.</span></span> <span data-ttu-id="a9acc-269">現時点でサポートされているイベントは、とだけです `OnNewMessageCompose` `OnNewAppointmentOrganizer` 。</span><span class="sxs-lookup"><span data-stu-id="a9acc-269">Currently, the only supported events are `OnNewMessageCompose` and `OnNewAppointmentOrganizer`.</span></span> <span data-ttu-id="a9acc-270">このシナリオの詳細については、「[イベントベースのライセンス認証用に Outlook アドインを構成](../../outlook/autolaunch.md)する」の記事を参照してください。</span><span class="sxs-lookup"><span data-stu-id="a9acc-270">To learn more about this scenario, see the [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md) article.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="a9acc-271">子要素</span><span class="sxs-lookup"><span data-stu-id="a9acc-271">Child elements</span></span>

|  <span data-ttu-id="a9acc-272">要素</span><span class="sxs-lookup"><span data-stu-id="a9acc-272">Element</span></span> |  <span data-ttu-id="a9acc-273">説明</span><span class="sxs-lookup"><span data-stu-id="a9acc-273">Description</span></span>  |
|:-----|:-----|
| [<span data-ttu-id="a9acc-274">LaunchEvents</span><span class="sxs-lookup"><span data-stu-id="a9acc-274">LaunchEvents</span></span>](launchevents.md) |  <span data-ttu-id="a9acc-275">イベントベースのライセンス認証のための[Launchevent](launchevent.md)のリスト。</span><span class="sxs-lookup"><span data-stu-id="a9acc-275">List of [LaunchEvent](launchevent.md) for event-based activation.</span></span>  |
| [<span data-ttu-id="a9acc-276">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="a9acc-276">SourceLocation</span></span>](sourcelocation.md) |  <span data-ttu-id="a9acc-277">ソース JavaScript ファイルの場所。</span><span class="sxs-lookup"><span data-stu-id="a9acc-277">The location of the source JavaScript file.</span></span>  |

#### <a name="example"></a><span data-ttu-id="a9acc-278">例</span><span class="sxs-lookup"><span data-stu-id="a9acc-278">Example</span></span>

```xml
<ExtensionPoint xsi:type="LaunchEvent">
  <LaunchEvents>
    <LaunchEvent Type="OnNewMessageCompose" FunctionName="onMessageComposeHandler"/>
    <LaunchEvent Type="OnNewAppointmentOrganizer" FunctionName="onAppointmentComposeHandler"/>
  </LaunchEvents>
  <!-- Identifies the runtime to be used (also referenced by the Runtime element). -->
  <SourceLocation resid="WebViewRuntime.Url"/>
</ExtensionPoint>
```

### <a name="events"></a><span data-ttu-id="a9acc-279">Events</span><span class="sxs-lookup"><span data-stu-id="a9acc-279">Events</span></span>

<span data-ttu-id="a9acc-280">この拡張点は、指定したイベントのイベント ハンドラーを追加します。</span><span class="sxs-lookup"><span data-stu-id="a9acc-280">This extension point adds an event handler for a specified event.</span></span> <span data-ttu-id="a9acc-281">この拡張点の使用の詳細については、「 [Outlook アドインの送信時機能](../../outlook/outlook-on-send-addins.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="a9acc-281">For more information about using this extension point, see [On-send feature for Outlook add-ins](../../outlook/outlook-on-send-addins.md).</span></span>

| <span data-ttu-id="a9acc-282">要素</span><span class="sxs-lookup"><span data-stu-id="a9acc-282">Element</span></span> | <span data-ttu-id="a9acc-283">説明</span><span class="sxs-lookup"><span data-stu-id="a9acc-283">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="a9acc-284">Event</span><span class="sxs-lookup"><span data-stu-id="a9acc-284">Event</span></span>](event.md) |  <span data-ttu-id="a9acc-285">イベントとイベント ハンドラーの関数を指定します。</span><span class="sxs-lookup"><span data-stu-id="a9acc-285">Specifies the event and event handler function.</span></span>  |

#### <a name="itemsend-event-example"></a><span data-ttu-id="a9acc-286">ItemSend イベントの例</span><span class="sxs-lookup"><span data-stu-id="a9acc-286">ItemSend event example</span></span>

```xml
<ExtensionPoint xsi:type="Events">
  <Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="itemSendHandler" />
</ExtensionPoint>
```

### <a name="detectedentity"></a><span data-ttu-id="a9acc-287">DetectedEntity</span><span class="sxs-lookup"><span data-stu-id="a9acc-287">DetectedEntity</span></span>

<span data-ttu-id="a9acc-288">この拡張点は、指定したエンティティの種類に対するコンテキスト アドインのアクティブ化を追加します。</span><span class="sxs-lookup"><span data-stu-id="a9acc-288">This extension point adds a contextual add-in activation on a specified entity type.</span></span>

<span data-ttu-id="a9acc-289">これを収容している [VersionOverrides](versionoverrides.md) 要素は、`xsi:type` 属性の値が `VersionOverridesV1_1` になっている必要があります。</span><span class="sxs-lookup"><span data-stu-id="a9acc-289">The containing [VersionOverrides](versionoverrides.md) element must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span>

> [!NOTE]
> <span data-ttu-id="a9acc-290">この要素の種類は、[要件セット 1.6 以降をサポートする Outlook クライアント ](../requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)が利用できます。</span><span class="sxs-lookup"><span data-stu-id="a9acc-290">This element type is available to [Outlook clients that support requirement sets 1.6 and later](../requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients).</span></span>

|  <span data-ttu-id="a9acc-291">要素</span><span class="sxs-lookup"><span data-stu-id="a9acc-291">Element</span></span> |  <span data-ttu-id="a9acc-292">説明</span><span class="sxs-lookup"><span data-stu-id="a9acc-292">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="a9acc-293">Label</span><span class="sxs-lookup"><span data-stu-id="a9acc-293">Label</span></span>](#label) |  <span data-ttu-id="a9acc-294">アドインのコンテキスト ウィンドウのラベルを指定します。</span><span class="sxs-lookup"><span data-stu-id="a9acc-294">Specifies the label for the add-in in the contextual window.</span></span>  |
|  [<span data-ttu-id="a9acc-295">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="a9acc-295">SourceLocation</span></span>](sourcelocation.md) |  <span data-ttu-id="a9acc-296">コンテキスト ウィンドウの URL を指定します。</span><span class="sxs-lookup"><span data-stu-id="a9acc-296">Specifies the URL for the contextual window.</span></span>  |
|  [<span data-ttu-id="a9acc-297">Rule</span><span class="sxs-lookup"><span data-stu-id="a9acc-297">Rule</span></span>](rule.md) |  <span data-ttu-id="a9acc-298">アドインをアクティブ化するタイミングを決定する 1 つ以上のルールを指定します。</span><span class="sxs-lookup"><span data-stu-id="a9acc-298">Specifies the rule or rules that determine when an add-in activates.</span></span>  |

#### <a name="label"></a><span data-ttu-id="a9acc-299">Label</span><span class="sxs-lookup"><span data-stu-id="a9acc-299">Label</span></span>

<span data-ttu-id="a9acc-300">必ず指定します。</span><span class="sxs-lookup"><span data-stu-id="a9acc-300">Required.</span></span> <span data-ttu-id="a9acc-301">グループのラベルです。</span><span class="sxs-lookup"><span data-stu-id="a9acc-301">The label of the group.</span></span> <span data-ttu-id="a9acc-302">**Resid**属性は、 [Resources](resources.md)要素の Short **strings**要素の**String**要素の**id**属性の値に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="a9acc-302">The **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

#### <a name="highlight-requirements"></a><span data-ttu-id="a9acc-303">強調表示の要件</span><span class="sxs-lookup"><span data-stu-id="a9acc-303">Highlight requirements</span></span>

<span data-ttu-id="a9acc-p119">ユーザーは、強調表示されたエンティティに対話型の操作を実行する方法でのみコンテキスト アドインを有効化できます。開発者は、`ItemHasKnownEntity` および `ItemHasRegularExpressionMatch` のルールの種類に対応する `Rule` 要素の `Highlight` 属性を使用して、強調表示にするエンティティを制御します。</span><span class="sxs-lookup"><span data-stu-id="a9acc-p119">The only way a user can activate a contextual add-in is to interact with a highlighted entity. Developers can control which entities are highlighted by using the `Highlight` attribute of the `Rule` element for `ItemHasKnownEntity` and `ItemHasRegularExpressionMatch` rule types.</span></span>

<span data-ttu-id="a9acc-p120">ただし、注意する必要のある制限があります。これらの制限は、ユーザーにアドインをアクティブ化する方法を提供するために、適用可能なメッセージや予定で強調表示されたエンティティが常に存在するようにするために実施されます。</span><span class="sxs-lookup"><span data-stu-id="a9acc-p120">However, there are some limitations to be aware of. These limitations are in place to ensure that there will always be a highlighted entity in applicable messages or appointments to give the user a way to activate the add-in.</span></span>

- <span data-ttu-id="a9acc-308">`EmailAddress` および `Url` のエンティティの種類は、強調表示できません。そのため、アドインをアクティブ化するためには使用できません。</span><span class="sxs-lookup"><span data-stu-id="a9acc-308">The `EmailAddress` and `Url` entity types cannot be highlighted, and therefore cannot be used to activate an add-in.</span></span>
- <span data-ttu-id="a9acc-309">単一のルールを使用する場合、`Highlight` は `all` に設定されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="a9acc-309">If using a single rule, `Highlight` MUST be set to `all`.</span></span>
- <span data-ttu-id="a9acc-310">複数のルールを組み合わせるために `Mode="AND"` で `RuleCollection` のルールの種類を使用する場合は、少なくとも 1 つのルールの `Highlight` が `all` に設定されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="a9acc-310">If using a `RuleCollection` rule type with `Mode="AND"` to combine multiple rules, at least one of the rules MUST have `Highlight` set to `all`.</span></span>
- <span data-ttu-id="a9acc-311">複数のルールを組み合わせるために `Mode="OR"` で `RuleCollection` のルールの種類を使用する場合は、すべてのルールの `Highlight` が `all` に設定されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="a9acc-311">If using a `RuleCollection` rule type with `Mode="OR"` to combine multiple rules, all of the rules MUST have `Highlight` set to `all`.</span></span>

#### <a name="detectedentity-event-example"></a><span data-ttu-id="a9acc-312">DetectedEntity イベントの例</span><span class="sxs-lookup"><span data-stu-id="a9acc-312">DetectedEntity event example</span></span>

```xml
<ExtensionPoint xsi:type="DetectedEntity">
  <Label resid="residLabelName"/>
  <!--If you opt to include RequestedHeight, it must be between 140px to 450px, inclusive.-->
  <!--<RequestedHeight>360</RequestedHeight>-->
  <SourceLocation resid="residDetectedEntityURL" />
  <Rule xsi:type="RuleCollection" Mode="And">
    <Rule xsi:type="ItemIs" ItemType="Message" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="MeetingSuggestion" Highlight="all" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" Highlight="none" />
  </Rule>
</ExtensionPoint>
```
