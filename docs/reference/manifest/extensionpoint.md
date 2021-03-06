---
title: マニフェスト ファイルの ExtensionPoint 要素
description: Office UI でアドインが機能を公開する場所を定義します。
ms.date: 02/12/2021
localization_priority: Normal
ms.openlocfilehash: e5b638969730be47c30c98d4fc231e58d492ac36
ms.sourcegitcommit: e7009c565b18c607fe0868db2e26e250ad308dce
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/05/2021
ms.locfileid: "50505466"
---
# <a name="extensionpoint-element"></a><span data-ttu-id="bd9c3-103">ExtensionPoint 要素</span><span class="sxs-lookup"><span data-stu-id="bd9c3-103">ExtensionPoint element</span></span>

 <span data-ttu-id="bd9c3-104">Office UI でアドインが機能を公開する場所を定義します。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-104">Defines where an add-in exposes functionality in the Office UI.</span></span> <span data-ttu-id="bd9c3-105">**ExtensionPoint** 要素は、[AllFormFactors](allformfactors.md)、[DesktopFormFactor](desktopformfactor.md)、[MobileFormFactor](mobileformfactor.md) の子要素です。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-105">The **ExtensionPoint** element is a child element of [AllFormFactors](allformfactors.md), [DesktopFormFactor](desktopformfactor.md) or [MobileFormFactor](mobileformfactor.md).</span></span>

## <a name="attributes"></a><span data-ttu-id="bd9c3-106">属性</span><span class="sxs-lookup"><span data-stu-id="bd9c3-106">Attributes</span></span>

|  <span data-ttu-id="bd9c3-107">属性</span><span class="sxs-lookup"><span data-stu-id="bd9c3-107">Attribute</span></span>  |  <span data-ttu-id="bd9c3-108">必須</span><span class="sxs-lookup"><span data-stu-id="bd9c3-108">Required</span></span>  |  <span data-ttu-id="bd9c3-109">説明</span><span class="sxs-lookup"><span data-stu-id="bd9c3-109">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="bd9c3-110">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="bd9c3-110">**xsi:type**</span></span>  |  <span data-ttu-id="bd9c3-111">はい</span><span class="sxs-lookup"><span data-stu-id="bd9c3-111">Yes</span></span>  | <span data-ttu-id="bd9c3-112">定義される拡張点の種類。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-112">The type of extension point being defined.</span></span>|

## <a name="extension-points-for-excel-only"></a><span data-ttu-id="bd9c3-113">Excel のみの拡張点</span><span class="sxs-lookup"><span data-stu-id="bd9c3-113">Extension points for Excel only</span></span>

- <span data-ttu-id="bd9c3-114">**CustomFunctions** - Excel 向けの JavaScript で記述されたカスタム関数。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-114">**CustomFunctions** - A custom function written in JavaScript for Excel.</span></span>

<span data-ttu-id="bd9c3-115">[この XML コード サンプル](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/manifest.xml)は、**CustomFunctions** 属性の値を持つ **ExtensionPoint** 要素を使用する方法と、使用する子要素を示しています。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-115">[This XML code sample](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/manifest.xml) shows how to use the **ExtensionPoint** element with the **CustomFunctions** attribute value, and the child elements to be used.</span></span>

## <a name="extension-points-for-word-excel-powerpoint-and-onenote-add-in-commands"></a><span data-ttu-id="bd9c3-116">Word、Excel、PowerPoint、OneNote アドイン コマンドの拡張点</span><span class="sxs-lookup"><span data-stu-id="bd9c3-116">Extension points for Word, Excel, PowerPoint, and OneNote add-in commands</span></span>

- <span data-ttu-id="bd9c3-117">**PrimaryCommandSurface** - Office のリボン。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-117">**PrimaryCommandSurface** - The ribbon in Office.</span></span>
- <span data-ttu-id="bd9c3-118">**ContextMenu** Office UI で右クリックしたときに表示されるショートカット メニュー。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-118">**ContextMenu** - The shortcut menu that appears when you right-click in the Office UI.</span></span>

<span data-ttu-id="bd9c3-119">次の例は、**PrimaryCommandSurface** と **ContextMenu** の属性値を持つ **ExtensionPoint** 要素を使用する方法と、各要素と併用する必要がある子要素を示しています。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-119">The following examples show how to use the **ExtensionPoint** element with **PrimaryCommandSurface** and **ContextMenu** attribute values, and the child elements that should be used with each.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="bd9c3-p102">ID 属性を含む要素では、一意の ID を指定してください。会社の名前と ID を使用することをお勧めします。たとえば、次の形式にします。<CustomTab id="mycompanyname.mygroupname"></span><span class="sxs-lookup"><span data-stu-id="bd9c3-p102">For elements that contain an ID attribute, make sure you provide a unique ID. We recommend that you use your company's name along with your ID. For example, use the following format. <CustomTab id="mycompanyname.mygroupname"></span></span>

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

#### <a name="child-elements"></a><span data-ttu-id="bd9c3-123">子要素</span><span class="sxs-lookup"><span data-stu-id="bd9c3-123">Child elements</span></span>
 
|<span data-ttu-id="bd9c3-124">要素</span><span class="sxs-lookup"><span data-stu-id="bd9c3-124">Element</span></span>|<span data-ttu-id="bd9c3-125">説明</span><span class="sxs-lookup"><span data-stu-id="bd9c3-125">Description</span></span>|
|:-----|:-----|
|<span data-ttu-id="bd9c3-126">**CustomTab**</span><span class="sxs-lookup"><span data-stu-id="bd9c3-126">**CustomTab**</span></span>|<span data-ttu-id="bd9c3-p103">カスタム タブをリボンに追加する必要がある場合は必須 (**PrimaryCommandSurface** を使用)。**CustomTab** 要素を使用する場合、**OfficeTab** 要素は使用できません。**id** 属性が必要です。 </span><span class="sxs-lookup"><span data-stu-id="bd9c3-p103">Required if you want to add a custom tab to the ribbon (using **PrimaryCommandSurface**). If you use the **CustomTab** element, you can't use the **OfficeTab** element. The **id** attribute is required.</span></span>|
|<span data-ttu-id="bd9c3-130">**OfficeTab**</span><span class="sxs-lookup"><span data-stu-id="bd9c3-130">**OfficeTab**</span></span>|<span data-ttu-id="bd9c3-131">既定のアプリ リボン タブ **(PrimaryCommandSurface** をOfficeを拡張する場合は必須です。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-131">Required if you want to extend a default Office app ribbon tab (using **PrimaryCommandSurface**).</span></span> <span data-ttu-id="bd9c3-132">OfficeTab 要素 **を使用する** 場合は **、CustomTab 要素を使用** することはできません。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-132">If you use the **OfficeTab** element, you can't use the **CustomTab** element.</span></span> <span data-ttu-id="bd9c3-133">詳細については、「[OfficeTab](officetab.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-133">For details, see [OfficeTab](officetab.md).</span></span>|
|<span data-ttu-id="bd9c3-134">**OfficeMenu**</span><span class="sxs-lookup"><span data-stu-id="bd9c3-134">**OfficeMenu**</span></span>|<span data-ttu-id="bd9c3-p105">既定のコンテキスト メニューにアドイン コマンドを追加する場合は必須 (**ContextMenu** を使用)。**id** 属性は以下に設定する必要があります。 </span><span class="sxs-lookup"><span data-stu-id="bd9c3-p105">Required if you're adding add-in commands to a default context menu (using **ContextMenu**). The **id** attribute must be set to: </span></span><br/> <span data-ttu-id="bd9c3-p106">Excel または Word の場合は - **ContextMenuText**。テキストが選択され、ユーザーが選択されたテキストを右クリックしたときに、コンテキスト メニューに項目が表示されます。 </span><span class="sxs-lookup"><span data-stu-id="bd9c3-p106">- **ContextMenuText** for Excel or Word. Displays the item on the context menu when text is selected and then the user right-clicks on the selected text. </span></span><br/> <span data-ttu-id="bd9c3-p107">Excel の場合は - **ContextMenuCell**。ユーザーがスプレッドシートのセルを右クリックすると、コンテキスト メニューに項目が表示されます。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-p107">- **ContextMenuCell** for Excel. Displays the  item on the context menu when the user right-clicks on a cell on the spreadsheet.</span></span>|
|<span data-ttu-id="bd9c3-141">**Group**</span><span class="sxs-lookup"><span data-stu-id="bd9c3-141">**Group**</span></span>|<span data-ttu-id="bd9c3-p108">タブのユーザー インターフェイスの拡張点のグループ。1 つのグループに、最大 6 個のコントロールを指定できます。**id** 属性が必要です。最大 125 文字の文字列です。 </span><span class="sxs-lookup"><span data-stu-id="bd9c3-p108">A group of user interface extension points on a tab. A group can have up to six controls. The **id** attribute is required. It's a string with a maximum of 125 characters.</span></span>|
|<span data-ttu-id="bd9c3-145">**Label**</span><span class="sxs-lookup"><span data-stu-id="bd9c3-145">**Label**</span></span>|<span data-ttu-id="bd9c3-146">必須。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-146">Required.</span></span> <span data-ttu-id="bd9c3-147">グループのラベルです。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-147">The label of the group.</span></span> <span data-ttu-id="bd9c3-148">**resid 属性** は 32 文字以内で **、String** 要素の **id** 属性の値に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-148">The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of a **String** element.</span></span> <span data-ttu-id="bd9c3-149">**String** 要素は、 **Resources** 要素の子要素である **ShortStrings** 要素の子要素です。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-149">The **String** element is a child element of the **ShortStrings** element, which is a child element of the **Resources** element.</span></span>|
|<span data-ttu-id="bd9c3-150">**Icon**</span><span class="sxs-lookup"><span data-stu-id="bd9c3-150">**Icon**</span></span>|<span data-ttu-id="bd9c3-151">必須。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-151">Required.</span></span> <span data-ttu-id="bd9c3-152">小さいフォーム ファクターのデバイス、または表示されるボタンが多すぎるときに使用されるグループのアイコンを指定します。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-152">Specifies the group's icon to be used on small form factor devices, or when too many buttons are displayed.</span></span> <span data-ttu-id="bd9c3-153">**resid 属性** は 32 文字以内で **、Image** 要素の **id** 属性の値に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-153">The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of an **Image** element.</span></span> <span data-ttu-id="bd9c3-154">**Image** 要素は、 **Resources** 要素の子要素である **Images** 要素の子要素です。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-154">The **Image** element is a child element of the **Images** element, which is a child element of the **Resources** element.</span></span> <span data-ttu-id="bd9c3-155">**size** 属性は、イメージのサイズをピクセル単位で指定します。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-155">The **size** attribute gives the size, in pixels, of the image.</span></span> <span data-ttu-id="bd9c3-156">3 つのイメージのサイズ (16、32、80) が必要です。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-156">Three image sizes are required: 16, 32, and 80.</span></span> <span data-ttu-id="bd9c3-157">5 つのオプションのサイズ (20、24、40、48、64) もサポートされています。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-157">Five optional sizes are also supported: 20, 24, 40, 48, and 64.</span></span>|
|<span data-ttu-id="bd9c3-158">**Tooltip**</span><span class="sxs-lookup"><span data-stu-id="bd9c3-158">**Tooltip**</span></span>|<span data-ttu-id="bd9c3-159">省略可能。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-159">Optional.</span></span> <span data-ttu-id="bd9c3-160">グループのツールヒント。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-160">The tooltip of the group.</span></span> <span data-ttu-id="bd9c3-161">**resid 属性** は 32 文字以内で **、String** 要素の **id** 属性の値に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-161">The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of a **String** element.</span></span> <span data-ttu-id="bd9c3-162">**String** 要素は、 **Resources** 要素の子要素である **LongStrings** 要素の子要素です。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-162">The **String** element is a child element of the **LongStrings** element, which is a child element of the **Resources** element.</span></span>|
|<span data-ttu-id="bd9c3-163">**Control**</span><span class="sxs-lookup"><span data-stu-id="bd9c3-163">**Control**</span></span>|<span data-ttu-id="bd9c3-164">各グループには、少なくとも 1 つのコントロールが必要です。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-164">Each group requires at least one control.</span></span> <span data-ttu-id="bd9c3-165">**コントロール要素** には、Button または **Menu** を **指定できます**。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-165">A **Control** element can be either a **Button** or a **Menu**.</span></span> <span data-ttu-id="bd9c3-166">メニュー **を使用** して、ボタン コントロールのドロップダウン リストを指定します。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-166">Use **Menu** to specify a drop-down list of button controls.</span></span> <span data-ttu-id="bd9c3-167">現在は、ボタンとメニューのみがサポートされています。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-167">Currently, only buttons and menus are supported.</span></span> <span data-ttu-id="bd9c3-168">詳細については、「[Button コントロール](control.md#button-control)」および「[Menu コントロール](control.md#menu-dropdown-button-controls)」のセクションを参照してください。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-168">See the [Button controls](control.md#button-control) and [Menu controls](control.md#menu-dropdown-button-controls) sections for more information.</span></span><br/><span data-ttu-id="bd9c3-169">**注:**  トラブルシューティングを容易にするために **、Control** 要素と関連する **Resources** 子要素を一度に 1 つ追加することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-169">**Note:**  To make troubleshooting easier, we recommend that a **Control** element and the related **Resources** child elements be added one at a time.</span></span>|
|<span data-ttu-id="bd9c3-170">**スクリプト**</span><span class="sxs-lookup"><span data-stu-id="bd9c3-170">**Script**</span></span>|<span data-ttu-id="bd9c3-171">カスタム関数の定義と登録コードを含む JavaScript ファイルにリンクします。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-171">Links to the JavaScript file with the custom function definition and registration code.</span></span> <span data-ttu-id="bd9c3-172">Developer Preview では、この要素は使用しません。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-172">This element is not used in the Developer Preview.</span></span> <span data-ttu-id="bd9c3-173">代わりに、HTML ページはすべての JavaScript ファイルを読み込みます。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-173">Instead, the HTML page is responsible for loading all JavaScript files.</span></span>|
|<span data-ttu-id="bd9c3-174">**Page**</span><span class="sxs-lookup"><span data-stu-id="bd9c3-174">**Page**</span></span>|<span data-ttu-id="bd9c3-175">カスタム関数についての HTML ページにリンクします。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-175">Links to the HTML page for your custom functions.</span></span>|

## <a name="extension-points-for-outlook"></a><span data-ttu-id="bd9c3-176">Outlook のみの拡張点</span><span class="sxs-lookup"><span data-stu-id="bd9c3-176">Extension points for Outlook</span></span>

- [<span data-ttu-id="bd9c3-177">MessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="bd9c3-177">MessageReadCommandSurface</span></span>](#messagereadcommandsurface)
- [<span data-ttu-id="bd9c3-178">MessageComposeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="bd9c3-178">MessageComposeCommandSurface</span></span>](#messagecomposecommandsurface)
- [<span data-ttu-id="bd9c3-179">AppointmentOrganizerCommandSurface</span><span class="sxs-lookup"><span data-stu-id="bd9c3-179">AppointmentOrganizerCommandSurface</span></span>](#appointmentorganizercommandsurface)
- [<span data-ttu-id="bd9c3-180">AppointmentAttendeeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="bd9c3-180">AppointmentAttendeeCommandSurface</span></span>](#appointmentattendeecommandsurface)
- <span data-ttu-id="bd9c3-181">[Module](#module) ([DesktopFormFactor](desktopformfactor.md) でのみ使用できます。)</span><span class="sxs-lookup"><span data-stu-id="bd9c3-181">[Module](#module) (Can only be used in the [DesktopFormFactor](desktopformfactor.md).)</span></span>
- [<span data-ttu-id="bd9c3-182">MobileMessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="bd9c3-182">MobileMessageReadCommandSurface</span></span>](#mobilemessagereadcommandsurface)
- [<span data-ttu-id="bd9c3-183">MobileOnlineMeetingCommandSurface</span><span class="sxs-lookup"><span data-stu-id="bd9c3-183">MobileOnlineMeetingCommandSurface</span></span>](#mobileonlinemeetingcommandsurface)
- [<span data-ttu-id="bd9c3-184">LaunchEvent</span><span class="sxs-lookup"><span data-stu-id="bd9c3-184">LaunchEvent</span></span>](#launchevent-preview)
- [<span data-ttu-id="bd9c3-185">Events</span><span class="sxs-lookup"><span data-stu-id="bd9c3-185">Events</span></span>](#events)
- [<span data-ttu-id="bd9c3-186">DetectedEntity</span><span class="sxs-lookup"><span data-stu-id="bd9c3-186">DetectedEntity</span></span>](#detectedentity)

### <a name="messagereadcommandsurface"></a><span data-ttu-id="bd9c3-187">MessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="bd9c3-187">MessageReadCommandSurface</span></span>

<span data-ttu-id="bd9c3-p114">この拡張点により、メールの閲覧ビューのコマンド サーフェスにボタンが配置されます。Outlook デスクトップでは、これはリボンに表示されます。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-p114">This extension point puts buttons in the command surface for the mail read view. In Outlook desktop, this appears in the ribbon.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="bd9c3-190">子要素</span><span class="sxs-lookup"><span data-stu-id="bd9c3-190">Child elements</span></span>

|  <span data-ttu-id="bd9c3-191">要素</span><span class="sxs-lookup"><span data-stu-id="bd9c3-191">Element</span></span> |  <span data-ttu-id="bd9c3-192">説明</span><span class="sxs-lookup"><span data-stu-id="bd9c3-192">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="bd9c3-193">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="bd9c3-193">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="bd9c3-194">コマンドを既定のリボン タブに追加します。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-194">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="bd9c3-195">CustomTab</span><span class="sxs-lookup"><span data-stu-id="bd9c3-195">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="bd9c3-196">コマンドをカスタム リボン タブに追加します。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-196">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="bd9c3-197">OfficeTab の例</span><span class="sxs-lookup"><span data-stu-id="bd9c3-197">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="bd9c3-198">CustomTab の例</span><span class="sxs-lookup"><span data-stu-id="bd9c3-198">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="messagecomposecommandsurface"></a><span data-ttu-id="bd9c3-199">MessageComposeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="bd9c3-199">MessageComposeCommandSurface</span></span>

<span data-ttu-id="bd9c3-200">この拡張点は、メールの新規作成フォームを使用してアドイン用のリボンにボタンを配置します。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-200">This extension point puts buttons on the ribbon for add-ins using mail compose form.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="bd9c3-201">子要素</span><span class="sxs-lookup"><span data-stu-id="bd9c3-201">Child elements</span></span>

|  <span data-ttu-id="bd9c3-202">要素</span><span class="sxs-lookup"><span data-stu-id="bd9c3-202">Element</span></span> |  <span data-ttu-id="bd9c3-203">説明</span><span class="sxs-lookup"><span data-stu-id="bd9c3-203">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="bd9c3-204">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="bd9c3-204">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="bd9c3-205">コマンドを既定のリボン タブに追加します。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-205">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="bd9c3-206">CustomTab</span><span class="sxs-lookup"><span data-stu-id="bd9c3-206">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="bd9c3-207">コマンドをカスタム リボン タブに追加します。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-207">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="bd9c3-208">OfficeTab の例</span><span class="sxs-lookup"><span data-stu-id="bd9c3-208">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="bd9c3-209">CustomTab の例</span><span class="sxs-lookup"><span data-stu-id="bd9c3-209">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="appointmentorganizercommandsurface"></a><span data-ttu-id="bd9c3-210">AppointmentOrganizerCommandSurface</span><span class="sxs-lookup"><span data-stu-id="bd9c3-210">AppointmentOrganizerCommandSurface</span></span>

<span data-ttu-id="bd9c3-211">この拡張点は、会議の開催者に表示されるフォームのリボンにボタンを配置します。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-211">This extension point puts buttons on the ribbon for the form that's displayed to the organizer of the meeting.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="bd9c3-212">子要素</span><span class="sxs-lookup"><span data-stu-id="bd9c3-212">Child elements</span></span>

|  <span data-ttu-id="bd9c3-213">要素</span><span class="sxs-lookup"><span data-stu-id="bd9c3-213">Element</span></span> |  <span data-ttu-id="bd9c3-214">説明</span><span class="sxs-lookup"><span data-stu-id="bd9c3-214">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="bd9c3-215">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="bd9c3-215">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="bd9c3-216">コマンドを既定のリボン タブに追加します。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-216">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="bd9c3-217">CustomTab</span><span class="sxs-lookup"><span data-stu-id="bd9c3-217">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="bd9c3-218">コマンドをカスタム リボン タブに追加します。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-218">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="bd9c3-219">OfficeTab の例</span><span class="sxs-lookup"><span data-stu-id="bd9c3-219">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="bd9c3-220">CustomTab の例</span><span class="sxs-lookup"><span data-stu-id="bd9c3-220">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="appointmentattendeecommandsurface"></a><span data-ttu-id="bd9c3-221">AppointmentAttendeeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="bd9c3-221">AppointmentAttendeeCommandSurface</span></span>

<span data-ttu-id="bd9c3-222">この拡張点は、会議の出席者に表示されるフォームのリボンにボタンを配置します。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-222">This extension point puts buttons on the ribbon for the form that's displayed to the attendee of the meeting.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="bd9c3-223">子要素</span><span class="sxs-lookup"><span data-stu-id="bd9c3-223">Child elements</span></span>

|  <span data-ttu-id="bd9c3-224">要素</span><span class="sxs-lookup"><span data-stu-id="bd9c3-224">Element</span></span> |  <span data-ttu-id="bd9c3-225">説明</span><span class="sxs-lookup"><span data-stu-id="bd9c3-225">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="bd9c3-226">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="bd9c3-226">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="bd9c3-227">コマンドを既定のリボン タブに追加します。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-227">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="bd9c3-228">CustomTab</span><span class="sxs-lookup"><span data-stu-id="bd9c3-228">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="bd9c3-229">コマンドをカスタム リボン タブに追加します。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-229">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="bd9c3-230">OfficeTab の例</span><span class="sxs-lookup"><span data-stu-id="bd9c3-230">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="bd9c3-231">CustomTab の例</span><span class="sxs-lookup"><span data-stu-id="bd9c3-231">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="module"></a><span data-ttu-id="bd9c3-232">Module</span><span class="sxs-lookup"><span data-stu-id="bd9c3-232">Module</span></span>

<span data-ttu-id="bd9c3-233">この拡張点は、モジュール拡張機能用のリボンにボタンを配置します。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-233">This extension point puts buttons on the ribbon for the module extension.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="bd9c3-234">メールボックスイベント [とアイテム](../objectmodel/preview-requirement-set/office.context.mailbox.md#events) イベント [の](../objectmodel/preview-requirement-set/office.context.mailbox.item.md#events) 登録は、この拡張ポイントでは使用できません。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-234">Registering [Mailbox](../objectmodel/preview-requirement-set/office.context.mailbox.md#events) and [Item](../objectmodel/preview-requirement-set/office.context.mailbox.item.md#events) events is not available with this extension point.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="bd9c3-235">子要素</span><span class="sxs-lookup"><span data-stu-id="bd9c3-235">Child elements</span></span>

|  <span data-ttu-id="bd9c3-236">要素</span><span class="sxs-lookup"><span data-stu-id="bd9c3-236">Element</span></span> |  <span data-ttu-id="bd9c3-237">説明</span><span class="sxs-lookup"><span data-stu-id="bd9c3-237">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="bd9c3-238">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="bd9c3-238">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="bd9c3-239">コマンドを既定のリボン タブに追加します。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-239">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="bd9c3-240">CustomTab</span><span class="sxs-lookup"><span data-stu-id="bd9c3-240">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="bd9c3-241">コマンドをカスタム リボン タブに追加します。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-241">Adds the command(s) to the custom ribbon tab.</span></span>  |

### <a name="mobilemessagereadcommandsurface"></a><span data-ttu-id="bd9c3-242">MobileMessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="bd9c3-242">MobileMessageReadCommandSurface</span></span>

<span data-ttu-id="bd9c3-243">この拡張点により、モバイル フォーム ファクターのメールの閲覧ビューのコマンド領域にボタンが配置されます。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-243">This extension point puts buttons in the command surface for the mail read view in the mobile form factor.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="bd9c3-244">子要素</span><span class="sxs-lookup"><span data-stu-id="bd9c3-244">Child elements</span></span>

|  <span data-ttu-id="bd9c3-245">要素</span><span class="sxs-lookup"><span data-stu-id="bd9c3-245">Element</span></span> |  <span data-ttu-id="bd9c3-246">説明</span><span class="sxs-lookup"><span data-stu-id="bd9c3-246">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="bd9c3-247">Group</span><span class="sxs-lookup"><span data-stu-id="bd9c3-247">Group</span></span>](group.md) |  <span data-ttu-id="bd9c3-248">コマンド領域にボタンのグループを追加します。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-248">Adds a group of buttons to the command surface.</span></span>  |

<span data-ttu-id="bd9c3-249">この種類の **ExtensionPoint** 要素には子要素を 1 つだけ含めることができます (**Group** 要素)。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-249">**ExtensionPoint** elements of this type can only have one child element: a **Group** element.</span></span>

<span data-ttu-id="bd9c3-250">この拡張点に含まれる **Control** 要素の **xsi:type** 属性を `MobileButton` に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-250">**Control** elements contained in this extension point must have the **xsi:type** attribute set to `MobileButton`.</span></span>

#### <a name="example"></a><span data-ttu-id="bd9c3-251">例</span><span class="sxs-lookup"><span data-stu-id="bd9c3-251">Example</span></span>

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

### <a name="mobileonlinemeetingcommandsurface"></a><span data-ttu-id="bd9c3-252">MobileOnlineMeetingCommandSurface</span><span class="sxs-lookup"><span data-stu-id="bd9c3-252">MobileOnlineMeetingCommandSurface</span></span>

<span data-ttu-id="bd9c3-253">この拡張ポイントは、モバイル フォーム ファクターの予定のコマンド 画面にモードに適したトグルを設定します。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-253">This extension point puts a mode-appropriate toggle in the command surface for an appointment in the mobile form factor.</span></span> <span data-ttu-id="bd9c3-254">会議の開催者は、オンライン会議を作成できます。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-254">A meeting organizer can create an online meeting.</span></span> <span data-ttu-id="bd9c3-255">その後、出席者はオンライン会議に参加できます。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-255">An attendee can subsequently join the online meeting.</span></span> <span data-ttu-id="bd9c3-256">このシナリオの詳細については、「オンライン会議プロバイダーの Outlook モバイル アドインを作成する」 [の記事を参照](../../outlook/online-meeting.md) してください。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-256">To learn more about this scenario, see the [Create an Outlook mobile add-in for an online-meeting provider](../../outlook/online-meeting.md) article.</span></span>

> [!NOTE]
> <span data-ttu-id="bd9c3-257">この拡張ポイントは、Microsoft 365 サブスクリプションを持つ Android および iOS でのみサポートされます。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-257">This extension point is only supported on Android and iOS with a Microsoft 365 subscription.</span></span>
>
> <span data-ttu-id="bd9c3-258">メールボックスイベント [とアイテム](../objectmodel/preview-requirement-set/office.context.mailbox.md#events) イベント [の](../objectmodel/preview-requirement-set/office.context.mailbox.item.md#events) 登録は、この拡張ポイントでは使用できません。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-258">Registering [Mailbox](../objectmodel/preview-requirement-set/office.context.mailbox.md#events) and [Item](../objectmodel/preview-requirement-set/office.context.mailbox.item.md#events) events is not available with this extension point.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="bd9c3-259">子要素</span><span class="sxs-lookup"><span data-stu-id="bd9c3-259">Child elements</span></span>

|  <span data-ttu-id="bd9c3-260">要素</span><span class="sxs-lookup"><span data-stu-id="bd9c3-260">Element</span></span> |  <span data-ttu-id="bd9c3-261">説明</span><span class="sxs-lookup"><span data-stu-id="bd9c3-261">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="bd9c3-262">Control</span><span class="sxs-lookup"><span data-stu-id="bd9c3-262">Control</span></span>](control.md) |  <span data-ttu-id="bd9c3-263">コマンド 画面にボタンを追加します。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-263">Adds a button to the command surface.</span></span>  |

<span data-ttu-id="bd9c3-264">`ExtensionPoint` この型の要素は、要素という 1 つの子要素のみを持 `Control` つ場合があります。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-264">`ExtensionPoint` elements of this type can only have one child element: a `Control` element.</span></span>

<span data-ttu-id="bd9c3-265">この `Control` 拡張ポイントに含まれる要素には、属性がに `xsi:type` 設定されている必要があります `MobileButton` 。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-265">The `Control` element contained in this extension point must have the `xsi:type` attribute set to `MobileButton`.</span></span>

<span data-ttu-id="bd9c3-266">画像 `Icon` は、16 進数コードまたは他の色形式で同等の値を使用 `#919191` して [グレースケールに設定する必要があります](https://convertingcolors.com/hex-color-919191.html)。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-266">The `Icon` images should be in grayscale using hex code `#919191` or its equivalent in [other color formats](https://convertingcolors.com/hex-color-919191.html).</span></span>

#### <a name="example"></a><span data-ttu-id="bd9c3-267">例</span><span class="sxs-lookup"><span data-stu-id="bd9c3-267">Example</span></span>

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

### <a name="launchevent-preview"></a><span data-ttu-id="bd9c3-268">LaunchEvent (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="bd9c3-268">LaunchEvent (preview)</span></span>

> [!NOTE]
> <span data-ttu-id="bd9c3-269">この拡張ポイントは、Outlook on [](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) the web および Microsoft 365 サブスクリプションを使用した Windows でのプレビューでのみサポートされます。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-269">This extension point is only supported in [preview](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) in Outlook on the web and on Windows with a Microsoft 365 subscription.</span></span>

<span data-ttu-id="bd9c3-270">この拡張ポイントを使用すると、デスクトップ フォーム ファクターでサポートされているイベントに基づいてアドインをアクティブ化できます。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-270">This extension point enables an add-in to activate based on supported events in the desktop form factor.</span></span> <span data-ttu-id="bd9c3-271">現在、サポートされている唯一のイベントは `OnNewMessageCompose` 、 と です `OnNewAppointmentOrganizer` 。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-271">Currently, the only supported events are `OnNewMessageCompose` and `OnNewAppointmentOrganizer`.</span></span> <span data-ttu-id="bd9c3-272">このシナリオの詳細については、「イベント ベースのライセンス認証用に Outlook アドインを構成 [する」を参照](../../outlook/autolaunch.md) してください。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-272">To learn more about this scenario, see the [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md) article.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="bd9c3-273">メールボックスイベント [とアイテム](../objectmodel/preview-requirement-set/office.context.mailbox.md#events) イベント [の](../objectmodel/preview-requirement-set/office.context.mailbox.item.md#events) 登録は、この拡張ポイントでは使用できません。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-273">Registering [Mailbox](../objectmodel/preview-requirement-set/office.context.mailbox.md#events) and [Item](../objectmodel/preview-requirement-set/office.context.mailbox.item.md#events) events is not available with this extension point.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="bd9c3-274">子要素</span><span class="sxs-lookup"><span data-stu-id="bd9c3-274">Child elements</span></span>

|  <span data-ttu-id="bd9c3-275">要素</span><span class="sxs-lookup"><span data-stu-id="bd9c3-275">Element</span></span> |  <span data-ttu-id="bd9c3-276">説明</span><span class="sxs-lookup"><span data-stu-id="bd9c3-276">Description</span></span>  |
|:-----|:-----|
| [<span data-ttu-id="bd9c3-277">LaunchEvents</span><span class="sxs-lookup"><span data-stu-id="bd9c3-277">LaunchEvents</span></span>](launchevents.md) |  <span data-ttu-id="bd9c3-278">イベント ベース [のアクティブ化の LaunchEvent](launchevent.md) の一覧。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-278">List of [LaunchEvent](launchevent.md) for event-based activation.</span></span>  |
| [<span data-ttu-id="bd9c3-279">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="bd9c3-279">SourceLocation</span></span>](sourcelocation.md) |  <span data-ttu-id="bd9c3-280">ソース JavaScript ファイルの場所。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-280">The location of the source JavaScript file.</span></span>  |

#### <a name="example"></a><span data-ttu-id="bd9c3-281">例</span><span class="sxs-lookup"><span data-stu-id="bd9c3-281">Example</span></span>

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

### <a name="events"></a><span data-ttu-id="bd9c3-282">Events</span><span class="sxs-lookup"><span data-stu-id="bd9c3-282">Events</span></span>

<span data-ttu-id="bd9c3-283">この拡張点は、指定したイベントのイベント ハンドラーを追加します。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-283">This extension point adds an event handler for a specified event.</span></span> <span data-ttu-id="bd9c3-284">この拡張ポイントの使用の詳細については、「Outlook アドインの送信時機能 [」を参照してください](../../outlook/outlook-on-send-addins.md)。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-284">For more information about using this extension point, see [On-send feature for Outlook add-ins](../../outlook/outlook-on-send-addins.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="bd9c3-285">メールボックスイベント [とアイテム](../objectmodel/preview-requirement-set/office.context.mailbox.md#events) イベント [の](../objectmodel/preview-requirement-set/office.context.mailbox.item.md#events) 登録は、この拡張ポイントでは使用できません。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-285">Registering [Mailbox](../objectmodel/preview-requirement-set/office.context.mailbox.md#events) and [Item](../objectmodel/preview-requirement-set/office.context.mailbox.item.md#events) events is not available with this extension point.</span></span>

| <span data-ttu-id="bd9c3-286">要素</span><span class="sxs-lookup"><span data-stu-id="bd9c3-286">Element</span></span> | <span data-ttu-id="bd9c3-287">説明</span><span class="sxs-lookup"><span data-stu-id="bd9c3-287">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="bd9c3-288">Event</span><span class="sxs-lookup"><span data-stu-id="bd9c3-288">Event</span></span>](event.md) |  <span data-ttu-id="bd9c3-289">イベントとイベント ハンドラーの関数を指定します。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-289">Specifies the event and event handler function.</span></span>  |

#### <a name="itemsend-event-example"></a><span data-ttu-id="bd9c3-290">ItemSend イベントの例</span><span class="sxs-lookup"><span data-stu-id="bd9c3-290">ItemSend event example</span></span>

```xml
<ExtensionPoint xsi:type="Events">
  <Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="itemSendHandler" />
</ExtensionPoint>
```

### <a name="detectedentity"></a><span data-ttu-id="bd9c3-291">DetectedEntity</span><span class="sxs-lookup"><span data-stu-id="bd9c3-291">DetectedEntity</span></span>

<span data-ttu-id="bd9c3-292">この拡張点は、指定したエンティティの種類に対するコンテキスト アドインのアクティブ化を追加します。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-292">This extension point adds a contextual add-in activation on a specified entity type.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="bd9c3-293">メールボックスイベント [とアイテム](../objectmodel/preview-requirement-set/office.context.mailbox.md#events) イベント [の](../objectmodel/preview-requirement-set/office.context.mailbox.item.md#events) 登録は、この拡張ポイントでは使用できません。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-293">Registering [Mailbox](../objectmodel/preview-requirement-set/office.context.mailbox.md#events) and [Item](../objectmodel/preview-requirement-set/office.context.mailbox.item.md#events) events is not available with this extension point.</span></span>

<span data-ttu-id="bd9c3-294">これを収容している [VersionOverrides](versionoverrides.md) 要素は、`xsi:type` 属性の値が `VersionOverridesV1_1` になっている必要があります。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-294">The containing [VersionOverrides](versionoverrides.md) element must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span>

> [!NOTE]
> <span data-ttu-id="bd9c3-295">この要素の種類は、[要件セット 1.6 以降をサポートする Outlook クライアント ](../requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)が利用できます。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-295">This element type is available to [Outlook clients that support requirement sets 1.6 and later](../requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients).</span></span>

|  <span data-ttu-id="bd9c3-296">要素</span><span class="sxs-lookup"><span data-stu-id="bd9c3-296">Element</span></span> |  <span data-ttu-id="bd9c3-297">説明</span><span class="sxs-lookup"><span data-stu-id="bd9c3-297">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="bd9c3-298">Label</span><span class="sxs-lookup"><span data-stu-id="bd9c3-298">Label</span></span>](#label) |  <span data-ttu-id="bd9c3-299">アドインのコンテキスト ウィンドウのラベルを指定します。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-299">Specifies the label for the add-in in the contextual window.</span></span>  |
|  [<span data-ttu-id="bd9c3-300">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="bd9c3-300">SourceLocation</span></span>](sourcelocation.md) |  <span data-ttu-id="bd9c3-301">コンテキスト ウィンドウの URL を指定します。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-301">Specifies the URL for the contextual window.</span></span>  |
|  [<span data-ttu-id="bd9c3-302">Rule</span><span class="sxs-lookup"><span data-stu-id="bd9c3-302">Rule</span></span>](rule.md) |  <span data-ttu-id="bd9c3-303">アドインをアクティブ化するタイミングを決定する 1 つ以上のルールを指定します。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-303">Specifies the rule or rules that determine when an add-in activates.</span></span>  |

#### <a name="label"></a><span data-ttu-id="bd9c3-304">Label</span><span class="sxs-lookup"><span data-stu-id="bd9c3-304">Label</span></span>

<span data-ttu-id="bd9c3-305">必ず指定します。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-305">Required.</span></span> <span data-ttu-id="bd9c3-306">グループのラベルです。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-306">The label of the group.</span></span> <span data-ttu-id="bd9c3-307">**resid 属性** は 32 文字以内で、Resources 要素の **ShortStrings** 要素の **String** 要素の **id** 属性の値に設定 [する必要](resources.md)があります。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-307">The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

#### <a name="highlight-requirements"></a><span data-ttu-id="bd9c3-308">強調表示の要件</span><span class="sxs-lookup"><span data-stu-id="bd9c3-308">Highlight requirements</span></span>

<span data-ttu-id="bd9c3-p119">ユーザーは、強調表示されたエンティティに対話型の操作を実行する方法でのみコンテキスト アドインを有効化できます。開発者は、`ItemHasKnownEntity` および `ItemHasRegularExpressionMatch` のルールの種類に対応する `Rule` 要素の `Highlight` 属性を使用して、強調表示にするエンティティを制御します。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-p119">The only way a user can activate a contextual add-in is to interact with a highlighted entity. Developers can control which entities are highlighted by using the `Highlight` attribute of the `Rule` element for `ItemHasKnownEntity` and `ItemHasRegularExpressionMatch` rule types.</span></span>

<span data-ttu-id="bd9c3-p120">ただし、注意する必要のある制限があります。これらの制限は、ユーザーにアドインをアクティブ化する方法を提供するために、適用可能なメッセージや予定で強調表示されたエンティティが常に存在するようにするために実施されます。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-p120">However, there are some limitations to be aware of. These limitations are in place to ensure that there will always be a highlighted entity in applicable messages or appointments to give the user a way to activate the add-in.</span></span>

- <span data-ttu-id="bd9c3-313">`EmailAddress` および `Url` のエンティティの種類は、強調表示できません。そのため、アドインをアクティブ化するためには使用できません。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-313">The `EmailAddress` and `Url` entity types cannot be highlighted, and therefore cannot be used to activate an add-in.</span></span>
- <span data-ttu-id="bd9c3-314">単一のルールを使用する場合、`Highlight` は `all` に設定されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-314">If using a single rule, `Highlight` MUST be set to `all`.</span></span>
- <span data-ttu-id="bd9c3-315">複数のルールを組み合わせるために `Mode="AND"` で `RuleCollection` のルールの種類を使用する場合は、少なくとも 1 つのルールの `Highlight` が `all` に設定されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-315">If using a `RuleCollection` rule type with `Mode="AND"` to combine multiple rules, at least one of the rules MUST have `Highlight` set to `all`.</span></span>
- <span data-ttu-id="bd9c3-316">複数のルールを組み合わせるために `Mode="OR"` で `RuleCollection` のルールの種類を使用する場合は、すべてのルールの `Highlight` が `all` に設定されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="bd9c3-316">If using a `RuleCollection` rule type with `Mode="OR"` to combine multiple rules, all of the rules MUST have `Highlight` set to `all`.</span></span>

#### <a name="detectedentity-event-example"></a><span data-ttu-id="bd9c3-317">DetectedEntity イベントの例</span><span class="sxs-lookup"><span data-stu-id="bd9c3-317">DetectedEntity event example</span></span>

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
