---
title: マニフェスト ファイルの ExtensionPoint 要素
description: Office UI でアドインが機能を公開する場所を定義します。
ms.date: 01/22/2021
localization_priority: Normal
ms.openlocfilehash: 96bf3a6835b1a0ab6e5aa85a837515a3071e5610
ms.sourcegitcommit: 6c5716d92312887e3d944bf12d9985560109b3c0
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/23/2021
ms.locfileid: "49944306"
---
# <a name="extensionpoint-element"></a><span data-ttu-id="1f970-103">ExtensionPoint 要素</span><span class="sxs-lookup"><span data-stu-id="1f970-103">ExtensionPoint element</span></span>

 <span data-ttu-id="1f970-104">Office UI でアドインが機能を公開する場所を定義します。</span><span class="sxs-lookup"><span data-stu-id="1f970-104">Defines where an add-in exposes functionality in the Office UI.</span></span> <span data-ttu-id="1f970-105">**ExtensionPoint** 要素は、[AllFormFactors](allformfactors.md)、[DesktopFormFactor](desktopformfactor.md)、[MobileFormFactor](mobileformfactor.md) の子要素です。</span><span class="sxs-lookup"><span data-stu-id="1f970-105">The **ExtensionPoint** element is a child element of [AllFormFactors](allformfactors.md), [DesktopFormFactor](desktopformfactor.md) or [MobileFormFactor](mobileformfactor.md).</span></span>

## <a name="attributes"></a><span data-ttu-id="1f970-106">属性</span><span class="sxs-lookup"><span data-stu-id="1f970-106">Attributes</span></span>

|  <span data-ttu-id="1f970-107">属性</span><span class="sxs-lookup"><span data-stu-id="1f970-107">Attribute</span></span>  |  <span data-ttu-id="1f970-108">必須</span><span class="sxs-lookup"><span data-stu-id="1f970-108">Required</span></span>  |  <span data-ttu-id="1f970-109">説明</span><span class="sxs-lookup"><span data-stu-id="1f970-109">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="1f970-110">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="1f970-110">**xsi:type**</span></span>  |  <span data-ttu-id="1f970-111">はい</span><span class="sxs-lookup"><span data-stu-id="1f970-111">Yes</span></span>  | <span data-ttu-id="1f970-112">定義される拡張点の種類。</span><span class="sxs-lookup"><span data-stu-id="1f970-112">The type of extension point being defined.</span></span>|

## <a name="extension-points-for-excel-only"></a><span data-ttu-id="1f970-113">Excel のみの拡張点</span><span class="sxs-lookup"><span data-stu-id="1f970-113">Extension points for Excel only</span></span>

- <span data-ttu-id="1f970-114">**CustomFunctions** - Excel 向けの JavaScript で記述されたカスタム関数。</span><span class="sxs-lookup"><span data-stu-id="1f970-114">**CustomFunctions** - A custom function written in JavaScript for Excel.</span></span>

<span data-ttu-id="1f970-115">[この XML コード サンプル](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/manifest.xml)は、**CustomFunctions** 属性の値を持つ **ExtensionPoint** 要素を使用する方法と、使用する子要素を示しています。</span><span class="sxs-lookup"><span data-stu-id="1f970-115">[This XML code sample](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/manifest.xml) shows how to use the **ExtensionPoint** element with the **CustomFunctions** attribute value, and the child elements to be used.</span></span>

## <a name="extension-points-for-word-excel-powerpoint-and-onenote-add-in-commands"></a><span data-ttu-id="1f970-116">Word、Excel、PowerPoint、OneNote アドイン コマンドの拡張点</span><span class="sxs-lookup"><span data-stu-id="1f970-116">Extension points for Word, Excel, PowerPoint, and OneNote add-in commands</span></span>

- <span data-ttu-id="1f970-117">**PrimaryCommandSurface** - Office のリボン。</span><span class="sxs-lookup"><span data-stu-id="1f970-117">**PrimaryCommandSurface** - The ribbon in Office.</span></span>
- <span data-ttu-id="1f970-118">**ContextMenu** Office UI で右クリックしたときに表示されるショートカット メニュー。</span><span class="sxs-lookup"><span data-stu-id="1f970-118">**ContextMenu** - The shortcut menu that appears when you right-click in the Office UI.</span></span>

<span data-ttu-id="1f970-119">次の例は、**PrimaryCommandSurface** と **ContextMenu** の属性値を持つ **ExtensionPoint** 要素を使用する方法と、各要素と併用する必要がある子要素を示しています。</span><span class="sxs-lookup"><span data-stu-id="1f970-119">The following examples show how to use the **ExtensionPoint** element with **PrimaryCommandSurface** and **ContextMenu** attribute values, and the child elements that should be used with each.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="1f970-p102">ID 属性を含む要素では、一意の ID を指定してください。会社の名前と ID を使用することをお勧めします。たとえば、次の形式にします。<CustomTab id="mycompanyname.mygroupname"></span><span class="sxs-lookup"><span data-stu-id="1f970-p102">For elements that contain an ID attribute, make sure you provide a unique ID. We recommend that you use your company's name along with your ID. For example, use the following format. <CustomTab id="mycompanyname.mygroupname"></span></span>

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

#### <a name="child-elements"></a><span data-ttu-id="1f970-123">子要素</span><span class="sxs-lookup"><span data-stu-id="1f970-123">Child elements</span></span>
 
|<span data-ttu-id="1f970-124">要素</span><span class="sxs-lookup"><span data-stu-id="1f970-124">Element</span></span>|<span data-ttu-id="1f970-125">説明</span><span class="sxs-lookup"><span data-stu-id="1f970-125">Description</span></span>|
|:-----|:-----|
|<span data-ttu-id="1f970-126">**CustomTab**</span><span class="sxs-lookup"><span data-stu-id="1f970-126">**CustomTab**</span></span>|<span data-ttu-id="1f970-p103">カスタム タブをリボンに追加する必要がある場合は必須 (**PrimaryCommandSurface** を使用)。**CustomTab** 要素を使用する場合、**OfficeTab** 要素は使用できません。**id** 属性が必要です。 </span><span class="sxs-lookup"><span data-stu-id="1f970-p103">Required if you want to add a custom tab to the ribbon (using **PrimaryCommandSurface**). If you use the **CustomTab** element, you can't use the **OfficeTab** element. The **id** attribute is required.</span></span>|
|<span data-ttu-id="1f970-130">**OfficeTab**</span><span class="sxs-lookup"><span data-stu-id="1f970-130">**OfficeTab**</span></span>|<span data-ttu-id="1f970-131">**(PrimaryCommandSurface** を使用して) アプリリボン タブOfficeを拡張する場合は必須です。</span><span class="sxs-lookup"><span data-stu-id="1f970-131">Required if you want to extend a default Office app ribbon tab (using **PrimaryCommandSurface**).</span></span> <span data-ttu-id="1f970-132">OfficeTab 要素 **を使用** する場合 **、CustomTab 要素は使用** することはできません。</span><span class="sxs-lookup"><span data-stu-id="1f970-132">If you use the **OfficeTab** element, you can't use the **CustomTab** element.</span></span> <span data-ttu-id="1f970-133">詳細については、「[OfficeTab](officetab.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="1f970-133">For details, see [OfficeTab](officetab.md).</span></span>|
|<span data-ttu-id="1f970-134">**OfficeMenu**</span><span class="sxs-lookup"><span data-stu-id="1f970-134">**OfficeMenu**</span></span>|<span data-ttu-id="1f970-p105">既定のコンテキスト メニューにアドイン コマンドを追加する場合は必須 (**ContextMenu** を使用)。**id** 属性は以下に設定する必要があります。 </span><span class="sxs-lookup"><span data-stu-id="1f970-p105">Required if you're adding add-in commands to a default context menu (using **ContextMenu**). The **id** attribute must be set to: </span></span><br/> <span data-ttu-id="1f970-p106">Excel または Word の場合は - **ContextMenuText**。テキストが選択され、ユーザーが選択されたテキストを右クリックしたときに、コンテキスト メニューに項目が表示されます。 </span><span class="sxs-lookup"><span data-stu-id="1f970-p106">- **ContextMenuText** for Excel or Word. Displays the item on the context menu when text is selected and then the user right-clicks on the selected text. </span></span><br/> <span data-ttu-id="1f970-p107">Excel の場合は - **ContextMenuCell**。ユーザーがスプレッドシートのセルを右クリックすると、コンテキスト メニューに項目が表示されます。</span><span class="sxs-lookup"><span data-stu-id="1f970-p107">- **ContextMenuCell** for Excel. Displays the  item on the context menu when the user right-clicks on a cell on the spreadsheet.</span></span>|
|<span data-ttu-id="1f970-141">**Group**</span><span class="sxs-lookup"><span data-stu-id="1f970-141">**Group**</span></span>|<span data-ttu-id="1f970-p108">タブのユーザー インターフェイスの拡張点のグループ。1 つのグループに、最大 6 個のコントロールを指定できます。**id** 属性が必要です。最大 125 文字の文字列です。 </span><span class="sxs-lookup"><span data-stu-id="1f970-p108">A group of user interface extension points on a tab. A group can have up to six controls. The **id** attribute is required. It's a string with a maximum of 125 characters.</span></span>|
|<span data-ttu-id="1f970-145">**Label**</span><span class="sxs-lookup"><span data-stu-id="1f970-145">**Label**</span></span>|<span data-ttu-id="1f970-146">必須。</span><span class="sxs-lookup"><span data-stu-id="1f970-146">Required.</span></span> <span data-ttu-id="1f970-147">グループのラベルです。</span><span class="sxs-lookup"><span data-stu-id="1f970-147">The label of the group.</span></span> <span data-ttu-id="1f970-148">**resid 属性** は 32 文字以内で **、String** 要素の **id** 属性の値に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="1f970-148">The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of a **String** element.</span></span> <span data-ttu-id="1f970-149">**String** 要素は、 **Resources** 要素の子要素である **ShortStrings** 要素の子要素です。</span><span class="sxs-lookup"><span data-stu-id="1f970-149">The **String** element is a child element of the **ShortStrings** element, which is a child element of the **Resources** element.</span></span>|
|<span data-ttu-id="1f970-150">**Icon**</span><span class="sxs-lookup"><span data-stu-id="1f970-150">**Icon**</span></span>|<span data-ttu-id="1f970-151">必須。</span><span class="sxs-lookup"><span data-stu-id="1f970-151">Required.</span></span> <span data-ttu-id="1f970-152">小さいフォーム ファクターのデバイス、または表示されるボタンが多すぎるときに使用されるグループのアイコンを指定します。</span><span class="sxs-lookup"><span data-stu-id="1f970-152">Specifies the group's icon to be used on small form factor devices, or when too many buttons are displayed.</span></span> <span data-ttu-id="1f970-153">**resid 属性** は 32 文字以内で **、Image** 要素の **id** 属性の値に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="1f970-153">The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of an **Image** element.</span></span> <span data-ttu-id="1f970-154">**Image** 要素は、 **Resources** 要素の子要素である **Images** 要素の子要素です。</span><span class="sxs-lookup"><span data-stu-id="1f970-154">The **Image** element is a child element of the **Images** element, which is a child element of the **Resources** element.</span></span> <span data-ttu-id="1f970-155">**size** 属性は、イメージのサイズをピクセル単位で指定します。</span><span class="sxs-lookup"><span data-stu-id="1f970-155">The **size** attribute gives the size, in pixels, of the image.</span></span> <span data-ttu-id="1f970-156">3 つのイメージのサイズ (16、32、80) が必要です。</span><span class="sxs-lookup"><span data-stu-id="1f970-156">Three image sizes are required: 16, 32, and 80.</span></span> <span data-ttu-id="1f970-157">5 つのオプションのサイズ (20、24、40、48、64) もサポートされています。</span><span class="sxs-lookup"><span data-stu-id="1f970-157">Five optional sizes are also supported: 20, 24, 40, 48, and 64.</span></span>|
|<span data-ttu-id="1f970-158">**Tooltip**</span><span class="sxs-lookup"><span data-stu-id="1f970-158">**Tooltip**</span></span>|<span data-ttu-id="1f970-159">省略可能。</span><span class="sxs-lookup"><span data-stu-id="1f970-159">Optional.</span></span> <span data-ttu-id="1f970-160">グループのツールヒント。</span><span class="sxs-lookup"><span data-stu-id="1f970-160">The tooltip of the group.</span></span> <span data-ttu-id="1f970-161">**resid 属性** は 32 文字以内で **、String** 要素の **id** 属性の値に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="1f970-161">The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of a **String** element.</span></span> <span data-ttu-id="1f970-162">**String** 要素は、 **Resources** 要素の子要素である **LongStrings** 要素の子要素です。</span><span class="sxs-lookup"><span data-stu-id="1f970-162">The **String** element is a child element of the **LongStrings** element, which is a child element of the **Resources** element.</span></span>|
|<span data-ttu-id="1f970-163">**Control**</span><span class="sxs-lookup"><span data-stu-id="1f970-163">**Control**</span></span>|<span data-ttu-id="1f970-164">各グループには、少なくとも 1 つのコントロールが必要です。</span><span class="sxs-lookup"><span data-stu-id="1f970-164">Each group requires at least one control.</span></span> <span data-ttu-id="1f970-165">コントロール **要素** には、ボタンまたは **メニュー** のいずれかを指定 **できます**。</span><span class="sxs-lookup"><span data-stu-id="1f970-165">A **Control** element can be either a **Button** or a **Menu**.</span></span> <span data-ttu-id="1f970-166">メニュー **を** 使用して、ボタン コントロールのドロップダウン リストを指定します。</span><span class="sxs-lookup"><span data-stu-id="1f970-166">Use **Menu** to specify a drop-down list of button controls.</span></span> <span data-ttu-id="1f970-167">現在は、ボタンとメニューのみがサポートされています。</span><span class="sxs-lookup"><span data-stu-id="1f970-167">Currently, only buttons and menus are supported.</span></span> <span data-ttu-id="1f970-168">詳細については、「[Button コントロール](control.md#button-control)」および「[Menu コントロール](control.md#menu-dropdown-button-controls)」のセクションを参照してください。</span><span class="sxs-lookup"><span data-stu-id="1f970-168">See the [Button controls](control.md#button-control) and [Menu controls](control.md#menu-dropdown-button-controls) sections for more information.</span></span><br/><span data-ttu-id="1f970-169">**注:**  トラブルシューティングを容易にするために **、Control** 要素と関連する **Resources** 子要素を一度に 1 つ追加することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="1f970-169">**Note:**  To make troubleshooting easier, we recommend that a **Control** element and the related **Resources** child elements be added one at a time.</span></span>|
|<span data-ttu-id="1f970-170">**スクリプト**</span><span class="sxs-lookup"><span data-stu-id="1f970-170">**Script**</span></span>|<span data-ttu-id="1f970-171">カスタム関数の定義と登録コードを含む JavaScript ファイルにリンクします。</span><span class="sxs-lookup"><span data-stu-id="1f970-171">Links to the JavaScript file with the custom function definition and registration code.</span></span> <span data-ttu-id="1f970-172">Developer Preview では、この要素は使用しません。</span><span class="sxs-lookup"><span data-stu-id="1f970-172">This element is not used in the Developer Preview.</span></span> <span data-ttu-id="1f970-173">代わりに、HTML ページはすべての JavaScript ファイルを読み込みます。</span><span class="sxs-lookup"><span data-stu-id="1f970-173">Instead, the HTML page is responsible for loading all JavaScript files.</span></span>|
|<span data-ttu-id="1f970-174">**Page**</span><span class="sxs-lookup"><span data-stu-id="1f970-174">**Page**</span></span>|<span data-ttu-id="1f970-175">カスタム関数についての HTML ページにリンクします。</span><span class="sxs-lookup"><span data-stu-id="1f970-175">Links to the HTML page for your custom functions.</span></span>|

## <a name="extension-points-for-outlook"></a><span data-ttu-id="1f970-176">Outlook のみの拡張点</span><span class="sxs-lookup"><span data-stu-id="1f970-176">Extension points for Outlook</span></span>

- [<span data-ttu-id="1f970-177">MessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="1f970-177">MessageReadCommandSurface</span></span>](#messagereadcommandsurface)
- [<span data-ttu-id="1f970-178">MessageComposeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="1f970-178">MessageComposeCommandSurface</span></span>](#messagecomposecommandsurface)
- [<span data-ttu-id="1f970-179">AppointmentOrganizerCommandSurface</span><span class="sxs-lookup"><span data-stu-id="1f970-179">AppointmentOrganizerCommandSurface</span></span>](#appointmentorganizercommandsurface)
- [<span data-ttu-id="1f970-180">AppointmentAttendeeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="1f970-180">AppointmentAttendeeCommandSurface</span></span>](#appointmentattendeecommandsurface)
- <span data-ttu-id="1f970-181">[Module](#module) ([DesktopFormFactor](desktopformfactor.md) でのみ使用できます。)</span><span class="sxs-lookup"><span data-stu-id="1f970-181">[Module](#module) (Can only be used in the [DesktopFormFactor](desktopformfactor.md).)</span></span>
- [<span data-ttu-id="1f970-182">MobileMessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="1f970-182">MobileMessageReadCommandSurface</span></span>](#mobilemessagereadcommandsurface)
- [<span data-ttu-id="1f970-183">MobileOnlineMeetingCommandSurface</span><span class="sxs-lookup"><span data-stu-id="1f970-183">MobileOnlineMeetingCommandSurface</span></span>](#mobileonlinemeetingcommandsurface)
- [<span data-ttu-id="1f970-184">LaunchEvent</span><span class="sxs-lookup"><span data-stu-id="1f970-184">LaunchEvent</span></span>](#launchevent-preview)
- [<span data-ttu-id="1f970-185">Events</span><span class="sxs-lookup"><span data-stu-id="1f970-185">Events</span></span>](#events)
- [<span data-ttu-id="1f970-186">DetectedEntity</span><span class="sxs-lookup"><span data-stu-id="1f970-186">DetectedEntity</span></span>](#detectedentity)

### <a name="messagereadcommandsurface"></a><span data-ttu-id="1f970-187">MessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="1f970-187">MessageReadCommandSurface</span></span>

<span data-ttu-id="1f970-p114">この拡張点により、メールの閲覧ビューのコマンド サーフェスにボタンが配置されます。Outlook デスクトップでは、これはリボンに表示されます。</span><span class="sxs-lookup"><span data-stu-id="1f970-p114">This extension point puts buttons in the command surface for the mail read view. In Outlook desktop, this appears in the ribbon.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="1f970-190">子要素</span><span class="sxs-lookup"><span data-stu-id="1f970-190">Child elements</span></span>

|  <span data-ttu-id="1f970-191">要素</span><span class="sxs-lookup"><span data-stu-id="1f970-191">Element</span></span> |  <span data-ttu-id="1f970-192">説明</span><span class="sxs-lookup"><span data-stu-id="1f970-192">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="1f970-193">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="1f970-193">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="1f970-194">コマンドを既定のリボン タブに追加します。</span><span class="sxs-lookup"><span data-stu-id="1f970-194">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="1f970-195">CustomTab</span><span class="sxs-lookup"><span data-stu-id="1f970-195">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="1f970-196">コマンドをカスタム リボン タブに追加します。</span><span class="sxs-lookup"><span data-stu-id="1f970-196">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="1f970-197">OfficeTab の例</span><span class="sxs-lookup"><span data-stu-id="1f970-197">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="1f970-198">CustomTab の例</span><span class="sxs-lookup"><span data-stu-id="1f970-198">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="messagecomposecommandsurface"></a><span data-ttu-id="1f970-199">MessageComposeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="1f970-199">MessageComposeCommandSurface</span></span>

<span data-ttu-id="1f970-200">この拡張点は、メールの新規作成フォームを使用してアドイン用のリボンにボタンを配置します。</span><span class="sxs-lookup"><span data-stu-id="1f970-200">This extension point puts buttons on the ribbon for add-ins using mail compose form.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="1f970-201">子要素</span><span class="sxs-lookup"><span data-stu-id="1f970-201">Child elements</span></span>

|  <span data-ttu-id="1f970-202">要素</span><span class="sxs-lookup"><span data-stu-id="1f970-202">Element</span></span> |  <span data-ttu-id="1f970-203">説明</span><span class="sxs-lookup"><span data-stu-id="1f970-203">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="1f970-204">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="1f970-204">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="1f970-205">コマンドを既定のリボン タブに追加します。</span><span class="sxs-lookup"><span data-stu-id="1f970-205">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="1f970-206">CustomTab</span><span class="sxs-lookup"><span data-stu-id="1f970-206">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="1f970-207">コマンドをカスタム リボン タブに追加します。</span><span class="sxs-lookup"><span data-stu-id="1f970-207">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="1f970-208">OfficeTab の例</span><span class="sxs-lookup"><span data-stu-id="1f970-208">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="1f970-209">CustomTab の例</span><span class="sxs-lookup"><span data-stu-id="1f970-209">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="appointmentorganizercommandsurface"></a><span data-ttu-id="1f970-210">AppointmentOrganizerCommandSurface</span><span class="sxs-lookup"><span data-stu-id="1f970-210">AppointmentOrganizerCommandSurface</span></span>

<span data-ttu-id="1f970-211">この拡張点は、会議の開催者に表示されるフォームのリボンにボタンを配置します。</span><span class="sxs-lookup"><span data-stu-id="1f970-211">This extension point puts buttons on the ribbon for the form that's displayed to the organizer of the meeting.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="1f970-212">子要素</span><span class="sxs-lookup"><span data-stu-id="1f970-212">Child elements</span></span>

|  <span data-ttu-id="1f970-213">要素</span><span class="sxs-lookup"><span data-stu-id="1f970-213">Element</span></span> |  <span data-ttu-id="1f970-214">説明</span><span class="sxs-lookup"><span data-stu-id="1f970-214">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="1f970-215">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="1f970-215">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="1f970-216">コマンドを既定のリボン タブに追加します。</span><span class="sxs-lookup"><span data-stu-id="1f970-216">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="1f970-217">CustomTab</span><span class="sxs-lookup"><span data-stu-id="1f970-217">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="1f970-218">コマンドをカスタム リボン タブに追加します。</span><span class="sxs-lookup"><span data-stu-id="1f970-218">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="1f970-219">OfficeTab の例</span><span class="sxs-lookup"><span data-stu-id="1f970-219">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="1f970-220">CustomTab の例</span><span class="sxs-lookup"><span data-stu-id="1f970-220">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="appointmentattendeecommandsurface"></a><span data-ttu-id="1f970-221">AppointmentAttendeeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="1f970-221">AppointmentAttendeeCommandSurface</span></span>

<span data-ttu-id="1f970-222">この拡張点は、会議の出席者に表示されるフォームのリボンにボタンを配置します。</span><span class="sxs-lookup"><span data-stu-id="1f970-222">This extension point puts buttons on the ribbon for the form that's displayed to the attendee of the meeting.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="1f970-223">子要素</span><span class="sxs-lookup"><span data-stu-id="1f970-223">Child elements</span></span>

|  <span data-ttu-id="1f970-224">要素</span><span class="sxs-lookup"><span data-stu-id="1f970-224">Element</span></span> |  <span data-ttu-id="1f970-225">説明</span><span class="sxs-lookup"><span data-stu-id="1f970-225">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="1f970-226">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="1f970-226">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="1f970-227">コマンドを既定のリボン タブに追加します。</span><span class="sxs-lookup"><span data-stu-id="1f970-227">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="1f970-228">CustomTab</span><span class="sxs-lookup"><span data-stu-id="1f970-228">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="1f970-229">コマンドをカスタム リボン タブに追加します。</span><span class="sxs-lookup"><span data-stu-id="1f970-229">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="1f970-230">OfficeTab の例</span><span class="sxs-lookup"><span data-stu-id="1f970-230">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="1f970-231">CustomTab の例</span><span class="sxs-lookup"><span data-stu-id="1f970-231">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="module"></a><span data-ttu-id="1f970-232">Module</span><span class="sxs-lookup"><span data-stu-id="1f970-232">Module</span></span>

<span data-ttu-id="1f970-233">この拡張点は、モジュール拡張機能用のリボンにボタンを配置します。</span><span class="sxs-lookup"><span data-stu-id="1f970-233">This extension point puts buttons on the ribbon for the module extension.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="1f970-234">子要素</span><span class="sxs-lookup"><span data-stu-id="1f970-234">Child elements</span></span>

|  <span data-ttu-id="1f970-235">要素</span><span class="sxs-lookup"><span data-stu-id="1f970-235">Element</span></span> |  <span data-ttu-id="1f970-236">説明</span><span class="sxs-lookup"><span data-stu-id="1f970-236">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="1f970-237">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="1f970-237">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="1f970-238">コマンドを既定のリボン タブに追加します。</span><span class="sxs-lookup"><span data-stu-id="1f970-238">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="1f970-239">CustomTab</span><span class="sxs-lookup"><span data-stu-id="1f970-239">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="1f970-240">コマンドをカスタム リボン タブに追加します。</span><span class="sxs-lookup"><span data-stu-id="1f970-240">Adds the command(s) to the custom ribbon tab.</span></span>  |

### <a name="mobilemessagereadcommandsurface"></a><span data-ttu-id="1f970-241">MobileMessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="1f970-241">MobileMessageReadCommandSurface</span></span>

<span data-ttu-id="1f970-242">この拡張点により、モバイル フォーム ファクターのメールの閲覧ビューのコマンド領域にボタンが配置されます。</span><span class="sxs-lookup"><span data-stu-id="1f970-242">This extension point puts buttons in the command surface for the mail read view in the mobile form factor.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="1f970-243">子要素</span><span class="sxs-lookup"><span data-stu-id="1f970-243">Child elements</span></span>

|  <span data-ttu-id="1f970-244">要素</span><span class="sxs-lookup"><span data-stu-id="1f970-244">Element</span></span> |  <span data-ttu-id="1f970-245">説明</span><span class="sxs-lookup"><span data-stu-id="1f970-245">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="1f970-246">Group</span><span class="sxs-lookup"><span data-stu-id="1f970-246">Group</span></span>](group.md) |  <span data-ttu-id="1f970-247">コマンド領域にボタンのグループを追加します。</span><span class="sxs-lookup"><span data-stu-id="1f970-247">Adds a group of buttons to the command surface.</span></span>  |

<span data-ttu-id="1f970-248">この種類の **ExtensionPoint** 要素には子要素を 1 つだけ含めることができます (**Group** 要素)。</span><span class="sxs-lookup"><span data-stu-id="1f970-248">**ExtensionPoint** elements of this type can only have one child element: a **Group** element.</span></span>

<span data-ttu-id="1f970-249">この拡張点に含まれる **Control** 要素の **xsi:type** 属性を `MobileButton` に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="1f970-249">**Control** elements contained in this extension point must have the **xsi:type** attribute set to `MobileButton`.</span></span>

#### <a name="example"></a><span data-ttu-id="1f970-250">例</span><span class="sxs-lookup"><span data-stu-id="1f970-250">Example</span></span>

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

### <a name="mobileonlinemeetingcommandsurface"></a><span data-ttu-id="1f970-251">MobileOnlineMeetingCommandSurface</span><span class="sxs-lookup"><span data-stu-id="1f970-251">MobileOnlineMeetingCommandSurface</span></span>

<span data-ttu-id="1f970-252">この拡張点は、モバイル フォーム ファクターの予定のコマンド サーフェスにモードに適したトグルを設定します。</span><span class="sxs-lookup"><span data-stu-id="1f970-252">This extension point puts a mode-appropriate toggle in the command surface for an appointment in the mobile form factor.</span></span> <span data-ttu-id="1f970-253">会議の開催者はオンライン会議を作成できます。</span><span class="sxs-lookup"><span data-stu-id="1f970-253">A meeting organizer can create an online meeting.</span></span> <span data-ttu-id="1f970-254">その後、出席者はオンライン会議に参加できます。</span><span class="sxs-lookup"><span data-stu-id="1f970-254">An attendee can subsequently join the online meeting.</span></span> <span data-ttu-id="1f970-255">このシナリオの詳細については、オンライン会議プロバイダー向け Outlook モバイル アドインの作成に関する記事 [を参照](../../outlook/online-meeting.md) してください。</span><span class="sxs-lookup"><span data-stu-id="1f970-255">To learn more about this scenario, see the [Create an Outlook mobile add-in for an online-meeting provider](../../outlook/online-meeting.md) article.</span></span>

> [!NOTE]
> <span data-ttu-id="1f970-256">この拡張点は、Microsoft 365 サブスクリプションを使用する Android でのみサポートされます。</span><span class="sxs-lookup"><span data-stu-id="1f970-256">This extension point is only supported on Android with a Microsoft 365 subscription.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="1f970-257">子要素</span><span class="sxs-lookup"><span data-stu-id="1f970-257">Child elements</span></span>

|  <span data-ttu-id="1f970-258">要素</span><span class="sxs-lookup"><span data-stu-id="1f970-258">Element</span></span> |  <span data-ttu-id="1f970-259">説明</span><span class="sxs-lookup"><span data-stu-id="1f970-259">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="1f970-260">Control</span><span class="sxs-lookup"><span data-stu-id="1f970-260">Control</span></span>](control.md) |  <span data-ttu-id="1f970-261">コマンド サーフェスにボタンを追加します。</span><span class="sxs-lookup"><span data-stu-id="1f970-261">Adds a button to the command surface.</span></span>  |

<span data-ttu-id="1f970-262">`ExtensionPoint` この型の要素は、1 つの子要素 (要素) のみを持 `Control` つ場合があります。</span><span class="sxs-lookup"><span data-stu-id="1f970-262">`ExtensionPoint` elements of this type can only have one child element: a `Control` element.</span></span>

<span data-ttu-id="1f970-263">この `Control` 拡張点に含まれる要素には、属性が設定 `xsi:type` されている必要があります `MobileButton` 。</span><span class="sxs-lookup"><span data-stu-id="1f970-263">The `Control` element contained in this extension point must have the `xsi:type` attribute set to `MobileButton`.</span></span>

<span data-ttu-id="1f970-264">イメージ `Icon` は、16 進数コードを使用してグレースケールで表示するか、他の色形式 `#919191` で同等 [の色を使用する必要があります](https://convertingcolors.com/hex-color-919191.html)。</span><span class="sxs-lookup"><span data-stu-id="1f970-264">The `Icon` images should be in grayscale using hex code `#919191` or its equivalent in [other color formats](https://convertingcolors.com/hex-color-919191.html).</span></span>

#### <a name="example"></a><span data-ttu-id="1f970-265">例</span><span class="sxs-lookup"><span data-stu-id="1f970-265">Example</span></span>

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

### <a name="launchevent-preview"></a><span data-ttu-id="1f970-266">LaunchEvent (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="1f970-266">LaunchEvent (preview)</span></span>

> [!NOTE]
> <span data-ttu-id="1f970-267">この拡張点は、Microsoft 365 サブスクリプションを使用する Outlook on the web のプレビューでのみサポートされます。 [](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md)</span><span class="sxs-lookup"><span data-stu-id="1f970-267">This extension point is only supported in [preview](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) in Outlook on the web with a Microsoft 365 subscription.</span></span>

<span data-ttu-id="1f970-268">この拡張点により、デスクトップ フォーム ファクターでサポートされているイベントに基づいてアドインをアクティブ化できます。</span><span class="sxs-lookup"><span data-stu-id="1f970-268">This extension point enables an add-in to activate based on supported events in the desktop form factor.</span></span> <span data-ttu-id="1f970-269">現在、サポートされている唯一のイベントは `OnNewMessageCompose` 次のとおりです `OnNewAppointmentOrganizer` 。</span><span class="sxs-lookup"><span data-stu-id="1f970-269">Currently, the only supported events are `OnNewMessageCompose` and `OnNewAppointmentOrganizer`.</span></span> <span data-ttu-id="1f970-270">このシナリオの詳細については、イベント ベースのアクティブ化に関する Outlook アドインの構成 [に関する記事を参照](../../outlook/autolaunch.md) してください。</span><span class="sxs-lookup"><span data-stu-id="1f970-270">To learn more about this scenario, see the [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md) article.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="1f970-271">子要素</span><span class="sxs-lookup"><span data-stu-id="1f970-271">Child elements</span></span>

|  <span data-ttu-id="1f970-272">要素</span><span class="sxs-lookup"><span data-stu-id="1f970-272">Element</span></span> |  <span data-ttu-id="1f970-273">説明</span><span class="sxs-lookup"><span data-stu-id="1f970-273">Description</span></span>  |
|:-----|:-----|
| [<span data-ttu-id="1f970-274">LaunchEvents</span><span class="sxs-lookup"><span data-stu-id="1f970-274">LaunchEvents</span></span>](launchevents.md) |  <span data-ttu-id="1f970-275">イベント ベース [のアクティブ化の LaunchEvent](launchevent.md) のリスト。</span><span class="sxs-lookup"><span data-stu-id="1f970-275">List of [LaunchEvent](launchevent.md) for event-based activation.</span></span>  |
| [<span data-ttu-id="1f970-276">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="1f970-276">SourceLocation</span></span>](sourcelocation.md) |  <span data-ttu-id="1f970-277">ソース JavaScript ファイルの場所。</span><span class="sxs-lookup"><span data-stu-id="1f970-277">The location of the source JavaScript file.</span></span>  |

#### <a name="example"></a><span data-ttu-id="1f970-278">例</span><span class="sxs-lookup"><span data-stu-id="1f970-278">Example</span></span>

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

### <a name="events"></a><span data-ttu-id="1f970-279">Events</span><span class="sxs-lookup"><span data-stu-id="1f970-279">Events</span></span>

<span data-ttu-id="1f970-280">この拡張点は、指定したイベントのイベント ハンドラーを追加します。</span><span class="sxs-lookup"><span data-stu-id="1f970-280">This extension point adds an event handler for a specified event.</span></span> <span data-ttu-id="1f970-281">この拡張点の使用の詳細については、Outlook アドインの送信時 [機能を参照してください](../../outlook/outlook-on-send-addins.md)。</span><span class="sxs-lookup"><span data-stu-id="1f970-281">For more information about using this extension point, see [On-send feature for Outlook add-ins](../../outlook/outlook-on-send-addins.md).</span></span>

| <span data-ttu-id="1f970-282">要素</span><span class="sxs-lookup"><span data-stu-id="1f970-282">Element</span></span> | <span data-ttu-id="1f970-283">説明</span><span class="sxs-lookup"><span data-stu-id="1f970-283">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="1f970-284">Event</span><span class="sxs-lookup"><span data-stu-id="1f970-284">Event</span></span>](event.md) |  <span data-ttu-id="1f970-285">イベントとイベント ハンドラーの関数を指定します。</span><span class="sxs-lookup"><span data-stu-id="1f970-285">Specifies the event and event handler function.</span></span>  |

#### <a name="itemsend-event-example"></a><span data-ttu-id="1f970-286">ItemSend イベントの例</span><span class="sxs-lookup"><span data-stu-id="1f970-286">ItemSend event example</span></span>

```xml
<ExtensionPoint xsi:type="Events">
  <Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="itemSendHandler" />
</ExtensionPoint>
```

### <a name="detectedentity"></a><span data-ttu-id="1f970-287">DetectedEntity</span><span class="sxs-lookup"><span data-stu-id="1f970-287">DetectedEntity</span></span>

<span data-ttu-id="1f970-288">この拡張点は、指定したエンティティの種類に対するコンテキスト アドインのアクティブ化を追加します。</span><span class="sxs-lookup"><span data-stu-id="1f970-288">This extension point adds a contextual add-in activation on a specified entity type.</span></span>

<span data-ttu-id="1f970-289">これを収容している [VersionOverrides](versionoverrides.md) 要素は、`xsi:type` 属性の値が `VersionOverridesV1_1` になっている必要があります。</span><span class="sxs-lookup"><span data-stu-id="1f970-289">The containing [VersionOverrides](versionoverrides.md) element must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span>

> [!NOTE]
> <span data-ttu-id="1f970-290">この要素の種類は、[要件セット 1.6 以降をサポートする Outlook クライアント ](../requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)が利用できます。</span><span class="sxs-lookup"><span data-stu-id="1f970-290">This element type is available to [Outlook clients that support requirement sets 1.6 and later](../requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients).</span></span>

|  <span data-ttu-id="1f970-291">要素</span><span class="sxs-lookup"><span data-stu-id="1f970-291">Element</span></span> |  <span data-ttu-id="1f970-292">説明</span><span class="sxs-lookup"><span data-stu-id="1f970-292">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="1f970-293">Label</span><span class="sxs-lookup"><span data-stu-id="1f970-293">Label</span></span>](#label) |  <span data-ttu-id="1f970-294">アドインのコンテキスト ウィンドウのラベルを指定します。</span><span class="sxs-lookup"><span data-stu-id="1f970-294">Specifies the label for the add-in in the contextual window.</span></span>  |
|  [<span data-ttu-id="1f970-295">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="1f970-295">SourceLocation</span></span>](sourcelocation.md) |  <span data-ttu-id="1f970-296">コンテキスト ウィンドウの URL を指定します。</span><span class="sxs-lookup"><span data-stu-id="1f970-296">Specifies the URL for the contextual window.</span></span>  |
|  [<span data-ttu-id="1f970-297">Rule</span><span class="sxs-lookup"><span data-stu-id="1f970-297">Rule</span></span>](rule.md) |  <span data-ttu-id="1f970-298">アドインをアクティブ化するタイミングを決定する 1 つ以上のルールを指定します。</span><span class="sxs-lookup"><span data-stu-id="1f970-298">Specifies the rule or rules that determine when an add-in activates.</span></span>  |

#### <a name="label"></a><span data-ttu-id="1f970-299">Label</span><span class="sxs-lookup"><span data-stu-id="1f970-299">Label</span></span>

<span data-ttu-id="1f970-300">必ず指定します。</span><span class="sxs-lookup"><span data-stu-id="1f970-300">Required.</span></span> <span data-ttu-id="1f970-301">グループのラベルです。</span><span class="sxs-lookup"><span data-stu-id="1f970-301">The label of the group.</span></span> <span data-ttu-id="1f970-302">**resid 属性** は 32 文字以内で [、Resources](resources.md)要素の **ShortStrings** 要素の **String** 要素の **id** 属性の値に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="1f970-302">The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

#### <a name="highlight-requirements"></a><span data-ttu-id="1f970-303">強調表示の要件</span><span class="sxs-lookup"><span data-stu-id="1f970-303">Highlight requirements</span></span>

<span data-ttu-id="1f970-p119">ユーザーは、強調表示されたエンティティに対話型の操作を実行する方法でのみコンテキスト アドインを有効化できます。開発者は、`ItemHasKnownEntity` および `ItemHasRegularExpressionMatch` のルールの種類に対応する `Rule` 要素の `Highlight` 属性を使用して、強調表示にするエンティティを制御します。</span><span class="sxs-lookup"><span data-stu-id="1f970-p119">The only way a user can activate a contextual add-in is to interact with a highlighted entity. Developers can control which entities are highlighted by using the `Highlight` attribute of the `Rule` element for `ItemHasKnownEntity` and `ItemHasRegularExpressionMatch` rule types.</span></span>

<span data-ttu-id="1f970-p120">ただし、注意する必要のある制限があります。これらの制限は、ユーザーにアドインをアクティブ化する方法を提供するために、適用可能なメッセージや予定で強調表示されたエンティティが常に存在するようにするために実施されます。</span><span class="sxs-lookup"><span data-stu-id="1f970-p120">However, there are some limitations to be aware of. These limitations are in place to ensure that there will always be a highlighted entity in applicable messages or appointments to give the user a way to activate the add-in.</span></span>

- <span data-ttu-id="1f970-308">`EmailAddress` および `Url` のエンティティの種類は、強調表示できません。そのため、アドインをアクティブ化するためには使用できません。</span><span class="sxs-lookup"><span data-stu-id="1f970-308">The `EmailAddress` and `Url` entity types cannot be highlighted, and therefore cannot be used to activate an add-in.</span></span>
- <span data-ttu-id="1f970-309">単一のルールを使用する場合、`Highlight` は `all` に設定されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="1f970-309">If using a single rule, `Highlight` MUST be set to `all`.</span></span>
- <span data-ttu-id="1f970-310">複数のルールを組み合わせるために `Mode="AND"` で `RuleCollection` のルールの種類を使用する場合は、少なくとも 1 つのルールの `Highlight` が `all` に設定されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="1f970-310">If using a `RuleCollection` rule type with `Mode="AND"` to combine multiple rules, at least one of the rules MUST have `Highlight` set to `all`.</span></span>
- <span data-ttu-id="1f970-311">複数のルールを組み合わせるために `Mode="OR"` で `RuleCollection` のルールの種類を使用する場合は、すべてのルールの `Highlight` が `all` に設定されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="1f970-311">If using a `RuleCollection` rule type with `Mode="OR"` to combine multiple rules, all of the rules MUST have `Highlight` set to `all`.</span></span>

#### <a name="detectedentity-event-example"></a><span data-ttu-id="1f970-312">DetectedEntity イベントの例</span><span class="sxs-lookup"><span data-stu-id="1f970-312">DetectedEntity event example</span></span>

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
