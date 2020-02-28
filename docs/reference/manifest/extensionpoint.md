---
title: マニフェスト ファイルの ExtensionPoint 要素
description: ''
ms.date: 09/05/2019
localization_priority: Normal
ms.openlocfilehash: 44075bd12c15b4ac9117a51d71fdcc7d6436a7ce
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/27/2020
ms.locfileid: "42324877"
---
# <a name="extensionpoint-element"></a><span data-ttu-id="cf0d1-102">ExtensionPoint 要素</span><span class="sxs-lookup"><span data-stu-id="cf0d1-102">ExtensionPoint element</span></span>

 <span data-ttu-id="cf0d1-103">Office UI でアドインが機能を公開する場所を定義します。</span><span class="sxs-lookup"><span data-stu-id="cf0d1-103">Defines where an add-in exposes functionality in the Office UI.</span></span> <span data-ttu-id="cf0d1-104">**ExtensionPoint** 要素は、[AllFormFactors](allformfactors.md)、[DesktopFormFactor](desktopformfactor.md)、[MobileFormFactor](mobileformfactor.md) の子要素です。</span><span class="sxs-lookup"><span data-stu-id="cf0d1-104">The **ExtensionPoint** element is a child element of [AllFormFactors](allformfactors.md), [DesktopFormFactor](desktopformfactor.md) or [MobileFormFactor](mobileformfactor.md).</span></span> 

## <a name="attributes"></a><span data-ttu-id="cf0d1-105">属性</span><span class="sxs-lookup"><span data-stu-id="cf0d1-105">Attributes</span></span>

|  <span data-ttu-id="cf0d1-106">属性</span><span class="sxs-lookup"><span data-stu-id="cf0d1-106">Attribute</span></span>  |  <span data-ttu-id="cf0d1-107">必須</span><span class="sxs-lookup"><span data-stu-id="cf0d1-107">Required</span></span>  |  <span data-ttu-id="cf0d1-108">説明</span><span class="sxs-lookup"><span data-stu-id="cf0d1-108">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="cf0d1-109">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="cf0d1-109">**xsi:type**</span></span>  |  <span data-ttu-id="cf0d1-110">はい</span><span class="sxs-lookup"><span data-stu-id="cf0d1-110">Yes</span></span>  | <span data-ttu-id="cf0d1-111">定義される拡張点の種類。</span><span class="sxs-lookup"><span data-stu-id="cf0d1-111">The type of extension point being defined.</span></span>|

## <a name="extension-points-for-excel-only"></a><span data-ttu-id="cf0d1-112">Excel のみの拡張点</span><span class="sxs-lookup"><span data-stu-id="cf0d1-112">Extension points for Excel only</span></span>

- <span data-ttu-id="cf0d1-113">**CustomFunctions** - Excel 向けの JavaScript で記述されたカスタム関数。</span><span class="sxs-lookup"><span data-stu-id="cf0d1-113">**CustomFunctions** - A custom function written in JavaScript for Excel.</span></span>

<span data-ttu-id="cf0d1-114">[この XML コード サンプル](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/manifest.xml)は、**CustomFunctions** 属性の値を持つ **ExtensionPoint** 要素を使用する方法と、使用する子要素を示しています。</span><span class="sxs-lookup"><span data-stu-id="cf0d1-114">[This XML code sample](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/manifest.xml) shows how to use the **ExtensionPoint** element with the **CustomFunctions** attribute value, and the child elements to be used.</span></span>

## <a name="extension-points-for-word-excel-powerpoint-and-onenote-add-in-commands"></a><span data-ttu-id="cf0d1-115">Word、Excel、PowerPoint、OneNote アドイン コマンドの拡張点</span><span class="sxs-lookup"><span data-stu-id="cf0d1-115">Extension points for Word, Excel, PowerPoint, and OneNote add-in commands</span></span>

- <span data-ttu-id="cf0d1-116">**PrimaryCommandSurface** - Office のリボン。</span><span class="sxs-lookup"><span data-stu-id="cf0d1-116">**PrimaryCommandSurface** - The ribbon in Office.</span></span>
- <span data-ttu-id="cf0d1-117">**ContextMenu**Office UI で右クリックしたときに表示されるショートカット メニュー。</span><span class="sxs-lookup"><span data-stu-id="cf0d1-117">**ContextMenu** - The shortcut menu that appears when you right-click in the Office UI.</span></span>

<span data-ttu-id="cf0d1-118">次の例は、**PrimaryCommandSurface** と **ContextMenu** の属性値を持つ **ExtensionPoint** 要素を使用する方法と、各要素と併用する必要がある子要素を示しています。</span><span class="sxs-lookup"><span data-stu-id="cf0d1-118">The following examples show how to use the **ExtensionPoint** element with **PrimaryCommandSurface** and **ContextMenu** attribute values, and the child elements that should be used with each.</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="cf0d1-p102">ID 属性を含む要素では、一意の ID を指定してください。会社の名前と ID を使用することをお勧めします。たとえば、次の形式にします。<CustomTab id="mycompanyname.mygroupname"></span><span class="sxs-lookup"><span data-stu-id="cf0d1-p102">For elements that contain an ID attribute, make sure you provide a unique ID. We recommend that you use your company's name along with your ID. For example, use the following format. <CustomTab id="mycompanyname.mygroupname"></span></span>

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

#### <a name="child-elements"></a><span data-ttu-id="cf0d1-122">子要素</span><span class="sxs-lookup"><span data-stu-id="cf0d1-122">Child elements</span></span>
 
|<span data-ttu-id="cf0d1-123">**Element**</span><span class="sxs-lookup"><span data-stu-id="cf0d1-123">**Element**</span></span>|<span data-ttu-id="cf0d1-124">**説明**</span><span class="sxs-lookup"><span data-stu-id="cf0d1-124">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="cf0d1-125">**CustomTab**</span><span class="sxs-lookup"><span data-stu-id="cf0d1-125">**CustomTab**</span></span>|<span data-ttu-id="cf0d1-p103">カスタム タブをリボンに追加する必要がある場合は必須 (**PrimaryCommandSurface** を使用)。**CustomTab** 要素を使用する場合、**OfficeTab** 要素は使用できません。**id** 属性が必要です。 </span><span class="sxs-lookup"><span data-stu-id="cf0d1-p103">Required if you want to add a custom tab to the ribbon (using **PrimaryCommandSurface**). If you use the **CustomTab** element, you can't use the **OfficeTab** element. The **id** attribute is required.</span></span>|
|<span data-ttu-id="cf0d1-129">**OfficeTab**</span><span class="sxs-lookup"><span data-stu-id="cf0d1-129">**OfficeTab**</span></span>|<span data-ttu-id="cf0d1-p104">既定の Office リボンタブ ( **Primarycommandsurface**を使用) を拡張する場合に必要です。**Officetab**要素を使用する場合、 **customtab**要素は使用できません。詳細については、「 [Officetab](officetab.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="cf0d1-p104">Required if you want to extend a default Office ribbon tab (using **PrimaryCommandSurface**). If you use the **OfficeTab** element, you can't use the **CustomTab** element. For details, see [OfficeTab](officetab.md).</span></span>|
|<span data-ttu-id="cf0d1-133">**OfficeMenu**</span><span class="sxs-lookup"><span data-stu-id="cf0d1-133">**OfficeMenu**</span></span>|<span data-ttu-id="cf0d1-p105">既定のコンテキスト メニューにアドイン コマンドを追加する場合は必須 (**ContextMenu** を使用)。**id** 属性は以下に設定する必要があります。 </span><span class="sxs-lookup"><span data-stu-id="cf0d1-p105">Required if you're adding add-in commands to a default context menu (using **ContextMenu**). The **id** attribute must be set to: </span></span><br/> <span data-ttu-id="cf0d1-p106">Excel または Word の場合は - **ContextMenuText**。テキストが選択され、ユーザーが選択されたテキストを右クリックしたときに、コンテキスト メニューに項目が表示されます。 </span><span class="sxs-lookup"><span data-stu-id="cf0d1-p106">- **ContextMenuText** for Excel or Word. Displays the item on the context menu when text is selected and then the user right-clicks on the selected text. </span></span><br/> <span data-ttu-id="cf0d1-p107">Excel の場合は - **ContextMenuCell**。ユーザーがスプレッドシートのセルを右クリックすると、コンテキスト メニューに項目が表示されます。</span><span class="sxs-lookup"><span data-stu-id="cf0d1-p107">- **ContextMenuCell** for Excel. Displays the  item on the context menu when the user right-clicks on a cell on the spreadsheet.</span></span>|
|<span data-ttu-id="cf0d1-140">**Group**</span><span class="sxs-lookup"><span data-stu-id="cf0d1-140">**Group**</span></span>|<span data-ttu-id="cf0d1-p108">タブのユーザー インターフェイスの拡張点のグループ。1 つのグループに、最大 6 個のコントロールを指定できます。**id** 属性が必要です。最大 125 文字の文字列です。 </span><span class="sxs-lookup"><span data-stu-id="cf0d1-p108">A group of user interface extension points on a tab. A group can have up to six controls. The **id** attribute is required. It's a string with a maximum of 125 characters.</span></span>|
|<span data-ttu-id="cf0d1-144">**Label**</span><span class="sxs-lookup"><span data-stu-id="cf0d1-144">**Label**</span></span>|<span data-ttu-id="cf0d1-p109">必須。グループのラベル。**resid** 属性は、**String** 要素の **id** 属性の値に設定する必要があります。**String** 要素は、**ShortStrings** 要素 (**Resources** 要素の子要素) の子要素です。 </span><span class="sxs-lookup"><span data-stu-id="cf0d1-p109">Required. The label of the group. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **ShortStrings** element, which is a child element of the **Resources** element.</span></span>|
|<span data-ttu-id="cf0d1-149">**Icon**</span><span class="sxs-lookup"><span data-stu-id="cf0d1-149">**Icon**</span></span>|<span data-ttu-id="cf0d1-p110">必須。小さいフォーム ファクターのデバイス、または多くのボタンが表示されるときに使用されるグループのアイコンを指定します。**resid** 属性は、**Image** 要素の **id** 属性の値に設定する必要があります。**Image** 要素は、**Images** 要素 (**Resources** 要素の子要素) の子要素です。**size** 属性は、イメージのサイズをピクセル単位で指定します。次の 3 つのイメージのサイズが必要です。16、32、および 80。次の 5 つのオプションのサイズもサポートされています。20、24、40、48、および 64。 </span><span class="sxs-lookup"><span data-stu-id="cf0d1-p110">Required. Specifies the group's icon to be used on small form factor devices, or when too many buttons are displayed. The **resid** attribute must be set to the value of the **id** attribute of an **Image** element. The **Image** element is a child element of the **Images** element, which is a child element of the **Resources** element. The **size** attribute gives the size, in pixels, of the image. Three image sizes are required: 16, 32, and 80. Five optional sizes are also supported: 20, 24, 40, 48, and 64.</span></span>|
|<span data-ttu-id="cf0d1-157">**Tooltip**</span><span class="sxs-lookup"><span data-stu-id="cf0d1-157">**Tooltip**</span></span>|<span data-ttu-id="cf0d1-p111">省略可能。グループのヒント。**resid** 属性は、**String** 要素の **id** 属性の値に設定する必要があります。**String** 要素は、**LongStrings** 要素 (**Resources** 要素の子要素) の子要素です。 </span><span class="sxs-lookup"><span data-stu-id="cf0d1-p111">Optional. The tooltip of the group. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **LongStrings** element, which is a child element of the **Resources** element.</span></span>|
|<span data-ttu-id="cf0d1-162">**Control**</span><span class="sxs-lookup"><span data-stu-id="cf0d1-162">**Control**</span></span>|<span data-ttu-id="cf0d1-163">各グループには、少なくとも 1 つのコントロールが必要です。</span><span class="sxs-lookup"><span data-stu-id="cf0d1-163">Each group requires at least one control.</span></span> <span data-ttu-id="cf0d1-164">**Control**要素には、**ボタン**または**メニュー**のいずれかを指定できます。</span><span class="sxs-lookup"><span data-stu-id="cf0d1-164">A **Control** element can be either a **Button** or a **Menu**.</span></span> <span data-ttu-id="cf0d1-165">**メニュー**を使用して、ボタンコントロールのドロップダウンリストを指定します。</span><span class="sxs-lookup"><span data-stu-id="cf0d1-165">Use **Menu** to specify a drop-down list of button controls.</span></span> <span data-ttu-id="cf0d1-166">現在は、ボタンとメニューのみがサポートされています。</span><span class="sxs-lookup"><span data-stu-id="cf0d1-166">Currently, only buttons and menus are supported.</span></span> <span data-ttu-id="cf0d1-167">詳細については、「[Button コントロール](control.md#button-control)」および「[Menu コントロール](control.md#menu-dropdown-button-controls)」のセクションを参照してください。</span><span class="sxs-lookup"><span data-stu-id="cf0d1-167">See the [Button controls](control.md#button-control) and [Menu controls](control.md#menu-dropdown-button-controls) sections for more information.</span></span><br/><span data-ttu-id="cf0d1-168">**注:** トラブルシューティングを簡単にするために、 **Control**要素と関連する**Resources**子要素を一度に1つずつ追加することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="cf0d1-168">**Note:**  To make troubleshooting easier, we recommend that a **Control** element and the related **Resources** child elements be added one at a time.</span></span>|
|<span data-ttu-id="cf0d1-169">**スクリプト**</span><span class="sxs-lookup"><span data-stu-id="cf0d1-169">**Script**</span></span>|<span data-ttu-id="cf0d1-170">カスタム関数の定義と登録コードを含む JavaScript ファイルにリンクします。</span><span class="sxs-lookup"><span data-stu-id="cf0d1-170">Links to the JavaScript file with the custom function definition and registration code.</span></span> <span data-ttu-id="cf0d1-171">Developer Preview では、この要素は使用しません。</span><span class="sxs-lookup"><span data-stu-id="cf0d1-171">This element is not used in the Developer Preview.</span></span> <span data-ttu-id="cf0d1-172">代わりに、HTML ページはすべての JavaScript ファイルを読み込みます。</span><span class="sxs-lookup"><span data-stu-id="cf0d1-172">Instead, the HTML page is responsible for loading all JavaScript files.</span></span>|
|<span data-ttu-id="cf0d1-173">**Page**</span><span class="sxs-lookup"><span data-stu-id="cf0d1-173">**Page**</span></span>|<span data-ttu-id="cf0d1-174">カスタム関数についての HTML ページにリンクします。</span><span class="sxs-lookup"><span data-stu-id="cf0d1-174">Links to the HTML page for your custom functions.</span></span>|

## <a name="extension-points-for-outlook"></a><span data-ttu-id="cf0d1-175">Outlook のみの拡張点</span><span class="sxs-lookup"><span data-stu-id="cf0d1-175">Extension points for Outlook</span></span>

- [<span data-ttu-id="cf0d1-176">MessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="cf0d1-176">MessageReadCommandSurface</span></span>](#messagereadcommandsurface) 
- [<span data-ttu-id="cf0d1-177">MessageComposeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="cf0d1-177">MessageComposeCommandSurface</span></span>](#messagecomposecommandsurface) 
- [<span data-ttu-id="cf0d1-178">AppointmentOrganizerCommandSurface</span><span class="sxs-lookup"><span data-stu-id="cf0d1-178">AppointmentOrganizerCommandSurface</span></span>](#appointmentorganizercommandsurface) 
- [<span data-ttu-id="cf0d1-179">AppointmentAttendeeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="cf0d1-179">AppointmentAttendeeCommandSurface</span></span>](#appointmentattendeecommandsurface)
- <span data-ttu-id="cf0d1-180">[Module](#module) ([DesktopFormFactor](desktopformfactor.md) でのみ使用できます。)</span><span class="sxs-lookup"><span data-stu-id="cf0d1-180">[Module](#module) (Can only be used in the [DesktopFormFactor](desktopformfactor.md).)</span></span>
- [<span data-ttu-id="cf0d1-181">MobileMessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="cf0d1-181">MobileMessageReadCommandSurface</span></span>](#mobilemessagereadcommandsurface)
- [<span data-ttu-id="cf0d1-182">Events</span><span class="sxs-lookup"><span data-stu-id="cf0d1-182">Events</span></span>](#events)
- [<span data-ttu-id="cf0d1-183">DetectedEntity</span><span class="sxs-lookup"><span data-stu-id="cf0d1-183">DetectedEntity</span></span>](#detectedentity)

### <a name="messagereadcommandsurface"></a><span data-ttu-id="cf0d1-184">MessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="cf0d1-184">MessageReadCommandSurface</span></span>
<span data-ttu-id="cf0d1-p114">この拡張点により、メールの閲覧ビューのコマンド サーフェスにボタンが配置されます。Outlook デスクトップでは、これはリボンに表示されます。</span><span class="sxs-lookup"><span data-stu-id="cf0d1-p114">This extension point puts buttons in the command surface for the mail read view. In Outlook desktop, this appears in the ribbon.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="cf0d1-187">子要素</span><span class="sxs-lookup"><span data-stu-id="cf0d1-187">Child elements</span></span>

|  <span data-ttu-id="cf0d1-188">要素</span><span class="sxs-lookup"><span data-stu-id="cf0d1-188">Element</span></span> |  <span data-ttu-id="cf0d1-189">説明</span><span class="sxs-lookup"><span data-stu-id="cf0d1-189">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="cf0d1-190">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="cf0d1-190">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="cf0d1-191">コマンドを既定のリボン タブに追加します。</span><span class="sxs-lookup"><span data-stu-id="cf0d1-191">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="cf0d1-192">CustomTab</span><span class="sxs-lookup"><span data-stu-id="cf0d1-192">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="cf0d1-193">コマンドをカスタム リボン タブに追加します。</span><span class="sxs-lookup"><span data-stu-id="cf0d1-193">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="cf0d1-194">OfficeTab の例</span><span class="sxs-lookup"><span data-stu-id="cf0d1-194">OfficeTab example</span></span>
```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="cf0d1-195">CustomTab の例</span><span class="sxs-lookup"><span data-stu-id="cf0d1-195">CustomTab example</span></span>
```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="messagecomposecommandsurface"></a><span data-ttu-id="cf0d1-196">MessageComposeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="cf0d1-196">MessageComposeCommandSurface</span></span>
<span data-ttu-id="cf0d1-197">この拡張点は、メールの新規作成フォームを使用してアドイン用のリボンにボタンを配置します。</span><span class="sxs-lookup"><span data-stu-id="cf0d1-197">This extension point puts buttons on the ribbon for add-ins using mail compose form.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="cf0d1-198">子要素</span><span class="sxs-lookup"><span data-stu-id="cf0d1-198">Child elements</span></span>

|  <span data-ttu-id="cf0d1-199">要素</span><span class="sxs-lookup"><span data-stu-id="cf0d1-199">Element</span></span> |  <span data-ttu-id="cf0d1-200">説明</span><span class="sxs-lookup"><span data-stu-id="cf0d1-200">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="cf0d1-201">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="cf0d1-201">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="cf0d1-202">コマンドを既定のリボン タブに追加します。</span><span class="sxs-lookup"><span data-stu-id="cf0d1-202">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="cf0d1-203">CustomTab</span><span class="sxs-lookup"><span data-stu-id="cf0d1-203">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="cf0d1-204">コマンドをカスタム リボン タブに追加します。</span><span class="sxs-lookup"><span data-stu-id="cf0d1-204">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="cf0d1-205">OfficeTab の例</span><span class="sxs-lookup"><span data-stu-id="cf0d1-205">OfficeTab example</span></span>
```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="cf0d1-206">CustomTab の例</span><span class="sxs-lookup"><span data-stu-id="cf0d1-206">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="appointmentorganizercommandsurface"></a><span data-ttu-id="cf0d1-207">AppointmentOrganizerCommandSurface</span><span class="sxs-lookup"><span data-stu-id="cf0d1-207">AppointmentOrganizerCommandSurface</span></span>

<span data-ttu-id="cf0d1-208">この拡張点は、会議の開催者に表示されるフォームのリボンにボタンを配置します。</span><span class="sxs-lookup"><span data-stu-id="cf0d1-208">This extension point puts buttons on the ribbon for the form that's displayed to the organizer of the meeting.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="cf0d1-209">子要素</span><span class="sxs-lookup"><span data-stu-id="cf0d1-209">Child elements</span></span>

|  <span data-ttu-id="cf0d1-210">要素</span><span class="sxs-lookup"><span data-stu-id="cf0d1-210">Element</span></span> |  <span data-ttu-id="cf0d1-211">説明</span><span class="sxs-lookup"><span data-stu-id="cf0d1-211">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="cf0d1-212">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="cf0d1-212">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="cf0d1-213">コマンドを既定のリボン タブに追加します。</span><span class="sxs-lookup"><span data-stu-id="cf0d1-213">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="cf0d1-214">CustomTab</span><span class="sxs-lookup"><span data-stu-id="cf0d1-214">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="cf0d1-215">コマンドをカスタム リボン タブに追加します。</span><span class="sxs-lookup"><span data-stu-id="cf0d1-215">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="cf0d1-216">OfficeTab の例</span><span class="sxs-lookup"><span data-stu-id="cf0d1-216">OfficeTab example</span></span>
```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="cf0d1-217">CustomTab の例</span><span class="sxs-lookup"><span data-stu-id="cf0d1-217">CustomTab example</span></span>
```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="appointmentattendeecommandsurface"></a><span data-ttu-id="cf0d1-218">AppointmentAttendeeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="cf0d1-218">AppointmentAttendeeCommandSurface</span></span>

<span data-ttu-id="cf0d1-219">この拡張点は、会議の出席者に表示されるフォームのリボンにボタンを配置します。</span><span class="sxs-lookup"><span data-stu-id="cf0d1-219">This extension point puts buttons on the ribbon for the form that's displayed to the attendee of the meeting.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="cf0d1-220">子要素</span><span class="sxs-lookup"><span data-stu-id="cf0d1-220">Child elements</span></span>

|  <span data-ttu-id="cf0d1-221">要素</span><span class="sxs-lookup"><span data-stu-id="cf0d1-221">Element</span></span> |  <span data-ttu-id="cf0d1-222">説明</span><span class="sxs-lookup"><span data-stu-id="cf0d1-222">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="cf0d1-223">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="cf0d1-223">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="cf0d1-224">コマンドを既定のリボン タブに追加します。</span><span class="sxs-lookup"><span data-stu-id="cf0d1-224">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="cf0d1-225">CustomTab</span><span class="sxs-lookup"><span data-stu-id="cf0d1-225">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="cf0d1-226">コマンドをカスタム リボン タブに追加します。</span><span class="sxs-lookup"><span data-stu-id="cf0d1-226">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="cf0d1-227">OfficeTab の例</span><span class="sxs-lookup"><span data-stu-id="cf0d1-227">OfficeTab example</span></span>
```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="cf0d1-228">CustomTab の例</span><span class="sxs-lookup"><span data-stu-id="cf0d1-228">CustomTab example</span></span>
```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="module"></a><span data-ttu-id="cf0d1-229">Module</span><span class="sxs-lookup"><span data-stu-id="cf0d1-229">Module</span></span>

<span data-ttu-id="cf0d1-230">この拡張点は、モジュール拡張機能用のリボンにボタンを配置します。</span><span class="sxs-lookup"><span data-stu-id="cf0d1-230">This extension point puts buttons on the ribbon for the module extension.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="cf0d1-231">子要素</span><span class="sxs-lookup"><span data-stu-id="cf0d1-231">Child elements</span></span>

|  <span data-ttu-id="cf0d1-232">要素</span><span class="sxs-lookup"><span data-stu-id="cf0d1-232">Element</span></span> |  <span data-ttu-id="cf0d1-233">説明</span><span class="sxs-lookup"><span data-stu-id="cf0d1-233">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="cf0d1-234">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="cf0d1-234">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="cf0d1-235">コマンドを既定のリボン タブに追加します。</span><span class="sxs-lookup"><span data-stu-id="cf0d1-235">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="cf0d1-236">CustomTab</span><span class="sxs-lookup"><span data-stu-id="cf0d1-236">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="cf0d1-237">コマンドをカスタム リボン タブに追加します。</span><span class="sxs-lookup"><span data-stu-id="cf0d1-237">Adds the command(s) to the custom ribbon tab.</span></span>  |

### <a name="mobilemessagereadcommandsurface"></a><span data-ttu-id="cf0d1-238">MobileMessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="cf0d1-238">MobileMessageReadCommandSurface</span></span>
<span data-ttu-id="cf0d1-239">この拡張点により、モバイル フォーム ファクターのメールの閲覧ビューのコマンド領域にボタンが配置されます。</span><span class="sxs-lookup"><span data-stu-id="cf0d1-239">This extension point puts buttons in the command surface for the mail read view in the mobile form factor.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="cf0d1-240">子要素</span><span class="sxs-lookup"><span data-stu-id="cf0d1-240">Child elements</span></span>

|  <span data-ttu-id="cf0d1-241">要素</span><span class="sxs-lookup"><span data-stu-id="cf0d1-241">Element</span></span> |  <span data-ttu-id="cf0d1-242">説明</span><span class="sxs-lookup"><span data-stu-id="cf0d1-242">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="cf0d1-243">Group</span><span class="sxs-lookup"><span data-stu-id="cf0d1-243">Group</span></span>](group.md) |  <span data-ttu-id="cf0d1-244">コマンド領域にボタンのグループを追加します。</span><span class="sxs-lookup"><span data-stu-id="cf0d1-244">Adds a group of buttons to the command surface.</span></span>  |

<span data-ttu-id="cf0d1-245">この種類の **ExtensionPoint** 要素には子要素を 1 つだけ含めることができます (**Group** 要素)。</span><span class="sxs-lookup"><span data-stu-id="cf0d1-245">**ExtensionPoint** elements of this type can only have one child element: a **Group** element.</span></span>

<span data-ttu-id="cf0d1-246">この拡張点に含まれる **Control** 要素の **xsi:type** 属性を `MobileButton` に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="cf0d1-246">**Control** elements contained in this extension point must have the **xsi:type** attribute set to `MobileButton`.</span></span>

#### <a name="example"></a><span data-ttu-id="cf0d1-247">例</span><span class="sxs-lookup"><span data-stu-id="cf0d1-247">Example</span></span>
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

### <a name="events"></a><span data-ttu-id="cf0d1-248">Events</span><span class="sxs-lookup"><span data-stu-id="cf0d1-248">Events</span></span>

<span data-ttu-id="cf0d1-249">この拡張点は、指定したイベントのイベント ハンドラーを追加します。</span><span class="sxs-lookup"><span data-stu-id="cf0d1-249">This extension point adds an event handler for a specified event.</span></span>

> [!NOTE]
> <span data-ttu-id="cf0d1-250">この要素の種類は、従来の Outlook on the web でサポートされており、Windows、Mac、および最新の Outlook on the web では[プレビュー](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md)されています。</span><span class="sxs-lookup"><span data-stu-id="cf0d1-250">This element type is supported by classic Outlook on the web, and in [preview](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) on Windows, Mac, and modern Outlook on the web.</span></span> <span data-ttu-id="cf0d1-251">Office 365 サブスクリプションも必要です。</span><span class="sxs-lookup"><span data-stu-id="cf0d1-251">An Office 365 subscription is also required.</span></span>

| <span data-ttu-id="cf0d1-252">要素</span><span class="sxs-lookup"><span data-stu-id="cf0d1-252">Element</span></span> | <span data-ttu-id="cf0d1-253">説明</span><span class="sxs-lookup"><span data-stu-id="cf0d1-253">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="cf0d1-254">Event</span><span class="sxs-lookup"><span data-stu-id="cf0d1-254">Event</span></span>](event.md) |  <span data-ttu-id="cf0d1-255">イベントとイベント ハンドラーの関数を指定します。</span><span class="sxs-lookup"><span data-stu-id="cf0d1-255">Specifies the event and event handler function.</span></span>  |

#### <a name="itemsend-event-example"></a><span data-ttu-id="cf0d1-256">ItemSend イベントの例</span><span class="sxs-lookup"><span data-stu-id="cf0d1-256">ItemSend event example</span></span>

```xml
<ExtensionPoint xsi:type="Events">
  <Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="itemSendHandler" />
</ExtensionPoint>
```

### <a name="detectedentity"></a><span data-ttu-id="cf0d1-257">DetectedEntity</span><span class="sxs-lookup"><span data-stu-id="cf0d1-257">DetectedEntity</span></span>

<span data-ttu-id="cf0d1-258">この拡張点は、指定したエンティティの種類に対するコンテキスト アドインのアクティブ化を追加します。</span><span class="sxs-lookup"><span data-stu-id="cf0d1-258">This extension point adds a contextual add-in activation on a specified entity type.</span></span>

<span data-ttu-id="cf0d1-259">これを収容している [VersionOverrides](versionoverrides.md) 要素は、`xsi:type` 属性の値が `VersionOverridesV1_1` になっている必要があります。</span><span class="sxs-lookup"><span data-stu-id="cf0d1-259">The containing [VersionOverrides](versionoverrides.md) element must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span>

> [!NOTE]
> <span data-ttu-id="cf0d1-260">この要素の種類は、[要件セット 1.6 以降をサポートする Outlook クライアント ](../requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)が利用できます。</span><span class="sxs-lookup"><span data-stu-id="cf0d1-260">This element type is available to [Outlook clients that support requirement sets 1.6 and later](../requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients).</span></span>

|  <span data-ttu-id="cf0d1-261">要素</span><span class="sxs-lookup"><span data-stu-id="cf0d1-261">Element</span></span> |  <span data-ttu-id="cf0d1-262">説明</span><span class="sxs-lookup"><span data-stu-id="cf0d1-262">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="cf0d1-263">Label</span><span class="sxs-lookup"><span data-stu-id="cf0d1-263">Label</span></span>](#label) |  <span data-ttu-id="cf0d1-264">アドインのコンテキスト ウィンドウのラベルを指定します。</span><span class="sxs-lookup"><span data-stu-id="cf0d1-264">Specifies the label for the add-in in the contextual window.</span></span>  |
|  [<span data-ttu-id="cf0d1-265">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="cf0d1-265">SourceLocation</span></span>](sourcelocation.md) |  <span data-ttu-id="cf0d1-266">コンテキスト ウィンドウの URL を指定します。</span><span class="sxs-lookup"><span data-stu-id="cf0d1-266">Specifies the URL for the contextual window.</span></span>  |
|  [<span data-ttu-id="cf0d1-267">Rule</span><span class="sxs-lookup"><span data-stu-id="cf0d1-267">Rule</span></span>](rule.md) |  <span data-ttu-id="cf0d1-268">アドインをアクティブ化するタイミングを決定する 1 つ以上のルールを指定します。</span><span class="sxs-lookup"><span data-stu-id="cf0d1-268">Specifies the rule or rules that determine when an add-in activates.</span></span>  |

#### <a name="label"></a><span data-ttu-id="cf0d1-269">Label</span><span class="sxs-lookup"><span data-stu-id="cf0d1-269">Label</span></span>

<span data-ttu-id="cf0d1-p116">必須。グループのラベルを示します。**Resid**属性は、 [Resources](resources.md)要素の Short **strings**要素の**String**要素の**id**属性の値に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="cf0d1-p116">Required. The label of the group. The **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

#### <a name="highlight-requirements"></a><span data-ttu-id="cf0d1-273">強調表示の要件</span><span class="sxs-lookup"><span data-stu-id="cf0d1-273">Highlight requirements</span></span>

<span data-ttu-id="cf0d1-p117">ユーザーは、強調表示されたエンティティに対話型の操作を実行する方法でのみコンテキスト アドインを有効化できます。開発者は、`ItemHasKnownEntity` および `ItemHasRegularExpressionMatch` のルールの種類に対応する `Rule` 要素の `Highlight` 属性を使用して、強調表示にするエンティティを制御します。</span><span class="sxs-lookup"><span data-stu-id="cf0d1-p117">The only way a user can activate a contextual add-in is to interact with a highlighted entity. Developers can control which entities are highlighted by using the `Highlight` attribute of the `Rule` element for `ItemHasKnownEntity` and `ItemHasRegularExpressionMatch` rule types.</span></span>

<span data-ttu-id="cf0d1-p118">ただし、注意する必要のある制限があります。これらの制限は、ユーザーにアドインをアクティブ化する方法を提供するために、適用可能なメッセージや予定で強調表示されたエンティティが常に存在するようにするために実施されます。</span><span class="sxs-lookup"><span data-stu-id="cf0d1-p118">However, there are some limitations to be aware of. These limitations are in place to ensure that there will always be a highlighted entity in applicable messages or appointments to give the user a way to activate the add-in.</span></span>

- <span data-ttu-id="cf0d1-278">`EmailAddress` および `Url` のエンティティの種類は、強調表示できません。そのため、アドインをアクティブ化するためには使用できません。</span><span class="sxs-lookup"><span data-stu-id="cf0d1-278">The `EmailAddress` and `Url` entity types cannot be highlighted, and therefore cannot be used to activate an add-in.</span></span>
- <span data-ttu-id="cf0d1-279">単一のルールを使用する場合、`Highlight` は `all` に設定されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="cf0d1-279">If using a single rule, `Highlight` MUST be set to `all`.</span></span>
- <span data-ttu-id="cf0d1-280">複数のルールを組み合わせるために `Mode="AND"` で `RuleCollection` のルールの種類を使用する場合は、少なくとも 1 つのルールの `Highlight` が `all` に設定されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="cf0d1-280">If using a `RuleCollection` rule type with `Mode="AND"` to combine multiple rules, at least one of the rules MUST have `Highlight` set to `all`.</span></span>
- <span data-ttu-id="cf0d1-281">複数のルールを組み合わせるために `Mode="OR"` で `RuleCollection` のルールの種類を使用する場合は、すべてのルールの `Highlight` が `all` に設定されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="cf0d1-281">If using a `RuleCollection` rule type with `Mode="OR"` to combine multiple rules, all of the rules MUST have `Highlight` set to `all`.</span></span>

#### <a name="detectedentity-event-example"></a><span data-ttu-id="cf0d1-282">DetectedEntity イベントの例</span><span class="sxs-lookup"><span data-stu-id="cf0d1-282">DetectedEntity event example</span></span>

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
