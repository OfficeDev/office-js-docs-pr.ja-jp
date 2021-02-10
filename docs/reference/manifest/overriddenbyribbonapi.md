---
title: マニフェスト ファイルの OverriddenByRibbonApi 要素
description: カスタム の操作別タブの一部である場合に、カスタム タブ、グループ、コントロール、またはメニュー項目を表示してはならないことを指定する方法について説明します。
ms.date: 01/29/2021
localization_priority: Normal
ms.openlocfilehash: 62aa484057221f9cd7f41af9c8b9210cdb5b3376
ms.sourcegitcommit: 4805454f7fc6c64368a35d014e24075faf3e7557
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/10/2021
ms.locfileid: "50174001"
---
# <a name="overriddenbyribbonapi-element"></a><span data-ttu-id="cd223-103">OverriddenByRibbonApi 要素</span><span class="sxs-lookup"><span data-stu-id="cd223-103">OverriddenByRibbonApi element</span></span>

<span data-ttu-id="cd223-104">カスタム コンテキスト タブをリボンにインストールする API [](control.md#menu-dropdown-button-controls) ([Office.ribbon.requestCreateControls)](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls-tabDefinition-)をサポートするアプリケーションとプラットフォームの組み合わせで[、CustomTab、Group、Button](customtab.md)[コントロール、](control.md#button-control)メニュー コントロール、またはメニュー項目を非表示にするかどうかを指定します。 [](group.md)</span><span class="sxs-lookup"><span data-stu-id="cd223-104">Specifies whether a [CustomTab](customtab.md), [Group](group.md), [Button](control.md#button-control) control, [Menu](control.md#menu-dropdown-button-controls) control, or menu item will be hidden on application and platform combinations that support the API ([Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls-tabDefinition-)) that installs custom contextual tabs on the ribbon.</span></span>

<span data-ttu-id="cd223-105">省略すると、既定値は `false` .</span><span class="sxs-lookup"><span data-stu-id="cd223-105">If it is omitted, the default is `false`.</span></span> <span data-ttu-id="cd223-106">使用する場合は、親要素の *最初の* 子要素である必要があります。</span><span class="sxs-lookup"><span data-stu-id="cd223-106">If it is used, it must be the *first* child element of its parent element.</span></span>

> [!NOTE]
> <span data-ttu-id="cd223-107">この要素の詳細については、「カスタム コンテキスト タブがサポートされていない場合に代替 UI エクスペリエンスを実装する」 [を参照してください](../../design/contextual-tabs.md#implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported)。</span><span class="sxs-lookup"><span data-stu-id="cd223-107">For a full understanding of this element, please read [Implement an alternate UI experience when custom contextual tabs are not supported](../../design/contextual-tabs.md#implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported).</span></span>

<span data-ttu-id="cd223-108">この要素の目的は、カスタム コンテキスト タブをサポートしないアプリケーションまたはプラットフォームでアドインが実行されている場合に、カスタム コンテキスト タブを実装するフォールバック エクスペリエンスをアドインに作成します。</span><span class="sxs-lookup"><span data-stu-id="cd223-108">The purpose of this element is to create a fallback experience in an add-in that implements custom contextual tabs when the add-in is running on an application or platform that doesn't support custom contextual tabs.</span></span> <span data-ttu-id="cd223-109">重要な戦略は、カスタム コンテキスト タブの一部またはすべてのグループとコントロールを 1 つ以上のカスタム コア タブ (非コンテキスト カスタム タブ) に複製 *する方法です* 。</span><span class="sxs-lookup"><span data-stu-id="cd223-109">The essential strategy is that you duplicate some or all of the groups and controls from your custom contextual tab onto one or more custom core tabs (that is, *noncontextual* custom tabs).</span></span> <span data-ttu-id="cd223-110">次に、カスタム コンテキスト タブがサポートされていないときにこれらのグループとコントロールが表示されますが、カスタム コンテキスト タブがサポートされている場合には表示されない場合は `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` **、CustomTab** 要素 **、Group** 要素 **、Control** 要素、またはメニュー **項目** 要素の最初の子要素として追加します。</span><span class="sxs-lookup"><span data-stu-id="cd223-110">Then, to ensure that these groups and controls appear when custom contextual tabs are *not* supported, but do not appear when custom contextual tabs *are* supported, you add `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` as the first child element of the **CustomTab**, **Group**, **Control**, or menu **Item** elements.</span></span> <span data-ttu-id="cd223-111">その効果は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="cd223-111">The effect of doing so is the following:</span></span>

- <span data-ttu-id="cd223-112">カスタム コンテキスト タブをサポートするアプリケーションとプラットフォームでアドインを実行する場合、複製されたタブ、グループ、およびコントロールはリボンに表示されません。</span><span class="sxs-lookup"><span data-stu-id="cd223-112">If the add-in runs on an application and platform that support custom contextual tabs, then the duplicated tabs, groups, and controls won't appear on the ribbon.</span></span> <span data-ttu-id="cd223-113">代わりに、アドインがメソッドを呼び出す際に、カスタム コンテキスト タブがインストール `requestCreateControls` されます。</span><span class="sxs-lookup"><span data-stu-id="cd223-113">Instead, the custom contextual tab will be installed when the add-in calls the `requestCreateControls` method.</span></span>
- <span data-ttu-id="cd223-114">カスタム コンテキスト タブをサポートしないアプリケーションまたはプラットフォームでアドインを実行すると、複製されたタブ、グループ、およびコントロールがリボンに表示されます。</span><span class="sxs-lookup"><span data-stu-id="cd223-114">If the add-in runs on an application or platform that *doesn't* support custom contextual tabs, then the duplicated tabs, groups, and controls will appear on the ribbon.</span></span>

## <a name="examples"></a><span data-ttu-id="cd223-115">例</span><span class="sxs-lookup"><span data-stu-id="cd223-115">Examples</span></span>

### <a name="overriding-an-entire-tab"></a><span data-ttu-id="cd223-116">タブ全体を上書きする</span><span class="sxs-lookup"><span data-stu-id="cd223-116">Overriding an entire tab</span></span>

```xml
<ExtensionPoint xsi:type="PrimaryCommandSurface">
  <CustomTab id="TabCustom1">
    <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
    <Group id="ContosoCustomTab.grp1">
      <Control  xsi:type="Button" id="MyButton">
        <!-- Child elements omitted. -->
      </Control>
    </Group>
    <Label resid="customTabLabel1"/>
  </CustomTab>
</ExtensionPoint>
```

### <a name="overriding-a-group"></a><span data-ttu-id="cd223-117">グループのオーバーライド</span><span class="sxs-lookup"><span data-stu-id="cd223-117">Overriding a group</span></span>

```xml
<ExtensionPoint xsi:type="PrimaryCommandSurface">
  <CustomTab id="TabCustom1">
    <Group id="ContosoCustomTab.grp1">
      <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
      <Control  xsi:type="Button" id="MyButton">
        <!-- Child elements omitted. -->
      </Control>
    </Group>
    <Label resid="customTabLabel1"/>
  </CustomTab>
</ExtensionPoint>
```

### <a name="overriding-a-control"></a><span data-ttu-id="cd223-118">コントロールのオーバーライド</span><span class="sxs-lookup"><span data-stu-id="cd223-118">Overriding a control</span></span>

```xml
<ExtensionPoint xsi:type="PrimaryCommandSurface">
  <CustomTab id="TabCustom1">
    <Group id="ContosoCustomTab.grp1">
      <Control  xsi:type="Button" id="MyButton">
        <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
        <!-- Other child elements omitted. -->
      </Control>
    </Group>
    <Label resid="customTabLabel1"/>
  </CustomTab>
</ExtensionPoint>
```

### <a name="overriding-a-menu-item"></a><span data-ttu-id="cd223-119">メニュー項目の上書き</span><span class="sxs-lookup"><span data-stu-id="cd223-119">Overriding a menu item</span></span>


```xml
<ExtensionPoint xsi:type="PrimaryCommandSurface">
  <CustomTab id="TabCustom1">
    <Group id="ContosoCustomTab.grp1">
      <Control  xsi:type="Menu" id="MyMenu">
        <!-- Other child elements omitted. -->
        <Items>
          <Item id="showGallery">
            <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
            <!-- Other child elements omitted. -->
          </Item>
        </Items>
      </Control>
    </Group>
    <Label resid="customTabLabel1"/>
  </CustomTab>
</ExtensionPoint>
```
