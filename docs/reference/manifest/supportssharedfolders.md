---
title: マニフェスト ファイルの SupportsSharedFolders 要素
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 776d44ec66c4e27a72e5487051bed1edf4b3dcaf
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432684"
---
# <a name="supportssharedfolders-element"></a><span data-ttu-id="3ad47-102">SupportsSharedFolders 要素</span><span class="sxs-lookup"><span data-stu-id="3ad47-102">SupportsSharedFolders element</span></span>

<span data-ttu-id="3ad47-103">代理人のシナリオで Outlook アドインが使用できるかどうかを定義します。</span><span class="sxs-lookup"><span data-stu-id="3ad47-103">It defines whether the add-in is available in delegate scenarios.</span></span> <span data-ttu-id="3ad47-104">**SupportsSharedFolders** 要素は、[DesktopFormFactor](desktopformfactor.md) の子要素です。</span><span class="sxs-lookup"><span data-stu-id="3ad47-104">The **ExtensionPoint** element is a child element of [AllFormFactors, DesktopFormFactor or MobileFormFactor](desktopformfactor.md).</span></span> <span data-ttu-id="3ad47-105">既定では *false* になっています。</span><span class="sxs-lookup"><span data-stu-id="3ad47-105">It is set to *false* by default.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="3ad47-106">この要素は、[Outlook アドイン要件セットのプレビュー](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md)の Exchange Online に対してのみ使用できます。</span><span class="sxs-lookup"><span data-stu-id="3ad47-106">This element is only available in the [Outlook add-ins Preview requirement set](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) against Exchange Online.</span></span> <span data-ttu-id="3ad47-107">この要素を使用するアドインは、AppSource に発行したり、一元展開によって展開したりすることはできません。</span><span class="sxs-lookup"><span data-stu-id="3ad47-107">Add-ins that use this element cannot be published to AppSource or deployed via centralized deployment.</span></span>

<span data-ttu-id="3ad47-108">**SupportsSharedFolders** 要素の例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="3ad47-108">The following is an example of how the **Rows** element should look.</span></span>

```XML
<DesktopFormFactor>
  <FunctionFile resid="residDesktopFuncUrl" />
  <SupportsSharedFolders>true</SupportsSharedFolders>
  <ExtensionPoint xsi:type="PrimaryCommandSurface">
    <!-- information about this extension point -->
  </ExtensionPoint>

  <!-- You can define more than one ExtensionPoint element as needed -->

</DesktopFormFactor>
```
