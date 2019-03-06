---
title: マニフェスト ファイルの SupportsSharedFolders 要素
description: ''
ms.date: 03/01/2019
localization_priority: Normal
ms.openlocfilehash: bfbce42c7d1aa5eefab40b528c5b622aa7d2d54f
ms.sourcegitcommit: 7ebd383f16ae5809bb6980a5f213b695d410e62c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/06/2019
ms.locfileid: "30413616"
---
# <a name="supportssharedfolders-element"></a><span data-ttu-id="b9996-102">SupportsSharedFolders 要素</span><span class="sxs-lookup"><span data-stu-id="b9996-102">SupportsSharedFolders element</span></span>

<span data-ttu-id="b9996-103">代理人のシナリオで Outlook アドインが使用できるかどうかを定義します。</span><span class="sxs-lookup"><span data-stu-id="b9996-103">Defines whether the Outlook add-in is available in delegate scenarios.</span></span> <span data-ttu-id="b9996-104">**SupportsSharedFolders** 要素は、[DesktopFormFactor](desktopformfactor.md) の子要素です。</span><span class="sxs-lookup"><span data-stu-id="b9996-104">The **SupportsSharedFolders** element is a child element of [DesktopFormFactor](desktopformfactor.md).</span></span> <span data-ttu-id="b9996-105">既定では *false* になっています。</span><span class="sxs-lookup"><span data-stu-id="b9996-105">It is set to *false* by default.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="b9996-106">Outlook アドインの代理人アクセスは現在プレビュー段階であり、Exchange Online に対して実行されるクライアントでのみサポートされています。</span><span class="sxs-lookup"><span data-stu-id="b9996-106">Delegate access for Outlook add-ins is currently in preview and only supported in clients that run against Exchange Online.</span></span> <span data-ttu-id="b9996-107">この要素を使用するアドインは、AppSource に発行したり、一元展開によって展開したりすることはできません。</span><span class="sxs-lookup"><span data-stu-id="b9996-107">Add-ins that use this element cannot be published to AppSource or deployed via centralized deployment.</span></span>

<span data-ttu-id="b9996-108">**SupportsSharedFolders** 要素の例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="b9996-108">The following is an example of the  **SupportsSharedFolders** element.</span></span>

```XML
<DesktopFormFactor>
  <FunctionFile resid="residDesktopFuncUrl" />
  <SupportsSharedFolders>true</SupportsSharedFolders>
  <ExtensionPoint xsi:type="MessageReadCommandSurface">
    <!-- configure selected extension point -->
  </ExtensionPoint>

  <!-- You can define more than one ExtensionPoint element as needed -->

</DesktopFormFactor>
```
