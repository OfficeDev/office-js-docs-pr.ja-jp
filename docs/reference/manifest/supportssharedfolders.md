---
title: マニフェスト ファイルの SupportsSharedFolders 要素
description: ''
ms.date: 04/02/2019
localization_priority: Normal
ms.openlocfilehash: 976f8ba00f6ac9ac32def56933af1077527b7e9c
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/24/2019
ms.locfileid: "32452040"
---
# <a name="supportssharedfolders-element"></a><span data-ttu-id="6f8c5-102">SupportsSharedFolders 要素</span><span class="sxs-lookup"><span data-stu-id="6f8c5-102">SupportsSharedFolders element</span></span>

<span data-ttu-id="6f8c5-103">代理人のシナリオで Outlook アドインが使用できるかどうかを定義します。</span><span class="sxs-lookup"><span data-stu-id="6f8c5-103">Defines whether the Outlook add-in is available in delegate scenarios.</span></span> <span data-ttu-id="6f8c5-104">**SupportsSharedFolders** 要素は、[DesktopFormFactor](desktopformfactor.md) の子要素です。</span><span class="sxs-lookup"><span data-stu-id="6f8c5-104">The **SupportsSharedFolders** element is a child element of [DesktopFormFactor](desktopformfactor.md).</span></span> <span data-ttu-id="6f8c5-105">既定では *false* になっています。</span><span class="sxs-lookup"><span data-stu-id="6f8c5-105">It is set to *false* by default.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="6f8c5-106">Outlook アドインの代理人アクセスは現在[プレビュー段階で](/office/dev/add-ins/reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview)あり、Exchange Online に対して実行されるクライアントでのみサポートされています。</span><span class="sxs-lookup"><span data-stu-id="6f8c5-106">Delegate access for Outlook add-ins is currently [in preview](/office/dev/add-ins/reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview) and only supported in clients that run against Exchange Online.</span></span> <span data-ttu-id="6f8c5-107">この要素を使用するアドインは、AppSource に発行したり、一元展開によって展開したりすることはできません。</span><span class="sxs-lookup"><span data-stu-id="6f8c5-107">Add-ins that use this element cannot be published to AppSource or deployed via centralized deployment.</span></span>

<span data-ttu-id="6f8c5-108">**SupportsSharedFolders** 要素の例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="6f8c5-108">The following is an example of the  **SupportsSharedFolders** element.</span></span>

```XML
...
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
    ...
    <Hosts>
      <Host xsi:type="MailHost">
        <DesktopFormFactor>
          <SupportsSharedFolders>true</SupportsSharedFolders>
          <FunctionFile resid="residDesktopFuncUrl" />
          <ExtensionPoint xsi:type="MessageReadCommandSurface">
            <!-- configure selected extension point -->
          </ExtensionPoint>

          <!-- You can define more than one ExtensionPoint element as needed -->

        </DesktopFormFactor>
      </Host>
    </Hosts>
    ...
  </VersionOverrides>
</VersionOverrides>
...
```
