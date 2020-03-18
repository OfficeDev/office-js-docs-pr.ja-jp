---
title: マニフェスト ファイルの SupportsSharedFolders 要素
description: SupportsSharedFolders 要素は、Outlook アドインが代理人のシナリオで利用できるかどうかを定義します。
ms.date: 01/14/2020
localization_priority: Normal
ms.openlocfilehash: 66a426b0c31bda61feb23cb83d63722898dfb503
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42717888"
---
# <a name="supportssharedfolders-element"></a><span data-ttu-id="ec601-103">SupportsSharedFolders 要素</span><span class="sxs-lookup"><span data-stu-id="ec601-103">SupportsSharedFolders element</span></span>

<span data-ttu-id="ec601-104">代理人のシナリオで Outlook アドインが使用できるかどうかを定義します。</span><span class="sxs-lookup"><span data-stu-id="ec601-104">Defines whether the Outlook add-in is available in delegate scenarios.</span></span> <span data-ttu-id="ec601-105">**SupportsSharedFolders** 要素は、[DesktopFormFactor](desktopformfactor.md) の子要素です。</span><span class="sxs-lookup"><span data-stu-id="ec601-105">The **SupportsSharedFolders** element is a child element of [DesktopFormFactor](desktopformfactor.md).</span></span> <span data-ttu-id="ec601-106">既定では *false* になっています。</span><span class="sxs-lookup"><span data-stu-id="ec601-106">It is set to *false* by default.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="ec601-107">Web 上の Outlook と Windows のみが**Supportssharedfolders**要素をサポートしています。</span><span class="sxs-lookup"><span data-stu-id="ec601-107">Only Outlook on the web and Windows support the **SupportsSharedFolders** element.</span></span>
>
> <span data-ttu-id="ec601-108">この要素のサポートは、要件セット1.8 で導入されました。</span><span class="sxs-lookup"><span data-stu-id="ec601-108">Support for this element was introduced in requirement set 1.8.</span></span> <span data-ttu-id="ec601-109">この要件セットをサポートする [クライアントおよびプラットフォーム](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) を参照してください。</span><span class="sxs-lookup"><span data-stu-id="ec601-109">See [clients and platforms](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.</span></span>

<span data-ttu-id="ec601-110">**Supportssharedfolders**要素の例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="ec601-110">The following is an example of the **SupportsSharedFolders** element.</span></span>

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
            <!-- Configure selected extension point. -->
          </ExtensionPoint>

          <!-- You can define more than one ExtensionPoint element as needed. -->

        </DesktopFormFactor>
      </Host>
    </Hosts>
    ...
  </VersionOverrides>
</VersionOverrides>
...
```
