---
title: マニフェスト ファイルの SupportsSharedFolders 要素
description: ''
ms.date: 01/14/2020
localization_priority: Normal
ms.openlocfilehash: e76d17b618e2aaf15724f15ee6695a932172bba3
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/27/2020
ms.locfileid: "42325228"
---
# <a name="supportssharedfolders-element"></a><span data-ttu-id="c0252-102">SupportsSharedFolders 要素</span><span class="sxs-lookup"><span data-stu-id="c0252-102">SupportsSharedFolders element</span></span>

<span data-ttu-id="c0252-103">代理人のシナリオで Outlook アドインが使用できるかどうかを定義します。</span><span class="sxs-lookup"><span data-stu-id="c0252-103">Defines whether the Outlook add-in is available in delegate scenarios.</span></span> <span data-ttu-id="c0252-104">**SupportsSharedFolders** 要素は、[DesktopFormFactor](desktopformfactor.md) の子要素です。</span><span class="sxs-lookup"><span data-stu-id="c0252-104">The **SupportsSharedFolders** element is a child element of [DesktopFormFactor](desktopformfactor.md).</span></span> <span data-ttu-id="c0252-105">既定では *false* になっています。</span><span class="sxs-lookup"><span data-stu-id="c0252-105">It is set to *false* by default.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="c0252-106">Web 上の Outlook と Windows のみが**Supportssharedfolders**要素をサポートしています。</span><span class="sxs-lookup"><span data-stu-id="c0252-106">Only Outlook on the web and Windows support the **SupportsSharedFolders** element.</span></span>
>
> <span data-ttu-id="c0252-107">この要素のサポートは、要件セット1.8 で導入されました。</span><span class="sxs-lookup"><span data-stu-id="c0252-107">Support for this element was introduced in requirement set 1.8.</span></span> <span data-ttu-id="c0252-108">この要件セットをサポートする [クライアントおよびプラットフォーム](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets#requirement-sets-supported-by-exchange-servers-and-outlook-clients) を参照してください。</span><span class="sxs-lookup"><span data-stu-id="c0252-108">See [clients and platforms](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.</span></span>

<span data-ttu-id="c0252-109">**Supportssharedfolders**要素の例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="c0252-109">The following is an example of the **SupportsSharedFolders** element.</span></span>

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
