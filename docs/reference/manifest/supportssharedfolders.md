---
title: マニフェスト ファイルの SupportsSharedFolders 要素
description: ''
ms.date: 01/14/2020
localization_priority: Normal
ms.openlocfilehash: 81401b79f4c443305e376df7a66a07d916393d17
ms.sourcegitcommit: 4079903c3cc45b7d8c041509a44e9fc38da399b1
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/11/2020
ms.locfileid: "42596754"
---
# <a name="supportssharedfolders-element"></a><span data-ttu-id="e560b-102">SupportsSharedFolders 要素</span><span class="sxs-lookup"><span data-stu-id="e560b-102">SupportsSharedFolders element</span></span>

<span data-ttu-id="e560b-103">代理人のシナリオで Outlook アドインが使用できるかどうかを定義します。</span><span class="sxs-lookup"><span data-stu-id="e560b-103">Defines whether the Outlook add-in is available in delegate scenarios.</span></span> <span data-ttu-id="e560b-104">**SupportsSharedFolders** 要素は、[DesktopFormFactor](desktopformfactor.md) の子要素です。</span><span class="sxs-lookup"><span data-stu-id="e560b-104">The **SupportsSharedFolders** element is a child element of [DesktopFormFactor](desktopformfactor.md).</span></span> <span data-ttu-id="e560b-105">既定では *false* になっています。</span><span class="sxs-lookup"><span data-stu-id="e560b-105">It is set to *false* by default.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="e560b-106">Web 上の Outlook と Windows のみが**Supportssharedfolders**要素をサポートしています。</span><span class="sxs-lookup"><span data-stu-id="e560b-106">Only Outlook on the web and Windows support the **SupportsSharedFolders** element.</span></span>
>
> <span data-ttu-id="e560b-107">この要素のサポートは、要件セット1.8 で導入されました。</span><span class="sxs-lookup"><span data-stu-id="e560b-107">Support for this element was introduced in requirement set 1.8.</span></span> <span data-ttu-id="e560b-108">この要件セットをサポートする [クライアントおよびプラットフォーム](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) を参照してください。</span><span class="sxs-lookup"><span data-stu-id="e560b-108">See [clients and platforms](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.</span></span>

<span data-ttu-id="e560b-109">**Supportssharedfolders**要素の例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="e560b-109">The following is an example of the **SupportsSharedFolders** element.</span></span>

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
