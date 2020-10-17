---
title: マニフェスト ファイルの SupportsSharedFolders 要素
description: SupportsSharedFolders 要素は、Outlook アドインが代理人のシナリオで利用できるかどうかを定義します。
ms.date: 10/15/2020
localization_priority: Normal
ms.openlocfilehash: 786a4763450d78cb16c9baafc81701758af54787
ms.sourcegitcommit: 6fa29989dfaec4dfa0f8df3fe5fb038d7afbae30
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/16/2020
ms.locfileid: "48487881"
---
# <a name="supportssharedfolders-element"></a><span data-ttu-id="cb06b-103">SupportsSharedFolders 要素</span><span class="sxs-lookup"><span data-stu-id="cb06b-103">SupportsSharedFolders element</span></span>

<span data-ttu-id="cb06b-104">代理人のシナリオで Outlook アドインが使用できるかどうかを定義します。</span><span class="sxs-lookup"><span data-stu-id="cb06b-104">Defines whether the Outlook add-in is available in delegate scenarios.</span></span> <span data-ttu-id="cb06b-105">**SupportsSharedFolders** 要素は、[DesktopFormFactor](desktopformfactor.md) の子要素です。</span><span class="sxs-lookup"><span data-stu-id="cb06b-105">The **SupportsSharedFolders** element is a child element of [DesktopFormFactor](desktopformfactor.md).</span></span> <span data-ttu-id="cb06b-106">既定では *false* になっています。</span><span class="sxs-lookup"><span data-stu-id="cb06b-106">It is set to *false* by default.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="cb06b-107">この要素のサポートは、要件セット1.8 で導入されました。</span><span class="sxs-lookup"><span data-stu-id="cb06b-107">Support for this element was introduced in requirement set 1.8.</span></span> <span data-ttu-id="cb06b-108">この要件セットをサポートする [クライアントおよびプラットフォーム](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) を参照してください。</span><span class="sxs-lookup"><span data-stu-id="cb06b-108">See [clients and platforms](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.</span></span>

<span data-ttu-id="cb06b-109">**Supportssharedfolders**要素の例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="cb06b-109">The following is an example of the **SupportsSharedFolders** element.</span></span>

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
