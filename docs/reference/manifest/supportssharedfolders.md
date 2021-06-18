---
title: マニフェスト ファイルの SupportsSharedFolders 要素
description: SupportsSharedFolders 要素は、共有フォルダー Outlook共有メールボックス のシナリオで使用できるかどうかを定義します。
ms.date: 06/15/2021
localization_priority: Normal
ms.openlocfilehash: 43f2c60664a6822b714023246cfa044e179e9a55
ms.sourcegitcommit: 0bf0e076f705af29193abe3dba98cbfcce17b24f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/18/2021
ms.locfileid: "53007784"
---
# <a name="supportssharedfolders-element"></a><span data-ttu-id="1e697-103">SupportsSharedFolders 要素</span><span class="sxs-lookup"><span data-stu-id="1e697-103">SupportsSharedFolders element</span></span>

<span data-ttu-id="1e697-104">共有メールボックス (プレビュー Outlook共有フォルダー (つまり、代理アクセス) のシナリオで、アドインを使用できるかどうかを定義します。</span><span class="sxs-lookup"><span data-stu-id="1e697-104">Defines whether the Outlook add-in is available in shared mailbox (now in preview) and shared folders (that is, delegate access) scenarios.</span></span> <span data-ttu-id="1e697-105">**SupportsSharedFolders** 要素は、[DesktopFormFactor](desktopformfactor.md) の子要素です。</span><span class="sxs-lookup"><span data-stu-id="1e697-105">The **SupportsSharedFolders** element is a child element of [DesktopFormFactor](desktopformfactor.md).</span></span> <span data-ttu-id="1e697-106">既定では *false* になっています。</span><span class="sxs-lookup"><span data-stu-id="1e697-106">It is set to *false* by default.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="1e697-107">この要素のサポートは、要件セット 1.8 で導入されました。</span><span class="sxs-lookup"><span data-stu-id="1e697-107">Support for this element was introduced in requirement set 1.8.</span></span> <span data-ttu-id="1e697-108">この要件セットをサポートする [クライアントおよびプラットフォーム](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) を参照してください。</span><span class="sxs-lookup"><span data-stu-id="1e697-108">See [clients and platforms](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.</span></span>

<span data-ttu-id="1e697-109">**SupportsSharedFolders 要素の例を次に示** します。</span><span class="sxs-lookup"><span data-stu-id="1e697-109">The following is an example of the **SupportsSharedFolders** element.</span></span>

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
