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
# <a name="supportssharedfolders-element"></a>SupportsSharedFolders 要素

代理人のシナリオで Outlook アドインが使用できるかどうかを定義します。 **SupportsSharedFolders** 要素は、[DesktopFormFactor](desktopformfactor.md) の子要素です。 既定では *false* になっています。

> [!IMPORTANT]
> Web 上の Outlook と Windows のみが**Supportssharedfolders**要素をサポートしています。
>
> この要素のサポートは、要件セット1.8 で導入されました。 この要件セットをサポートする [クライアントおよびプラットフォーム](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets#requirement-sets-supported-by-exchange-servers-and-outlook-clients) を参照してください。

**Supportssharedfolders**要素の例を次に示します。

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
