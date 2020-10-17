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
# <a name="supportssharedfolders-element"></a>SupportsSharedFolders 要素

代理人のシナリオで Outlook アドインが使用できるかどうかを定義します。 **SupportsSharedFolders** 要素は、[DesktopFormFactor](desktopformfactor.md) の子要素です。 既定では *false* になっています。

> [!IMPORTANT]
> この要素のサポートは、要件セット1.8 で導入されました。 この要件セットをサポートする [クライアントおよびプラットフォーム](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) を参照してください。

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
