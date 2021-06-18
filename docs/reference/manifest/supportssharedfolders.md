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
# <a name="supportssharedfolders-element"></a>SupportsSharedFolders 要素

共有メールボックス (プレビュー Outlook共有フォルダー (つまり、代理アクセス) のシナリオで、アドインを使用できるかどうかを定義します。 **SupportsSharedFolders** 要素は、[DesktopFormFactor](desktopformfactor.md) の子要素です。 既定では *false* になっています。

> [!IMPORTANT]
> この要素のサポートは、要件セット 1.8 で導入されました。 この要件セットをサポートする [クライアントおよびプラットフォーム](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) を参照してください。

**SupportsSharedFolders 要素の例を次に示** します。

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
