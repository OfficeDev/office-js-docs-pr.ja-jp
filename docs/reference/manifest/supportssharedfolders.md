---
title: マニフェスト ファイルの SupportsSharedFolders 要素
description: SupportsSharedFolders 要素は、共有フォルダー Outlook共有メールボックス のシナリオで使用できるかどうかを定義します。
ms.date: 09/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: 8e13393f10b12e0a3c5ca1b004b202eb2970d264
ms.sourcegitcommit: 489befc41e543a4fb3c504fd9b3f61322134c1ef
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/06/2021
ms.locfileid: "60138724"
---
# <a name="supportssharedfolders-element"></a>SupportsSharedFolders 要素

共有メールボックス (プレビュー Outlook共有フォルダー (つまり、代理アクセス) のシナリオで、アドインを使用できるかどうかを定義します。 **SupportsSharedFolders** 要素は、[DesktopFormFactor](desktopformfactor.md) の子要素です。 既定では *false* になっています。

> [!IMPORTANT]
> この要素のサポートは、要件セット 1.8 で導入されました。 この要件セットをサポートする [クライアントおよびプラットフォーム](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) を参照してください。

**アドインの種類:** メール

**次の VersionOverrides スキーマでのみ有効です**。

- メール 1.1

詳細については、「マニフェストの [バージョンオーバーライド」を参照してください](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)。

**次の要件セットに関連付けられている**。

- [Mailbox 1.8](../../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md)

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
