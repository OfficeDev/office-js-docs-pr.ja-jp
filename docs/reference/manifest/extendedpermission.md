---
title: マニフェスト ファイルの ExtendedPermission 要素
description: アドインが関連付けられた API または機能にアクセスするために必要な拡張アクセス許可を定義します。
ms.date: 10/15/2020
localization_priority: Normal
ms.openlocfilehash: 37859350cfaffdc14ab91d5026d67aa0a736ac56
ms.sourcegitcommit: 3fa8c754a47bab909e559ae3e5d4237ba27fdbe4
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/30/2021
ms.locfileid: "53671759"
---
# <a name="extendedpermission-element"></a>`ExtendedPermission` 要素

アドインが関連付けられた API または機能にアクセスするために必要な拡張アクセス許可を定義します。 要素 `ExtendedPermission` は [ExtendedPermissions の子要素です](extendedpermissions.md)。

> [!IMPORTANT]
> この要素のサポートは、要件セット 1.9 で導入されました。 この要件セットをサポートする [クライアントおよびプラットフォーム](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) を参照してください。

## <a name="available-extended-permissions"></a>使用可能な拡張アクセス許可

使用可能な値は次のとおりです。

|使用可能な値|説明|Hosts|
|---|---|---|
|`AppendOnSend`|アドインがアプリケーション を使用[Office。Body.appendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#appendOnSendAsync_data__options__callback_) API。|Outlook|

## <a name="extendedpermission-example"></a>`ExtendedPermission` 例

次に、要素の例を示 `ExtendedPermission` します。

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
    <ExtendedPermissions>
      <ExtendedPermission>AppendOnSend</ExtendedPermission>
    </ExtendedPermissions>
  </VersionOverrides>
</VersionOverrides>
...
```

## <a name="contained-in"></a>含まれる場所

[ExtendedPermissions](extendedpermissions.md)
