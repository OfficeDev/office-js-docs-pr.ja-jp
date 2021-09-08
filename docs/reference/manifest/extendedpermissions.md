---
title: マニフェスト ファイルの ExtendedPermissions 要素
description: アドインが関連付けられた API または機能にアクセスするために必要な拡張アクセス許可のコレクションを定義します。
ms.date: 10/15/2020
localization_priority: Normal
ms.openlocfilehash: 1e3aa16c160613d34ef2c4f9c25bc2ffe4970816
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/08/2021
ms.locfileid: "58938851"
---
# <a name="extendedpermissions-element"></a>ExtendedPermissions 要素

アドインが関連付けられた API または機能にアクセスするために必要な拡張アクセス許可のコレクションを定義します。 要素 `ExtendedPermissions` は [VersionOverrides の子要素です](versionoverrides.md)。

> [!IMPORTANT]
> この要素のサポートは、要件セット 1.9 で導入されました。 この要件セットをサポートする [クライアントおよびプラットフォーム](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) を参照してください。

## <a name="child-elements"></a>子要素

|  要素 |  必須  |  説明  |
|:-----|:-----:|:-----|
|  [ExtendedPermission](extendedpermission.md)    |  いいえ   | アドインが関連付けられた API または機能にアクセスするために必要な拡張アクセス許可を定義します。 |

## <a name="extendedpermissions-example"></a>`ExtendedPermissions` 例

次に、要素の例を示 `ExtendedPermissions` します。

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

[VersionOverrides](versionoverrides.md)

## <a name="can-contain"></a>含めることができるもの

[ExtendedPermission](extendedpermission.md)
