---
title: マニフェスト ファイルの ExtendedPermission 要素
description: アドインが関連付けられた API または機能にアクセスするために必要な拡張アクセス許可を定義します。
ms.date: 01/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: 5ed3745da87c2fa04839a8fbd1c677f62ad771dc
ms.sourcegitcommit: 9b0e70bb296a84adfaea0d6fee54916be9e13031
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/14/2022
ms.locfileid: "62042141"
---
# <a name="extendedpermission-element"></a>`ExtendedPermission` 要素

アドインが関連付けられた API または機能にアクセスするために必要な拡張アクセス許可を定義します。 要素 `ExtendedPermission` は [ExtendedPermissions の子要素です](extendedpermissions.md)。

> [!IMPORTANT]
> この要素のサポートは、要件セット 1.9 で導入されました。 この要件セットをサポートする [クライアントおよびプラットフォーム](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) を参照してください。

**アドインの種類:** メール

**次の VersionOverrides スキーマでのみ有効です**。

- メール 1.1

詳細については、「マニフェストの [バージョンオーバーライド」を参照してください](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)。

**次の要件セットに関連付けられている**。

- [Mailbox 1.9](../../reference/objectmodel/requirement-set-1.9/outlook-requirement-set-1.9.md)

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
