---
title: マニフェストファイルの ExtendedPermission 要素
description: アドインが関連する API または機能にアクセスするために必要な拡張アクセス許可を定義します。
ms.date: 03/05/2020
localization_priority: Normal
ms.openlocfilehash: 138acafb359e2b6e386b34fde7201b1b2c4b3177
ms.sourcegitcommit: 83f9a2fdff81ca421cd23feea103b9b60895cab4
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/11/2020
ms.locfileid: "47430927"
---
# <a name="extendedpermission-element"></a>`ExtendedPermission` 項目

アドインが関連する API または機能にアクセスするために必要な拡張アクセス許可を定義します。 `ExtendedPermission`要素は、 [extendedpermissions](extendedpermissions.md)の子要素です。

> [!IMPORTANT]
> この要素は、Exchange Online に対して [Outlook アドインのプレビュー要件が設定](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) されている場合にのみ使用できます。 この要素を使用するアドインは、AppSource に発行したり、一元展開によって展開したりすることはできません。

## <a name="available-extended-permissions"></a>利用可能な拡張アクセス許可

使用可能な値は次のとおりです。

|利用可能な値|説明|Hosts|
|---|---|---|
|`AppendOnSend`|アドインが [Office. appendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#appendonsendasync-data--options--callback-) API を使用していることを宣言します。|Outlook|

## <a name="extendedpermission-example"></a>`ExtendedPermission` 例

要素の例を次に示し `ExtendedPermission` ます。

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
