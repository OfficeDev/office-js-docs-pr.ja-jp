---
title: マニフェストファイルの ExtendedPermissions 要素
description: ''
ms.date: 03/05/2020
localization_priority: Normal
ms.openlocfilehash: 966378b8bbed66960d7a99c4a82df75ace1c9161
ms.sourcegitcommit: a0262ea40cd23f221e69bcb0223110f011265d13
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/12/2020
ms.locfileid: "42605815"
---
# <a name="extendedpermissions-element"></a>ExtendedPermissions 要素

アドインが関連する Api または機能にアクセスするために必要な拡張アクセス許可のコレクションを定義します。 `ExtendedPermissions`要素は[versionoverrides](versionoverrides.md)の子要素です。

> [!IMPORTANT]
> この要素は、Exchange Online に対して[Outlook アドインのプレビュー要件が設定](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md)されている場合にのみ使用できます。 この要素を使用するアドインは、AppSource に発行したり、一元展開によって展開したりすることはできません。

## <a name="child-elements"></a>子要素

|  要素 |  必須  |  説明  |
|:-----|:-----:|:-----|
|  [ExtendedPermission](extendedpermission.md)    |  いいえ   | アドインが関連する API または機能にアクセスするために必要な拡張アクセス許可を定義します。 |

## <a name="extendedpermissions-example"></a>`ExtendedPermissions`例

`ExtendedPermissions`要素の例を次に示します。

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
