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
# <a name="extendedpermissions-element"></a><span data-ttu-id="c5fe6-102">ExtendedPermissions 要素</span><span class="sxs-lookup"><span data-stu-id="c5fe6-102">ExtendedPermissions element</span></span>

<span data-ttu-id="c5fe6-103">アドインが関連する Api または機能にアクセスするために必要な拡張アクセス許可のコレクションを定義します。</span><span class="sxs-lookup"><span data-stu-id="c5fe6-103">Defines the collection of extended permissions the add-in needs to access associated APIs or features.</span></span> <span data-ttu-id="c5fe6-104">`ExtendedPermissions`要素は[versionoverrides](versionoverrides.md)の子要素です。</span><span class="sxs-lookup"><span data-stu-id="c5fe6-104">The `ExtendedPermissions` element is a child element of [VersionOverrides](versionoverrides.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="c5fe6-105">この要素は、Exchange Online に対して[Outlook アドインのプレビュー要件が設定](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md)されている場合にのみ使用できます。</span><span class="sxs-lookup"><span data-stu-id="c5fe6-105">This element is only available in the [Outlook add-ins preview requirement set](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) against Exchange Online.</span></span> <span data-ttu-id="c5fe6-106">この要素を使用するアドインは、AppSource に発行したり、一元展開によって展開したりすることはできません。</span><span class="sxs-lookup"><span data-stu-id="c5fe6-106">Add-ins that use this element cannot be published to AppSource or deployed via centralized deployment.</span></span>

## <a name="child-elements"></a><span data-ttu-id="c5fe6-107">子要素</span><span class="sxs-lookup"><span data-stu-id="c5fe6-107">Child elements</span></span>

|  <span data-ttu-id="c5fe6-108">要素</span><span class="sxs-lookup"><span data-stu-id="c5fe6-108">Element</span></span> |  <span data-ttu-id="c5fe6-109">必須</span><span class="sxs-lookup"><span data-stu-id="c5fe6-109">Required</span></span>  |  <span data-ttu-id="c5fe6-110">説明</span><span class="sxs-lookup"><span data-stu-id="c5fe6-110">Description</span></span>  |
|:-----|:-----:|:-----|
|  [<span data-ttu-id="c5fe6-111">ExtendedPermission</span><span class="sxs-lookup"><span data-stu-id="c5fe6-111">ExtendedPermission</span></span>](extendedpermission.md)    |  <span data-ttu-id="c5fe6-112">いいえ</span><span class="sxs-lookup"><span data-stu-id="c5fe6-112">No</span></span>   | <span data-ttu-id="c5fe6-113">アドインが関連する API または機能にアクセスするために必要な拡張アクセス許可を定義します。</span><span class="sxs-lookup"><span data-stu-id="c5fe6-113">Defines an extended permission needed for the add-in to access the associated API or feature.</span></span> |

## <a name="extendedpermissions-example"></a><span data-ttu-id="c5fe6-114">`ExtendedPermissions`例</span><span class="sxs-lookup"><span data-stu-id="c5fe6-114">`ExtendedPermissions` example</span></span>

<span data-ttu-id="c5fe6-115">`ExtendedPermissions`要素の例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="c5fe6-115">The following is an example of the `ExtendedPermissions` element.</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="c5fe6-116">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="c5fe6-116">Contained in</span></span>

[<span data-ttu-id="c5fe6-117">VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="c5fe6-117">VersionOverrides</span></span>](versionoverrides.md)

## <a name="can-contain"></a><span data-ttu-id="c5fe6-118">含めることができるもの</span><span class="sxs-lookup"><span data-stu-id="c5fe6-118">Can contain</span></span>

[<span data-ttu-id="c5fe6-119">ExtendedPermission</span><span class="sxs-lookup"><span data-stu-id="c5fe6-119">ExtendedPermission</span></span>](extendedpermission.md)
