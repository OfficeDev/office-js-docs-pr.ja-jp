---
title: マニフェストファイルの ExtendedPermissions 要素
description: アドインが関連する Api または機能にアクセスするために必要な拡張アクセス許可のコレクションを定義します。
ms.date: 03/05/2020
localization_priority: Normal
ms.openlocfilehash: 86d898052af6ba0e6f6bc8b341fff9f0f8408967
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718224"
---
# <a name="extendedpermissions-element"></a><span data-ttu-id="d2d74-103">ExtendedPermissions 要素</span><span class="sxs-lookup"><span data-stu-id="d2d74-103">ExtendedPermissions element</span></span>

<span data-ttu-id="d2d74-104">アドインが関連する Api または機能にアクセスするために必要な拡張アクセス許可のコレクションを定義します。</span><span class="sxs-lookup"><span data-stu-id="d2d74-104">Defines the collection of extended permissions the add-in needs to access associated APIs or features.</span></span> <span data-ttu-id="d2d74-105">`ExtendedPermissions`要素は[versionoverrides](versionoverrides.md)の子要素です。</span><span class="sxs-lookup"><span data-stu-id="d2d74-105">The `ExtendedPermissions` element is a child element of [VersionOverrides](versionoverrides.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="d2d74-106">この要素は、Exchange Online に対して[Outlook アドインのプレビュー要件が設定](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md)されている場合にのみ使用できます。</span><span class="sxs-lookup"><span data-stu-id="d2d74-106">This element is only available in the [Outlook add-ins preview requirement set](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) against Exchange Online.</span></span> <span data-ttu-id="d2d74-107">この要素を使用するアドインは、AppSource に発行したり、一元展開によって展開したりすることはできません。</span><span class="sxs-lookup"><span data-stu-id="d2d74-107">Add-ins that use this element cannot be published to AppSource or deployed via centralized deployment.</span></span>

## <a name="child-elements"></a><span data-ttu-id="d2d74-108">子要素</span><span class="sxs-lookup"><span data-stu-id="d2d74-108">Child elements</span></span>

|  <span data-ttu-id="d2d74-109">要素</span><span class="sxs-lookup"><span data-stu-id="d2d74-109">Element</span></span> |  <span data-ttu-id="d2d74-110">必須</span><span class="sxs-lookup"><span data-stu-id="d2d74-110">Required</span></span>  |  <span data-ttu-id="d2d74-111">説明</span><span class="sxs-lookup"><span data-stu-id="d2d74-111">Description</span></span>  |
|:-----|:-----:|:-----|
|  [<span data-ttu-id="d2d74-112">ExtendedPermission</span><span class="sxs-lookup"><span data-stu-id="d2d74-112">ExtendedPermission</span></span>](extendedpermission.md)    |  <span data-ttu-id="d2d74-113">いいえ</span><span class="sxs-lookup"><span data-stu-id="d2d74-113">No</span></span>   | <span data-ttu-id="d2d74-114">アドインが関連する API または機能にアクセスするために必要な拡張アクセス許可を定義します。</span><span class="sxs-lookup"><span data-stu-id="d2d74-114">Defines an extended permission needed for the add-in to access the associated API or feature.</span></span> |

## <a name="extendedpermissions-example"></a><span data-ttu-id="d2d74-115">`ExtendedPermissions`例</span><span class="sxs-lookup"><span data-stu-id="d2d74-115">`ExtendedPermissions` example</span></span>

<span data-ttu-id="d2d74-116">`ExtendedPermissions`要素の例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="d2d74-116">The following is an example of the `ExtendedPermissions` element.</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="d2d74-117">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="d2d74-117">Contained in</span></span>

[<span data-ttu-id="d2d74-118">VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="d2d74-118">VersionOverrides</span></span>](versionoverrides.md)

## <a name="can-contain"></a><span data-ttu-id="d2d74-119">含めることができるもの</span><span class="sxs-lookup"><span data-stu-id="d2d74-119">Can contain</span></span>

[<span data-ttu-id="d2d74-120">ExtendedPermission</span><span class="sxs-lookup"><span data-stu-id="d2d74-120">ExtendedPermission</span></span>](extendedpermission.md)
