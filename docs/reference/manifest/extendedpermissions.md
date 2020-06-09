---
title: マニフェストファイルの ExtendedPermissions 要素
description: アドインが関連する Api または機能にアクセスするために必要な拡張アクセス許可のコレクションを定義します。
ms.date: 03/05/2020
localization_priority: Normal
ms.openlocfilehash: cf59d13d794f8f303da6cc0ca39066584bc3f56c
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611534"
---
# <a name="extendedpermissions-element"></a><span data-ttu-id="40059-103">ExtendedPermissions 要素</span><span class="sxs-lookup"><span data-stu-id="40059-103">ExtendedPermissions element</span></span>

<span data-ttu-id="40059-104">アドインが関連する Api または機能にアクセスするために必要な拡張アクセス許可のコレクションを定義します。</span><span class="sxs-lookup"><span data-stu-id="40059-104">Defines the collection of extended permissions the add-in needs to access associated APIs or features.</span></span> <span data-ttu-id="40059-105">`ExtendedPermissions`要素は[versionoverrides](versionoverrides.md)の子要素です。</span><span class="sxs-lookup"><span data-stu-id="40059-105">The `ExtendedPermissions` element is a child element of [VersionOverrides](versionoverrides.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="40059-106">この要素は、Exchange Online に対して[Outlook アドインのプレビュー要件が設定](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md)されている場合にのみ使用できます。</span><span class="sxs-lookup"><span data-stu-id="40059-106">This element is only available in the [Outlook add-ins preview requirement set](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) against Exchange Online.</span></span> <span data-ttu-id="40059-107">この要素を使用するアドインは、AppSource に発行したり、一元展開によって展開したりすることはできません。</span><span class="sxs-lookup"><span data-stu-id="40059-107">Add-ins that use this element cannot be published to AppSource or deployed via centralized deployment.</span></span>

## <a name="child-elements"></a><span data-ttu-id="40059-108">子要素</span><span class="sxs-lookup"><span data-stu-id="40059-108">Child elements</span></span>

|  <span data-ttu-id="40059-109">要素</span><span class="sxs-lookup"><span data-stu-id="40059-109">Element</span></span> |  <span data-ttu-id="40059-110">必須</span><span class="sxs-lookup"><span data-stu-id="40059-110">Required</span></span>  |  <span data-ttu-id="40059-111">説明</span><span class="sxs-lookup"><span data-stu-id="40059-111">Description</span></span>  |
|:-----|:-----:|:-----|
|  [<span data-ttu-id="40059-112">ExtendedPermission</span><span class="sxs-lookup"><span data-stu-id="40059-112">ExtendedPermission</span></span>](extendedpermission.md)    |  <span data-ttu-id="40059-113">いいえ</span><span class="sxs-lookup"><span data-stu-id="40059-113">No</span></span>   | <span data-ttu-id="40059-114">アドインが関連する API または機能にアクセスするために必要な拡張アクセス許可を定義します。</span><span class="sxs-lookup"><span data-stu-id="40059-114">Defines an extended permission needed for the add-in to access the associated API or feature.</span></span> |

## <a name="extendedpermissions-example"></a><span data-ttu-id="40059-115">`ExtendedPermissions`例</span><span class="sxs-lookup"><span data-stu-id="40059-115">`ExtendedPermissions` example</span></span>

<span data-ttu-id="40059-116">要素の例を次に示し `ExtendedPermissions` ます。</span><span class="sxs-lookup"><span data-stu-id="40059-116">The following is an example of the `ExtendedPermissions` element.</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="40059-117">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="40059-117">Contained in</span></span>

[<span data-ttu-id="40059-118">VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="40059-118">VersionOverrides</span></span>](versionoverrides.md)

## <a name="can-contain"></a><span data-ttu-id="40059-119">含めることができるもの</span><span class="sxs-lookup"><span data-stu-id="40059-119">Can contain</span></span>

[<span data-ttu-id="40059-120">ExtendedPermission</span><span class="sxs-lookup"><span data-stu-id="40059-120">ExtendedPermission</span></span>](extendedpermission.md)
