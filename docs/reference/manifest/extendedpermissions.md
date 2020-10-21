---
title: マニフェストファイルの ExtendedPermissions 要素
description: アドインが関連する Api または機能にアクセスするために必要な拡張アクセス許可のコレクションを定義します。
ms.date: 10/15/2020
localization_priority: Normal
ms.openlocfilehash: 1e3aa16c160613d34ef2c4f9c25bc2ffe4970816
ms.sourcegitcommit: 4e7c74ad67ea8bf6b47d65b2fde54a967090f65b
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/20/2020
ms.locfileid: "48626443"
---
# <a name="extendedpermissions-element"></a><span data-ttu-id="044ec-103">ExtendedPermissions 要素</span><span class="sxs-lookup"><span data-stu-id="044ec-103">ExtendedPermissions element</span></span>

<span data-ttu-id="044ec-104">アドインが関連する Api または機能にアクセスするために必要な拡張アクセス許可のコレクションを定義します。</span><span class="sxs-lookup"><span data-stu-id="044ec-104">Defines the collection of extended permissions the add-in needs to access associated APIs or features.</span></span> <span data-ttu-id="044ec-105">`ExtendedPermissions`要素は[versionoverrides](versionoverrides.md)の子要素です。</span><span class="sxs-lookup"><span data-stu-id="044ec-105">The `ExtendedPermissions` element is a child element of [VersionOverrides](versionoverrides.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="044ec-106">この要素のサポートは、要件セット1.9 で導入されました。</span><span class="sxs-lookup"><span data-stu-id="044ec-106">Support for this element was introduced in requirement set 1.9.</span></span> <span data-ttu-id="044ec-107">この要件セットをサポートする [クライアントおよびプラットフォーム](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) を参照してください。</span><span class="sxs-lookup"><span data-stu-id="044ec-107">See [clients and platforms](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.</span></span>

## <a name="child-elements"></a><span data-ttu-id="044ec-108">子要素</span><span class="sxs-lookup"><span data-stu-id="044ec-108">Child elements</span></span>

|  <span data-ttu-id="044ec-109">要素</span><span class="sxs-lookup"><span data-stu-id="044ec-109">Element</span></span> |  <span data-ttu-id="044ec-110">必須</span><span class="sxs-lookup"><span data-stu-id="044ec-110">Required</span></span>  |  <span data-ttu-id="044ec-111">説明</span><span class="sxs-lookup"><span data-stu-id="044ec-111">Description</span></span>  |
|:-----|:-----:|:-----|
|  [<span data-ttu-id="044ec-112">ExtendedPermission</span><span class="sxs-lookup"><span data-stu-id="044ec-112">ExtendedPermission</span></span>](extendedpermission.md)    |  <span data-ttu-id="044ec-113">いいえ</span><span class="sxs-lookup"><span data-stu-id="044ec-113">No</span></span>   | <span data-ttu-id="044ec-114">アドインが関連する API または機能にアクセスするために必要な拡張アクセス許可を定義します。</span><span class="sxs-lookup"><span data-stu-id="044ec-114">Defines an extended permission needed for the add-in to access the associated API or feature.</span></span> |

## <a name="extendedpermissions-example"></a><span data-ttu-id="044ec-115">`ExtendedPermissions` 例</span><span class="sxs-lookup"><span data-stu-id="044ec-115">`ExtendedPermissions` example</span></span>

<span data-ttu-id="044ec-116">要素の例を次に示し `ExtendedPermissions` ます。</span><span class="sxs-lookup"><span data-stu-id="044ec-116">The following is an example of the `ExtendedPermissions` element.</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="044ec-117">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="044ec-117">Contained in</span></span>

[<span data-ttu-id="044ec-118">VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="044ec-118">VersionOverrides</span></span>](versionoverrides.md)

## <a name="can-contain"></a><span data-ttu-id="044ec-119">含めることができるもの</span><span class="sxs-lookup"><span data-stu-id="044ec-119">Can contain</span></span>

[<span data-ttu-id="044ec-120">ExtendedPermission</span><span class="sxs-lookup"><span data-stu-id="044ec-120">ExtendedPermission</span></span>](extendedpermission.md)
