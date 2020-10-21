---
title: マニフェストファイルの ExtendedPermission 要素
description: アドインが関連する API または機能にアクセスするために必要な拡張アクセス許可を定義します。
ms.date: 10/15/2020
localization_priority: Normal
ms.openlocfilehash: 996cac59c44220d05165c7be6ae7c3d79d853271
ms.sourcegitcommit: 4e7c74ad67ea8bf6b47d65b2fde54a967090f65b
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/20/2020
ms.locfileid: "48626401"
---
# <a name="extendedpermission-element"></a><span data-ttu-id="fcda5-103">`ExtendedPermission` 項目</span><span class="sxs-lookup"><span data-stu-id="fcda5-103">`ExtendedPermission` element</span></span>

<span data-ttu-id="fcda5-104">アドインが関連する API または機能にアクセスするために必要な拡張アクセス許可を定義します。</span><span class="sxs-lookup"><span data-stu-id="fcda5-104">Defines an extended permission the add-in needs to access the associated API or feature.</span></span> <span data-ttu-id="fcda5-105">`ExtendedPermission`要素は、 [extendedpermissions](extendedpermissions.md)の子要素です。</span><span class="sxs-lookup"><span data-stu-id="fcda5-105">The `ExtendedPermission` element is a child element of [ExtendedPermissions](extendedpermissions.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="fcda5-106">この要素のサポートは、要件セット1.9 で導入されました。</span><span class="sxs-lookup"><span data-stu-id="fcda5-106">Support for this element was introduced in requirement set 1.9.</span></span> <span data-ttu-id="fcda5-107">この要件セットをサポートする [クライアントおよびプラットフォーム](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) を参照してください。</span><span class="sxs-lookup"><span data-stu-id="fcda5-107">See [clients and platforms](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.</span></span>

## <a name="available-extended-permissions"></a><span data-ttu-id="fcda5-108">利用可能な拡張アクセス許可</span><span class="sxs-lookup"><span data-stu-id="fcda5-108">Available extended permissions</span></span>

<span data-ttu-id="fcda5-109">使用可能な値は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="fcda5-109">The following are the available values.</span></span>

|<span data-ttu-id="fcda5-110">利用可能な値</span><span class="sxs-lookup"><span data-stu-id="fcda5-110">Available value</span></span>|<span data-ttu-id="fcda5-111">説明</span><span class="sxs-lookup"><span data-stu-id="fcda5-111">Description</span></span>|<span data-ttu-id="fcda5-112">Hosts</span><span class="sxs-lookup"><span data-stu-id="fcda5-112">Hosts</span></span>|
|---|---|---|
|`AppendOnSend`|<span data-ttu-id="fcda5-113">アドインが [Office. appendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#appendonsendasync-data--options--callback-) API を使用していることを宣言します。</span><span class="sxs-lookup"><span data-stu-id="fcda5-113">Declares that the add-in is using the [Office.Body.appendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#appendonsendasync-data--options--callback-) API.</span></span>|<span data-ttu-id="fcda5-114">Outlook</span><span class="sxs-lookup"><span data-stu-id="fcda5-114">Outlook</span></span>|

## <a name="extendedpermission-example"></a><span data-ttu-id="fcda5-115">`ExtendedPermission` 例</span><span class="sxs-lookup"><span data-stu-id="fcda5-115">`ExtendedPermission` example</span></span>

<span data-ttu-id="fcda5-116">要素の例を次に示し `ExtendedPermission` ます。</span><span class="sxs-lookup"><span data-stu-id="fcda5-116">The following is an example of the `ExtendedPermission` element.</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="fcda5-117">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="fcda5-117">Contained in</span></span>

[<span data-ttu-id="fcda5-118">ExtendedPermissions</span><span class="sxs-lookup"><span data-stu-id="fcda5-118">ExtendedPermissions</span></span>](extendedpermissions.md)
