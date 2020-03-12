---
title: マニフェストファイルの ExtendedPermission 要素
description: ''
ms.date: 03/05/2020
localization_priority: Normal
ms.openlocfilehash: 6c41684fc922f5845559250311edd8182788cfc5
ms.sourcegitcommit: a0262ea40cd23f221e69bcb0223110f011265d13
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/12/2020
ms.locfileid: "42605810"
---
# <a name="extendedpermission-element"></a><span data-ttu-id="e5d4b-102">`ExtendedPermission`項目</span><span class="sxs-lookup"><span data-stu-id="e5d4b-102">`ExtendedPermission` element</span></span>

<span data-ttu-id="e5d4b-103">アドインが関連する API または機能にアクセスするために必要な拡張アクセス許可を定義します。</span><span class="sxs-lookup"><span data-stu-id="e5d4b-103">Defines an extended permission the add-in needs to access the associated API or feature.</span></span> <span data-ttu-id="e5d4b-104">要素`ExtendedPermission`は、 [extendedpermissions](extendedpermissions.md)の子要素です。</span><span class="sxs-lookup"><span data-stu-id="e5d4b-104">The `ExtendedPermission` element is a child element of [ExtendedPermissions](extendedpermissions.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="e5d4b-105">この要素は、Exchange Online に対して[Outlook アドインのプレビュー要件が設定](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md)されている場合にのみ使用できます。</span><span class="sxs-lookup"><span data-stu-id="e5d4b-105">This element is only available in the [Outlook add-ins preview requirement set](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) against Exchange Online.</span></span> <span data-ttu-id="e5d4b-106">この要素を使用するアドインは、AppSource に発行したり、一元展開によって展開したりすることはできません。</span><span class="sxs-lookup"><span data-stu-id="e5d4b-106">Add-ins that use this element cannot be published to AppSource or deployed via centralized deployment.</span></span>

## <a name="available-extended-permissions"></a><span data-ttu-id="e5d4b-107">利用可能な拡張アクセス許可</span><span class="sxs-lookup"><span data-stu-id="e5d4b-107">Available extended permissions</span></span>

<span data-ttu-id="e5d4b-108">使用可能な値は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="e5d4b-108">The following are the available values.</span></span>

|<span data-ttu-id="e5d4b-109">利用可能な値</span><span class="sxs-lookup"><span data-stu-id="e5d4b-109">Available value</span></span>|<span data-ttu-id="e5d4b-110">説明</span><span class="sxs-lookup"><span data-stu-id="e5d4b-110">Description</span></span>|<span data-ttu-id="e5d4b-111">Hosts</span><span class="sxs-lookup"><span data-stu-id="e5d4b-111">Hosts</span></span>|
|---|---|---|
|`AppendOnSend`|<span data-ttu-id="e5d4b-112">アドインが[Office. appendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-preview#appendonsendasync-data--options--callback-) API を使用していることを宣言します。</span><span class="sxs-lookup"><span data-stu-id="e5d4b-112">Declares that the add-in is using the [Office.Body.appendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-preview#appendonsendasync-data--options--callback-) API.</span></span>|<span data-ttu-id="e5d4b-113">Outlook</span><span class="sxs-lookup"><span data-stu-id="e5d4b-113">Outlook</span></span>|

## <a name="extendedpermission-example"></a><span data-ttu-id="e5d4b-114">`ExtendedPermission`例</span><span class="sxs-lookup"><span data-stu-id="e5d4b-114">`ExtendedPermission` example</span></span>

<span data-ttu-id="e5d4b-115">`ExtendedPermission`要素の例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="e5d4b-115">The following is an example of the `ExtendedPermission` element.</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="e5d4b-116">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="e5d4b-116">Contained in</span></span>

[<span data-ttu-id="e5d4b-117">ExtendedPermissions</span><span class="sxs-lookup"><span data-stu-id="e5d4b-117">ExtendedPermissions</span></span>](extendedpermissions.md)
