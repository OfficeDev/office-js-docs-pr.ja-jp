---
title: マニフェストファイルの ExtendedPermission 要素
description: アドインが関連する API または機能にアクセスするために必要な拡張アクセス許可を定義します。
ms.date: 03/05/2020
localization_priority: Normal
ms.openlocfilehash: ca4c2d12cb9a5be159c22712b631d0bde42e48ed
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611541"
---
# <a name="extendedpermission-element"></a><span data-ttu-id="60f26-103">`ExtendedPermission`項目</span><span class="sxs-lookup"><span data-stu-id="60f26-103">`ExtendedPermission` element</span></span>

<span data-ttu-id="60f26-104">アドインが関連する API または機能にアクセスするために必要な拡張アクセス許可を定義します。</span><span class="sxs-lookup"><span data-stu-id="60f26-104">Defines an extended permission the add-in needs to access the associated API or feature.</span></span> <span data-ttu-id="60f26-105">`ExtendedPermission`要素は、 [extendedpermissions](extendedpermissions.md)の子要素です。</span><span class="sxs-lookup"><span data-stu-id="60f26-105">The `ExtendedPermission` element is a child element of [ExtendedPermissions](extendedpermissions.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="60f26-106">この要素は、Exchange Online に対して[Outlook アドインのプレビュー要件が設定](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md)されている場合にのみ使用できます。</span><span class="sxs-lookup"><span data-stu-id="60f26-106">This element is only available in the [Outlook add-ins preview requirement set](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) against Exchange Online.</span></span> <span data-ttu-id="60f26-107">この要素を使用するアドインは、AppSource に発行したり、一元展開によって展開したりすることはできません。</span><span class="sxs-lookup"><span data-stu-id="60f26-107">Add-ins that use this element cannot be published to AppSource or deployed via centralized deployment.</span></span>

## <a name="available-extended-permissions"></a><span data-ttu-id="60f26-108">利用可能な拡張アクセス許可</span><span class="sxs-lookup"><span data-stu-id="60f26-108">Available extended permissions</span></span>

<span data-ttu-id="60f26-109">使用可能な値は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="60f26-109">The following are the available values.</span></span>

|<span data-ttu-id="60f26-110">利用可能な値</span><span class="sxs-lookup"><span data-stu-id="60f26-110">Available value</span></span>|<span data-ttu-id="60f26-111">Description</span><span class="sxs-lookup"><span data-stu-id="60f26-111">Description</span></span>|<span data-ttu-id="60f26-112">Hosts</span><span class="sxs-lookup"><span data-stu-id="60f26-112">Hosts</span></span>|
|---|---|---|
|`AppendOnSend`|<span data-ttu-id="60f26-113">アドインが[Office. appendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-preview#appendonsendasync-data--options--callback-) API を使用していることを宣言します。</span><span class="sxs-lookup"><span data-stu-id="60f26-113">Declares that the add-in is using the [Office.Body.appendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-preview#appendonsendasync-data--options--callback-) API.</span></span>|<span data-ttu-id="60f26-114">Outlook</span><span class="sxs-lookup"><span data-stu-id="60f26-114">Outlook</span></span>|

## <a name="extendedpermission-example"></a><span data-ttu-id="60f26-115">`ExtendedPermission`例</span><span class="sxs-lookup"><span data-stu-id="60f26-115">`ExtendedPermission` example</span></span>

<span data-ttu-id="60f26-116">要素の例を次に示し `ExtendedPermission` ます。</span><span class="sxs-lookup"><span data-stu-id="60f26-116">The following is an example of the `ExtendedPermission` element.</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="60f26-117">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="60f26-117">Contained in</span></span>

[<span data-ttu-id="60f26-118">ExtendedPermissions</span><span class="sxs-lookup"><span data-stu-id="60f26-118">ExtendedPermissions</span></span>](extendedpermissions.md)
