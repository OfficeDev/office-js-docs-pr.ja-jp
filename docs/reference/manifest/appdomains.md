---
title: マニフェスト ファイルの AppDomains 要素
description: '`SourceLocation`Office アドインが使用する、office によって信頼される必要がある、要素で指定されているドメインに加えて、すべてのドメインを一覧表示します。'
ms.date: 06/12/2020
localization_priority: Normal
ms.openlocfilehash: 751e4ad2ffa5fd50739a855fad48964473b154f1
ms.sourcegitcommit: 9eed5201a3ef556f77ba3b6790f007358188d57d
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/17/2020
ms.locfileid: "44778656"
---
# <a name="appdomains-element"></a><span data-ttu-id="f8276-103">AppDomains 要素</span><span class="sxs-lookup"><span data-stu-id="f8276-103">AppDomains element</span></span>

<span data-ttu-id="f8276-104">Office `SourceLocation` アドインが使用し、office によって信頼されるようにする必要がある、要素で指定されているドメインに加えて、すべてのドメインを一覧表示します。</span><span class="sxs-lookup"><span data-stu-id="f8276-104">Lists any domains, in addition to the domain specified in the `SourceLocation` element, that your Office Add-in will use and that should be trusted by Office.</span></span> <span data-ttu-id="f8276-105">これにより、ドメイン内のページは、アドイン内の Iframe から Office.js Api を呼び出すことができるようになり、他の効果があります。</span><span class="sxs-lookup"><span data-stu-id="f8276-105">This enables pages in the domains to make calls to Office.js APIs from IFrames within the add-in and has other effects.</span></span> <span data-ttu-id="f8276-106">追加の各ドメインに、**AppDomain** 要素を指定します。</span><span class="sxs-lookup"><span data-stu-id="f8276-106">For each additional domain, specify an **AppDomain** element.</span></span>

 <span data-ttu-id="f8276-107">**アドインの種類:** コンテンツ、作業ウィンドウ、メール</span><span class="sxs-lookup"><span data-stu-id="f8276-107">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="f8276-108">構文</span><span class="sxs-lookup"><span data-stu-id="f8276-108">Syntax</span></span>

```XML
<AppDomains>
    <AppDomain>AppDomain1</AppDomain>
    <AppDomain>AppDomain2</AppDomain>
</AppDomains>
```

> [!IMPORTANT]
> <span data-ttu-id="f8276-109">**AppDomain**要素の値には、いくつかの制限があります。</span><span class="sxs-lookup"><span data-stu-id="f8276-109">There are restrictions on what can be the value of a **AppDomain** element.</span></span> <span data-ttu-id="f8276-110">詳細については、「 [AppDomain](appdomain.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="f8276-110">For more information, see [AppDomain](appdomain.md).</span></span>

## <a name="contained-in"></a><span data-ttu-id="f8276-111">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="f8276-111">Contained in</span></span>

[<span data-ttu-id="f8276-112">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="f8276-112">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="f8276-113">含めることができるもの</span><span class="sxs-lookup"><span data-stu-id="f8276-113">Can contain</span></span>

[<span data-ttu-id="f8276-114">AppDomain</span><span class="sxs-lookup"><span data-stu-id="f8276-114">AppDomain</span></span>](appdomain.md)

## <a name="remarks"></a><span data-ttu-id="f8276-115">解説</span><span class="sxs-lookup"><span data-stu-id="f8276-115">Remarks</span></span>

<span data-ttu-id="f8276-116">アドインは、既定では [SourceLocation](sourcelocation.md) 要素で指定されたものと同じ場所のドメインのページを読み込みます。</span><span class="sxs-lookup"><span data-stu-id="f8276-116">By default, your add-in can load any page that is in the same domain as the location specified in the [SourceLocation element](sourcelocation.md).</span></span> <span data-ttu-id="f8276-117">この要素は空にできません。</span><span class="sxs-lookup"><span data-stu-id="f8276-117">This element can't be empty.</span></span>
