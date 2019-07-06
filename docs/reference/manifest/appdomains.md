---
title: マニフェスト ファイルの AppDomains 要素
description: ''
ms.date: 07/03/2019
localization_priority: Normal
ms.openlocfilehash: b6db3d46d004021f25edd5733566544010abb457
ms.sourcegitcommit: c3673cc693fa7070e1b397922bd735ba3f9342f3
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/05/2019
ms.locfileid: "35575332"
---
# <a name="appdomains-element"></a><span data-ttu-id="e5f44-102">AppDomains 要素</span><span class="sxs-lookup"><span data-stu-id="e5f44-102">AppDomains element</span></span>

<span data-ttu-id="e5f44-103">Office アドインがページの読み込みに使用する`SourceLocation`要素に指定されているドメインに加えて、すべてのドメインを一覧表示します。</span><span class="sxs-lookup"><span data-stu-id="e5f44-103">Lists any domains in addition to the domain specified in the `SourceLocation` element that your Office Add-in will use to load pages.</span></span> <span data-ttu-id="e5f44-104">また、アドイン内の Iframe から Office .js API 呼び出しを行うことができる信頼されたドメインも一覧表示されます。</span><span class="sxs-lookup"><span data-stu-id="e5f44-104">It also lists trusted domains from which Office.js API calls can be made from IFrames within the add-in.</span></span> <span data-ttu-id="e5f44-105">追加の各ドメインに、AppDomain 要素を指定します。</span><span class="sxs-lookup"><span data-stu-id="e5f44-105">For each additional domain, specify an AppDomain element.</span></span>

 <span data-ttu-id="e5f44-106">**アドインの種類:** コンテンツ、作業ウィンドウ、メール</span><span class="sxs-lookup"><span data-stu-id="e5f44-106">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="e5f44-107">構文</span><span class="sxs-lookup"><span data-stu-id="e5f44-107">Syntax</span></span>

```XML
<AppDomains>
    <AppDomain>AppDomain1</AppDomain>
    <AppDomain>AppDomain2</AppDomain>
</AppDomains>
```

> [!IMPORTANT]
> <span data-ttu-id="e5f44-108">すべての **AppDomain** 要素の値には、プロトコル (例: `<AppDomain>https://myappdomain<AppDomain>`) が含まれている必要があります。</span><span class="sxs-lookup"><span data-stu-id="e5f44-108">The value of each **AppDomain** element must include the protocol (e.g., `<AppDomain>https://myappdomain<AppDomain>`).</span></span>

## <a name="contained-in"></a><span data-ttu-id="e5f44-109">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="e5f44-109">Contained in</span></span>

[<span data-ttu-id="e5f44-110">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="e5f44-110">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="e5f44-111">含めることができるもの</span><span class="sxs-lookup"><span data-stu-id="e5f44-111">Can contain</span></span>

[<span data-ttu-id="e5f44-112">AppDomain</span><span class="sxs-lookup"><span data-stu-id="e5f44-112">AppDomain</span></span>](appdomain.md)

## <a name="remarks"></a><span data-ttu-id="e5f44-113">解説</span><span class="sxs-lookup"><span data-stu-id="e5f44-113">Remarks</span></span>

<span data-ttu-id="e5f44-114">アドインは、既定では [SourceLocation](sourcelocation.md) 要素で指定されたものと同じ場所のドメインのページを読み込みます。</span><span class="sxs-lookup"><span data-stu-id="e5f44-114">By default, your add-in can load any page that is in the same domain as the location specified in the [SourceLocation element](sourcelocation.md).</span></span> <span data-ttu-id="e5f44-115">アドインと同じドメインにないページを読み込む場合は、**AppDomains** 要素と **AppDomain** 要素を使用してドメインを指定します。</span><span class="sxs-lookup"><span data-stu-id="e5f44-115">To load pages that are not in the same domain as the add-in, specify their domains by using the **AppDomains** and **AppDomain** elements.</span></span> <span data-ttu-id="e5f44-116">この要素は空にできません。</span><span class="sxs-lookup"><span data-stu-id="e5f44-116">This element can't be empty.</span></span>
