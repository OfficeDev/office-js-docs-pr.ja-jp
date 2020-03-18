---
title: マニフェスト ファイルの AppDomains 要素
description: Office アドインがページの読み込みに使用する`SourceLocation`要素に指定されているドメインに加えて、すべてのドメインを一覧表示します。
ms.date: 07/03/2019
localization_priority: Normal
ms.openlocfilehash: f60579d773e81a7e8006bafcf1c151874af42aeb
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720702"
---
# <a name="appdomains-element"></a><span data-ttu-id="d8157-103">AppDomains 要素</span><span class="sxs-lookup"><span data-stu-id="d8157-103">AppDomains element</span></span>

<span data-ttu-id="d8157-104">Office アドインがページの読み込みに使用する`SourceLocation`要素に指定されているドメインに加えて、すべてのドメインを一覧表示します。</span><span class="sxs-lookup"><span data-stu-id="d8157-104">Lists any domains in addition to the domain specified in the `SourceLocation` element that your Office Add-in will use to load pages.</span></span> <span data-ttu-id="d8157-105">また、アドイン内の Iframe から Office .js API 呼び出しを行うことができる信頼されたドメインも一覧表示されます。</span><span class="sxs-lookup"><span data-stu-id="d8157-105">It also lists trusted domains from which Office.js API calls can be made from IFrames within the add-in.</span></span> <span data-ttu-id="d8157-106">追加の各ドメインに、AppDomain 要素を指定します。</span><span class="sxs-lookup"><span data-stu-id="d8157-106">For each additional domain, specify an AppDomain element.</span></span>

 <span data-ttu-id="d8157-107">**アドインの種類:** コンテンツ、作業ウィンドウ、メール</span><span class="sxs-lookup"><span data-stu-id="d8157-107">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="d8157-108">構文</span><span class="sxs-lookup"><span data-stu-id="d8157-108">Syntax</span></span>

```XML
<AppDomains>
    <AppDomain>AppDomain1</AppDomain>
    <AppDomain>AppDomain2</AppDomain>
</AppDomains>
```

> [!IMPORTANT]
> <span data-ttu-id="d8157-109">すべての **AppDomain** 要素の値には、プロトコル (例: `<AppDomain>https://myappdomain<AppDomain>`) が含まれている必要があります。</span><span class="sxs-lookup"><span data-stu-id="d8157-109">The value of each **AppDomain** element must include the protocol (e.g., `<AppDomain>https://myappdomain<AppDomain>`).</span></span>

## <a name="contained-in"></a><span data-ttu-id="d8157-110">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="d8157-110">Contained in</span></span>

[<span data-ttu-id="d8157-111">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="d8157-111">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="d8157-112">含めることができるもの</span><span class="sxs-lookup"><span data-stu-id="d8157-112">Can contain</span></span>

[<span data-ttu-id="d8157-113">AppDomain</span><span class="sxs-lookup"><span data-stu-id="d8157-113">AppDomain</span></span>](appdomain.md)

## <a name="remarks"></a><span data-ttu-id="d8157-114">解説</span><span class="sxs-lookup"><span data-stu-id="d8157-114">Remarks</span></span>

<span data-ttu-id="d8157-115">アドインは、既定では [SourceLocation](sourcelocation.md) 要素で指定されたものと同じ場所のドメインのページを読み込みます。</span><span class="sxs-lookup"><span data-stu-id="d8157-115">By default, your add-in can load any page that is in the same domain as the location specified in the [SourceLocation element](sourcelocation.md).</span></span> <span data-ttu-id="d8157-116">アドインと同じドメインにないページを読み込む場合は、**AppDomains** 要素と **AppDomain** 要素を使用してドメインを指定します。</span><span class="sxs-lookup"><span data-stu-id="d8157-116">To load pages that are not in the same domain as the add-in, specify their domains by using the **AppDomains** and **AppDomain** elements.</span></span> <span data-ttu-id="d8157-117">この要素は空にできません。</span><span class="sxs-lookup"><span data-stu-id="d8157-117">This element can't be empty.</span></span>
