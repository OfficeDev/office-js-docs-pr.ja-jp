---
title: マニフェスト ファイルの AppDomains 要素
description: Office アドインがページの読み込みに使用する要素に指定されているドメインに加えて、すべてのドメインを一覧表示 `SourceLocation` します。
ms.date: 07/03/2019
localization_priority: Normal
ms.openlocfilehash: 9183f1815e97bd8d4ac1a7e2cf72d5547d153f7e
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608769"
---
# <a name="appdomains-element"></a><span data-ttu-id="1f0d4-103">AppDomains 要素</span><span class="sxs-lookup"><span data-stu-id="1f0d4-103">AppDomains element</span></span>

<span data-ttu-id="1f0d4-104">Office アドインがページの読み込みに使用する要素に指定されているドメインに加えて、すべてのドメインを一覧表示 `SourceLocation` します。</span><span class="sxs-lookup"><span data-stu-id="1f0d4-104">Lists any domains in addition to the domain specified in the `SourceLocation` element that your Office Add-in will use to load pages.</span></span> <span data-ttu-id="1f0d4-105">また、アドイン内の Iframe から Office .js API 呼び出しを行うことができる信頼されたドメインも一覧表示されます。</span><span class="sxs-lookup"><span data-stu-id="1f0d4-105">It also lists trusted domains from which Office.js API calls can be made from IFrames within the add-in.</span></span> <span data-ttu-id="1f0d4-106">追加の各ドメインに、AppDomain 要素を指定します。</span><span class="sxs-lookup"><span data-stu-id="1f0d4-106">For each additional domain, specify an AppDomain element.</span></span>

 <span data-ttu-id="1f0d4-107">**アドインの種類:** コンテンツ、作業ウィンドウ、メール</span><span class="sxs-lookup"><span data-stu-id="1f0d4-107">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="1f0d4-108">構文</span><span class="sxs-lookup"><span data-stu-id="1f0d4-108">Syntax</span></span>

```XML
<AppDomains>
    <AppDomain>AppDomain1</AppDomain>
    <AppDomain>AppDomain2</AppDomain>
</AppDomains>
```

> [!IMPORTANT]
> <span data-ttu-id="1f0d4-109">すべての **AppDomain** 要素の値には、プロトコル (例: `<AppDomain>https://myappdomain<AppDomain>`) が含まれている必要があります。</span><span class="sxs-lookup"><span data-stu-id="1f0d4-109">The value of each **AppDomain** element must include the protocol (e.g., `<AppDomain>https://myappdomain<AppDomain>`).</span></span>

## <a name="contained-in"></a><span data-ttu-id="1f0d4-110">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="1f0d4-110">Contained in</span></span>

[<span data-ttu-id="1f0d4-111">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="1f0d4-111">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="1f0d4-112">含めることができるもの</span><span class="sxs-lookup"><span data-stu-id="1f0d4-112">Can contain</span></span>

[<span data-ttu-id="1f0d4-113">AppDomain</span><span class="sxs-lookup"><span data-stu-id="1f0d4-113">AppDomain</span></span>](appdomain.md)

## <a name="remarks"></a><span data-ttu-id="1f0d4-114">解説</span><span class="sxs-lookup"><span data-stu-id="1f0d4-114">Remarks</span></span>

<span data-ttu-id="1f0d4-115">アドインは、既定では [SourceLocation](sourcelocation.md) 要素で指定されたものと同じ場所のドメインのページを読み込みます。</span><span class="sxs-lookup"><span data-stu-id="1f0d4-115">By default, your add-in can load any page that is in the same domain as the location specified in the [SourceLocation element](sourcelocation.md).</span></span> <span data-ttu-id="1f0d4-116">アドインと同じドメインにないページを読み込む場合は、**AppDomains** 要素と **AppDomain** 要素を使用してドメインを指定します。</span><span class="sxs-lookup"><span data-stu-id="1f0d4-116">To load pages that are not in the same domain as the add-in, specify their domains by using the **AppDomains** and **AppDomain** elements.</span></span> <span data-ttu-id="1f0d4-117">この要素は空にできません。</span><span class="sxs-lookup"><span data-stu-id="1f0d4-117">This element can't be empty.</span></span>
