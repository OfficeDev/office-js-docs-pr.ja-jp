---
title: マニフェスト ファイルの AppDomains 要素
description: ''
ms.date: 12/13/2018
localization_priority: Normal
ms.openlocfilehash: 65391c9529e7ddaa9726d0b58accf90c5b9babef
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450647"
---
# <a name="appdomains-element"></a><span data-ttu-id="4ad58-102">AppDomains 要素</span><span class="sxs-lookup"><span data-stu-id="4ad58-102">AppDomains element</span></span>

<span data-ttu-id="4ad58-p101">Office アドイン でページを読み込むのに使う SourceLocation 要素で指定されたドメインの他に、任意のドメインを一覧表示します。追加の各ドメインに、AppDomain 要素を指定します。</span><span class="sxs-lookup"><span data-stu-id="4ad58-p101">Lists any domains in addition to the domain specified in the SourceLocation element that your Office Add-in will use to load pages. For each additional domain, specify an AppDomain element.</span></span>

 <span data-ttu-id="4ad58-105">**アドインの種類:** コンテンツ、作業ウィンドウ、メール</span><span class="sxs-lookup"><span data-stu-id="4ad58-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="4ad58-106">構文</span><span class="sxs-lookup"><span data-stu-id="4ad58-106">Syntax</span></span>

```XML
<AppDomains>
    <AppDomain>AppDomain1</AppDomain>
    <AppDomain>AppDomain2</AppDomain>
</AppDomains>
```

> [!IMPORTANT]
> <span data-ttu-id="4ad58-107">すべての **AppDomain** 要素の値には、プロトコル (例: `<AppDomain>https://myappdomain<AppDomain>`) が含まれている必要があります。</span><span class="sxs-lookup"><span data-stu-id="4ad58-107">The value of each **AppDomain** element must include the protocol (e.g., `<AppDomain>https://myappdomain<AppDomain>`).</span></span>

## <a name="contained-in"></a><span data-ttu-id="4ad58-108">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="4ad58-108">Contained in</span></span>

[<span data-ttu-id="4ad58-109">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="4ad58-109">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="4ad58-110">含めることができるもの</span><span class="sxs-lookup"><span data-stu-id="4ad58-110">Can contain</span></span>

[<span data-ttu-id="4ad58-111">AppDomain</span><span class="sxs-lookup"><span data-stu-id="4ad58-111">AppDomain</span></span>](appdomain.md)

## <a name="remarks"></a><span data-ttu-id="4ad58-112">解説</span><span class="sxs-lookup"><span data-stu-id="4ad58-112">Remarks</span></span>

<span data-ttu-id="4ad58-113">アドインは、既定では [SourceLocation](sourcelocation.md) 要素で指定されたものと同じ場所のドメインのページを読み込みます。</span><span class="sxs-lookup"><span data-stu-id="4ad58-113">By default, your add-in can load any page that is in the same domain as the location specified in the [SourceLocation element](sourcelocation.md).</span></span> <span data-ttu-id="4ad58-114">アドインと同じドメインにないページを読み込む場合は、**AppDomains** 要素と **AppDomain** 要素を使用してドメインを指定します。</span><span class="sxs-lookup"><span data-stu-id="4ad58-114">To load pages that are not in the same domain as the add-in, specify their domains by using the **AppDomains** and **AppDomain** elements.</span></span> <span data-ttu-id="4ad58-115">この要素は空にできません。</span><span class="sxs-lookup"><span data-stu-id="4ad58-115">This element can't be empty.</span></span>
