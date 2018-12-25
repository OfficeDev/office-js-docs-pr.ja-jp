---
title: マニフェスト ファイルの AppDomains 要素
description: ''
ms.date: 12/13/2018
ms.openlocfilehash: cc2f5ade0bdda214c85490f8e474b42f921edbe8
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433680"
---
# <a name="appdomains-element"></a><span data-ttu-id="24e8d-102">AppDomains 要素</span><span class="sxs-lookup"><span data-stu-id="24e8d-102">AppDomains element</span></span>

<span data-ttu-id="24e8d-p101">Office アドイン でページを読み込むのに使う SourceLocation 要素で指定されたドメインの他に、任意のドメインを一覧表示します。追加の各ドメインに、AppDomain 要素を指定します。</span><span class="sxs-lookup"><span data-stu-id="24e8d-p101">Lists any domains in addition to the domain specified in the SourceLocation element that your Office Add-in will use to load pages. For each additional domain, specify an AppDomain element.</span></span>

 <span data-ttu-id="24e8d-105">**アドインの種類:** コンテンツ、作業ウィンドウ、メール</span><span class="sxs-lookup"><span data-stu-id="24e8d-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="24e8d-106">構文</span><span class="sxs-lookup"><span data-stu-id="24e8d-106">Syntax</span></span>

```XML
<AppDomains>
    <AppDomain>AppDomain1</AppDomain>
    <AppDomain>AppDomain2</AppDomain>
</AppDomains>
```

> [!IMPORTANT]
> <span data-ttu-id="24e8d-107">すべての **AppDomain** 要素の値には、プロトコル (例: `<AppDomain>https://myappdomain<AppDomain>`) が含まれている必要があります。</span><span class="sxs-lookup"><span data-stu-id="24e8d-107">The value of each **AppDomain** element must include the protocol (e.g., `<AppDomain>https://myappdomain<AppDomain>`).</span></span>

## <a name="contained-in"></a><span data-ttu-id="24e8d-108">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="24e8d-108">Contained in</span></span>

[<span data-ttu-id="24e8d-109">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="24e8d-109">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="24e8d-110">含めることができるもの</span><span class="sxs-lookup"><span data-stu-id="24e8d-110">Can contain</span></span>

[<span data-ttu-id="24e8d-111">AppDomain</span><span class="sxs-lookup"><span data-stu-id="24e8d-111">AppDomain</span></span>](appdomain.md)

## <a name="remarks"></a><span data-ttu-id="24e8d-112">解説</span><span class="sxs-lookup"><span data-stu-id="24e8d-112">Remarks</span></span>

<span data-ttu-id="24e8d-113">アドインは、既定では [SourceLocation](sourcelocation.md) 要素で指定されたものと同じ場所のドメインのページを読み込みます。</span><span class="sxs-lookup"><span data-stu-id="24e8d-113">By default, your add-in can load any page that is in the same domain as the location specified in the [SourceLocation element](sourcelocation.md).</span></span> <span data-ttu-id="24e8d-114">アドインと同じドメインにないページを読み込む場合は、**AppDomains** 要素と **AppDomain** 要素を使用してドメインを指定します。</span><span class="sxs-lookup"><span data-stu-id="24e8d-114">To load pages that are not in the same domain as the add-in, specify their domains by using the **AppDomains** and **AppDomain** elements.</span></span> <span data-ttu-id="24e8d-115">この要素は空にできません。</span><span class="sxs-lookup"><span data-stu-id="24e8d-115">This element can't be empty.</span></span>
