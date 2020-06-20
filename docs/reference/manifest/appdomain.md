---
title: マニフェスト ファイルの AppDomain 要素
description: アドインで使用される追加のドメインを指定します。 Office によって信頼される必要があります。
ms.date: 06/12/2020
localization_priority: Normal
ms.openlocfilehash: ae49944afceada559b39353cd119e26a21fd3d15
ms.sourcegitcommit: 9eed5201a3ef556f77ba3b6790f007358188d57d
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/17/2020
ms.locfileid: "44778649"
---
# <a name="appdomain-element"></a><span data-ttu-id="e12c6-103">AppDomain 要素</span><span class="sxs-lookup"><span data-stu-id="e12c6-103">AppDomain element</span></span>

<span data-ttu-id="e12c6-104">[SourceLocation 要素](sourcelocation.md)で指定されているものに加えて、Office が信頼する必要がある追加のドメインを指定します。</span><span class="sxs-lookup"><span data-stu-id="e12c6-104">Specifies an additional domain that Office should trust, in addition to the one specified in the [SourceLocation element](sourcelocation.md).</span></span> <span data-ttu-id="e12c6-105">ドメインの指定には、次のような影響があります。</span><span class="sxs-lookup"><span data-stu-id="e12c6-105">Specifying a domain has these effects:</span></span>

- <span data-ttu-id="e12c6-106">これにより、ドメイン内のページ、ルート、またはその他のリソースを、デスクトップの Office プラットフォーム上のアドインのルート作業ウィンドウで直接開くことができます。</span><span class="sxs-lookup"><span data-stu-id="e12c6-106">It enables pages, routes, or other resources in the domain to be opened directly in the root task pane of the add-in on desktop Office platforms.</span></span> <span data-ttu-id="e12c6-107">(Web 上の Office、または IFrame でリソースを開くために**AppDomain**でドメインを指定する必要はありません)。または、[ダイアログ API](../../develop/dialog-api-in-office-add-ins.md)で開いたダイアログでリソースを開く必要はありません。</span><span class="sxs-lookup"><span data-stu-id="e12c6-107">(Specifying a domain in an **AppDomain** isn't necessary for Office on the web or to open a resource in an IFrame, nor it is necessary for opening a resource in a dialog opened with the [Dialog API](../../develop/dialog-api-in-office-add-ins.md).)</span></span>
- <span data-ttu-id="e12c6-108">これにより、ドメイン内のページは、アドイン内の Iframe から Office.js API 呼び出しを実行できるようになります。</span><span class="sxs-lookup"><span data-stu-id="e12c6-108">It enables pages in the domain to make Office.js API calls from IFrames within the add-in.</span></span>

<span data-ttu-id="e12c6-109">**アドインの種類:** コンテンツ、作業ウィンドウ、メール</span><span class="sxs-lookup"><span data-stu-id="e12c6-109">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="e12c6-110">構文</span><span class="sxs-lookup"><span data-stu-id="e12c6-110">Syntax</span></span>

```XML
<AppDomain>string</AppDomain>
```

> [!IMPORTANT]
> 1. <span data-ttu-id="e12c6-111">**AppDomain** 要素の値には、プロトコル (例: `<AppDomain>https://myappdomain.com</AppDomain>`) が含まれている必要があります。</span><span class="sxs-lookup"><span data-stu-id="e12c6-111">The value of the **AppDomain** element must include the protocol (e.g., `<AppDomain>https://myappdomain.com</AppDomain>`).</span></span>
> 2. <span data-ttu-id="e12c6-112">ドメインの明示的なポートがある場合は、そのポートを含めます (例: `<AppDomain>https://myappdomain.com:9999</AppDomain>` )。</span><span class="sxs-lookup"><span data-stu-id="e12c6-112">If there is an explicit port for the domain, include it (e.g.,`<AppDomain>https://myappdomain.com:9999</AppDomain>`).</span></span>
> 3. <span data-ttu-id="e12c6-113">サブドメインを信頼する必要がある場合は、それを含めます (例: `<AppDomain>https://mysubdomain.myappdomain.com</AppDomain>` )。</span><span class="sxs-lookup"><span data-stu-id="e12c6-113">If a subdomain needs to be trusted, include it (e.g.,`<AppDomain>https://mysubdomain.myappdomain.com</AppDomain>`).</span></span> <span data-ttu-id="e12c6-114">サブドメイン `mysubdomain.mydomain.com` と `mydomain.com` ドメインが異なる。</span><span class="sxs-lookup"><span data-stu-id="e12c6-114">The subdomain `mysubdomain.mydomain.com` and `mydomain.com` are different domains.</span></span> <span data-ttu-id="e12c6-115">両方を信頼する必要がある場合は、どちらも別個の**AppDomain**要素にする必要があります。</span><span class="sxs-lookup"><span data-stu-id="e12c6-115">If both need to be trusted, then both need to be in separate **AppDomain** elements.</span></span>
> 4. <span data-ttu-id="e12c6-116">[SourceLocation 要素](sourcelocation.md)で指定されたドメインと同じドメインを一覧表示することはできません。誤解を招く可能性があります。</span><span class="sxs-lookup"><span data-stu-id="e12c6-116">Listing the same domain as the one specified in the [SourceLocation element](sourcelocation.md) has no effect and may be misleading.</span></span> <span data-ttu-id="e12c6-117">特に、を開発する場合は、 `localhost` 用の**AppDomain**要素を作成する必要はありません `localhost` 。</span><span class="sxs-lookup"><span data-stu-id="e12c6-117">In particular, when you are developing on `localhost`, you don't need to create an **AppDomain** element for `localhost`.</span></span>
> 5. <span data-ttu-id="e12c6-118">ドメインを超える URL のセグメントは含めないでください。</span><span class="sxs-lookup"><span data-stu-id="e12c6-118">Don't include any segments of a URL past the domain.</span></span> <span data-ttu-id="e12c6-119">たとえば、ページの完全な URL を含めないでください。</span><span class="sxs-lookup"><span data-stu-id="e12c6-119">For example, don't include the full URL of a page.</span></span>
> 6. <span data-ttu-id="e12c6-120">値には、末尾にスラッシュ "/" を付け*ない*でください。</span><span class="sxs-lookup"><span data-stu-id="e12c6-120">Do *not* put a closing slash, "/", on the value.</span></span>

## <a name="contained-in"></a><span data-ttu-id="e12c6-121">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="e12c6-121">Contained in</span></span>

[<span data-ttu-id="e12c6-122">AppDomains</span><span class="sxs-lookup"><span data-stu-id="e12c6-122">AppDomains</span></span>](appdomains.md)

## <a name="remarks"></a><span data-ttu-id="e12c6-123">注釈</span><span class="sxs-lookup"><span data-stu-id="e12c6-123">Remarks</span></span>

<span data-ttu-id="e12c6-124">詳細については、「[Office アドインの XML マニフェスト](../../develop/add-in-manifests.md)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="e12c6-124">For more information, see [Office Add-ins XML manifest](../../develop/add-in-manifests.md).</span></span>
