---
title: マニフェスト ファイルの AppDomain 要素
description: ''
ms.date: 05/15/2019
localization_priority: Normal
ms.openlocfilehash: b1d71648cc7646eec246f3d0a8113c843eed2e74
ms.sourcegitcommit: b0e71ae0ae09c57b843d4de277081845c108a645
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/21/2019
ms.locfileid: "34337196"
---
# <a name="appdomain-element"></a><span data-ttu-id="f01b5-102">AppDomain 要素</span><span class="sxs-lookup"><span data-stu-id="f01b5-102">AppDomain element</span></span>

<span data-ttu-id="f01b5-103">アドイン ウィンドウにページを読み込むために使用される追加のドメインを指定します。</span><span class="sxs-lookup"><span data-stu-id="f01b5-103">Specifies an additional domain that will be used to load pages in the add-in window.</span></span>

<span data-ttu-id="f01b5-104">**アドインの種類:** コンテンツ、作業ウィンドウ、メール</span><span class="sxs-lookup"><span data-stu-id="f01b5-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="f01b5-105">構文</span><span class="sxs-lookup"><span data-stu-id="f01b5-105">Syntax</span></span>

```XML
<AppDomain>string</AppDomain>
```

> [!IMPORTANT]
> 1. <span data-ttu-id="f01b5-106">**AppDomain** 要素の値には、プロトコル (例: `<AppDomain>https://myappdomain</AppDomain>`) が含まれている必要があります。</span><span class="sxs-lookup"><span data-stu-id="f01b5-106">The value of the **AppDomain** element must include the protocol (e.g., `<AppDomain>https://myappdomain</AppDomain>`).</span></span>
> 2. <span data-ttu-id="f01b5-107">値には、末尾にスラッシュ "/" を付け*ない*でください。</span><span class="sxs-lookup"><span data-stu-id="f01b5-107">Do *not* put a closing slash, "/", on the value.</span></span>

## <a name="contained-in"></a><span data-ttu-id="f01b5-108">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="f01b5-108">Contained in</span></span>

[<span data-ttu-id="f01b5-109">AppDomains</span><span class="sxs-lookup"><span data-stu-id="f01b5-109">AppDomains</span></span>](appdomains.md)

## <a name="remarks"></a><span data-ttu-id="f01b5-110">解説</span><span class="sxs-lookup"><span data-stu-id="f01b5-110">Remarks</span></span>

<span data-ttu-id="f01b5-111">**AppDomain** 要素は、[SourceLocation](sourcelocation.md) 要素で指定したドメイン以外のものを追加指定するために使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="f01b5-111">**AppDomain** elements should be used to specify any additional domains other than the one specified in the [SourceLocation element](sourcelocation.md).</span></span> <span data-ttu-id="f01b5-112">詳細については、「[Office アドイン XML マニフェスト](/office/dev/add-ins/develop/add-in-manifests)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="f01b5-112">For more information, see [Office Add-ins XML manifest](/office/dev/add-ins/develop/add-in-manifests).</span></span>
