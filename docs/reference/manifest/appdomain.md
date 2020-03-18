---
title: マニフェスト ファイルの AppDomain 要素
description: アドインウィンドウにページを読み込む追加のドメインを指定します。
ms.date: 07/03/2019
localization_priority: Normal
ms.openlocfilehash: 6990f759df806f24b1d617c036bc1a452e6da38f
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718455"
---
# <a name="appdomain-element"></a><span data-ttu-id="5f342-103">AppDomain 要素</span><span class="sxs-lookup"><span data-stu-id="5f342-103">AppDomain element</span></span>

<span data-ttu-id="5f342-104">アドインウィンドウにページを読み込む追加のドメインを指定します。</span><span class="sxs-lookup"><span data-stu-id="5f342-104">Specifies additional domains that load pages in the add-in window.</span></span> <span data-ttu-id="5f342-105">また、アドイン内の Iframe から Office .js API 呼び出しを行うことができる信頼されたドメインも一覧表示されます。</span><span class="sxs-lookup"><span data-stu-id="5f342-105">It also lists trusted domains from which Office.js API calls can be made from IFrames within the add-in.</span></span>

<span data-ttu-id="5f342-106">**アドインの種類:** コンテンツ、作業ウィンドウ、メール</span><span class="sxs-lookup"><span data-stu-id="5f342-106">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="5f342-107">構文</span><span class="sxs-lookup"><span data-stu-id="5f342-107">Syntax</span></span>

```XML
<AppDomain>string</AppDomain>
```

> [!IMPORTANT]
> 1. <span data-ttu-id="5f342-108">**AppDomain** 要素の値には、プロトコル (例: `<AppDomain>https://myappdomain</AppDomain>`) が含まれている必要があります。</span><span class="sxs-lookup"><span data-stu-id="5f342-108">The value of the **AppDomain** element must include the protocol (e.g., `<AppDomain>https://myappdomain</AppDomain>`).</span></span>
> 2. <span data-ttu-id="5f342-109">値には、末尾にスラッシュ "/" を付け*ない*でください。</span><span class="sxs-lookup"><span data-stu-id="5f342-109">Do *not* put a closing slash, "/", on the value.</span></span>

## <a name="contained-in"></a><span data-ttu-id="5f342-110">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="5f342-110">Contained in</span></span>

[<span data-ttu-id="5f342-111">AppDomains</span><span class="sxs-lookup"><span data-stu-id="5f342-111">AppDomains</span></span>](appdomains.md)

## <a name="remarks"></a><span data-ttu-id="5f342-112">解説</span><span class="sxs-lookup"><span data-stu-id="5f342-112">Remarks</span></span>

<span data-ttu-id="5f342-113">**AppDomain** 要素は、[SourceLocation](sourcelocation.md) 要素で指定したドメイン以外のものを追加指定するために使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="5f342-113">**AppDomain** elements should be used to specify any additional domains other than the one specified in the [SourceLocation element](sourcelocation.md).</span></span> <span data-ttu-id="5f342-114">詳細については、「[Office アドイン XML マニフェスト](../../develop/add-in-manifests.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="5f342-114">For more information, see [Office Add-ins XML manifest](../../develop/add-in-manifests.md).</span></span>
