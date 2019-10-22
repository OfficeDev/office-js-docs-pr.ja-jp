---
title: マニフェスト ファイルの AppDomain 要素
description: ''
ms.date: 07/03/2019
localization_priority: Normal
ms.openlocfilehash: 2f65302d1ac3d85f2867cd13501bc67606cd00b5
ms.sourcegitcommit: b3996b1444e520b44cf752e76eef50908386ca26
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/21/2019
ms.locfileid: "35575640"
---
# <a name="appdomain-element"></a><span data-ttu-id="5327a-102">AppDomain 要素</span><span class="sxs-lookup"><span data-stu-id="5327a-102">AppDomain element</span></span>

<span data-ttu-id="5327a-103">アドインウィンドウにページを読み込む追加のドメインを指定します。</span><span class="sxs-lookup"><span data-stu-id="5327a-103">Specifies additional domains that load pages in the add-in window.</span></span> <span data-ttu-id="5327a-104">また、アドイン内の Iframe から Office .js API 呼び出しを行うことができる信頼されたドメインも一覧表示されます。</span><span class="sxs-lookup"><span data-stu-id="5327a-104">It also lists trusted domains from which Office.js API calls can be made from IFrames within the add-in.</span></span>

<span data-ttu-id="5327a-105">**アドインの種類:** コンテンツ、作業ウィンドウ、メール</span><span class="sxs-lookup"><span data-stu-id="5327a-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="5327a-106">構文</span><span class="sxs-lookup"><span data-stu-id="5327a-106">Syntax</span></span>

```XML
<AppDomain>string</AppDomain>
```

> [!IMPORTANT]
> 1. <span data-ttu-id="5327a-107">**AppDomain** 要素の値には、プロトコル (例: `<AppDomain>https://myappdomain</AppDomain>`) が含まれている必要があります。</span><span class="sxs-lookup"><span data-stu-id="5327a-107">The value of the **AppDomain** element must include the protocol (e.g., `<AppDomain>https://myappdomain</AppDomain>`).</span></span>
> 2. <span data-ttu-id="5327a-108">値には、末尾にスラッシュ "/" を付け*ない*でください。</span><span class="sxs-lookup"><span data-stu-id="5327a-108">Do *not* put a closing slash, "/", on the value.</span></span>

## <a name="contained-in"></a><span data-ttu-id="5327a-109">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="5327a-109">Contained in</span></span>

[<span data-ttu-id="5327a-110">AppDomains</span><span class="sxs-lookup"><span data-stu-id="5327a-110">AppDomains</span></span>](appdomains.md)

## <a name="remarks"></a><span data-ttu-id="5327a-111">解説</span><span class="sxs-lookup"><span data-stu-id="5327a-111">Remarks</span></span>

<span data-ttu-id="5327a-112">**AppDomain** 要素は、[SourceLocation](sourcelocation.md) 要素で指定したドメイン以外のものを追加指定するために使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="5327a-112">**AppDomain** elements should be used to specify any additional domains other than the one specified in the [SourceLocation element](sourcelocation.md).</span></span> <span data-ttu-id="5327a-113">詳細については、「[Office アドイン XML マニフェスト](/office/dev/add-ins/develop/add-in-manifests)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="5327a-113">For more information, see [Office Add-ins XML manifest](/office/dev/add-ins/develop/add-in-manifests).</span></span>
