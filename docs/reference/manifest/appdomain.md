---
title: マニフェスト ファイルの AppDomain 要素
description: ''
ms.date: 12/13/2018
ms.openlocfilehash: 2b55f2c1ea7a2a3dc7dec42c913d74006c0f2e3b
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433069"
---
# <a name="appdomain-element"></a><span data-ttu-id="f0a6a-102">AppDomain 要素</span><span class="sxs-lookup"><span data-stu-id="f0a6a-102">AppDomain element</span></span>

<span data-ttu-id="f0a6a-103">アドイン ウィンドウにページを読み込むために使用される追加のドメインを指定します。</span><span class="sxs-lookup"><span data-stu-id="f0a6a-103">Specifies an additional domain that will be used to load pages in the add-in window.</span></span>

<span data-ttu-id="f0a6a-104">**アドインの種類:** コンテンツ、作業ウィンドウ、メール</span><span class="sxs-lookup"><span data-stu-id="f0a6a-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="f0a6a-105">構文</span><span class="sxs-lookup"><span data-stu-id="f0a6a-105">Syntax</span></span>

```XML
<AppDomain>string</AppDomain>
```

> [!IMPORTANT]
> <span data-ttu-id="f0a6a-106">**AppDomain** 要素の値には、プロトコル (例: `<AppDomain>https://myappdomain<AppDomain>`) が含まれている必要があります。</span><span class="sxs-lookup"><span data-stu-id="f0a6a-106">The value of the **AppDomain** element must include the protocol (e.g., `<AppDomain>https://myappdomain<AppDomain>`).</span></span>

## <a name="contained-in"></a><span data-ttu-id="f0a6a-107">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="f0a6a-107">Contained in</span></span>

[<span data-ttu-id="f0a6a-108">AppDomains</span><span class="sxs-lookup"><span data-stu-id="f0a6a-108">AppDomains</span></span>](appdomains.md)

## <a name="remarks"></a><span data-ttu-id="f0a6a-109">解説</span><span class="sxs-lookup"><span data-stu-id="f0a6a-109">Remarks</span></span>

<span data-ttu-id="f0a6a-110">**AppDomain** 要素は、[SourceLocation](sourcelocation.md) 要素で指定したドメイン以外のものを追加指定するために使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="f0a6a-110">**AppDomain** elements should be used to specify any additional domains other than the one specified in the [SourceLocation element](sourcelocation.md).</span></span> <span data-ttu-id="f0a6a-111">詳細については、「[Office アドイン XML マニフェスト](/office/dev/add-ins/develop/add-in-manifests)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="f0a6a-111">For more information, see [Office Add-ins XML manifest](/office/dev/add-ins/develop/add-in-manifests).</span></span>
