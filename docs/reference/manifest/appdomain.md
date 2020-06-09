---
title: マニフェスト ファイルの AppDomain 要素
description: アドインウィンドウにページを読み込む追加のドメインを指定します。
ms.date: 07/03/2019
localization_priority: Normal
ms.openlocfilehash: ddacae6d8aa45ccccd3a8acbb42de48b152fb9d2
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608776"
---
# <a name="appdomain-element"></a><span data-ttu-id="69a52-103">AppDomain 要素</span><span class="sxs-lookup"><span data-stu-id="69a52-103">AppDomain element</span></span>

<span data-ttu-id="69a52-104">アドインウィンドウにページを読み込む追加のドメインを指定します。</span><span class="sxs-lookup"><span data-stu-id="69a52-104">Specifies additional domains that load pages in the add-in window.</span></span> <span data-ttu-id="69a52-105">また、アドイン内の Iframe から Office .js API 呼び出しを行うことができる信頼されたドメインも一覧表示されます。</span><span class="sxs-lookup"><span data-stu-id="69a52-105">It also lists trusted domains from which Office.js API calls can be made from IFrames within the add-in.</span></span>

<span data-ttu-id="69a52-106">**アドインの種類:** コンテンツ、作業ウィンドウ、メール</span><span class="sxs-lookup"><span data-stu-id="69a52-106">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="69a52-107">構文</span><span class="sxs-lookup"><span data-stu-id="69a52-107">Syntax</span></span>

```XML
<AppDomain>string</AppDomain>
```

> [!IMPORTANT]
> 1. <span data-ttu-id="69a52-108">**AppDomain** 要素の値には、プロトコル (例: `<AppDomain>https://myappdomain</AppDomain>`) が含まれている必要があります。</span><span class="sxs-lookup"><span data-stu-id="69a52-108">The value of the **AppDomain** element must include the protocol (e.g., `<AppDomain>https://myappdomain</AppDomain>`).</span></span>
> 2. <span data-ttu-id="69a52-109">値には、末尾にスラッシュ "/" を付け*ない*でください。</span><span class="sxs-lookup"><span data-stu-id="69a52-109">Do *not* put a closing slash, "/", on the value.</span></span>

## <a name="contained-in"></a><span data-ttu-id="69a52-110">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="69a52-110">Contained in</span></span>

[<span data-ttu-id="69a52-111">AppDomains</span><span class="sxs-lookup"><span data-stu-id="69a52-111">AppDomains</span></span>](appdomains.md)

## <a name="remarks"></a><span data-ttu-id="69a52-112">解説</span><span class="sxs-lookup"><span data-stu-id="69a52-112">Remarks</span></span>

<span data-ttu-id="69a52-113">**AppDomain** 要素は、[SourceLocation](sourcelocation.md) 要素で指定したドメイン以外のものを追加指定するために使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="69a52-113">**AppDomain** elements should be used to specify any additional domains other than the one specified in the [SourceLocation element](sourcelocation.md).</span></span> <span data-ttu-id="69a52-114">詳細については、「[Office アドイン XML マニフェスト](../../develop/add-in-manifests.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="69a52-114">For more information, see [Office Add-ins XML manifest](../../develop/add-in-manifests.md).</span></span>
