---
title: マニフェストファイル内のランタイム
description: ''
ms.date: 01/06/2020
localization_priority: Normal
ms.openlocfilehash: ec2b85a92325eb4e36c61f731369ec54d44ef169
ms.sourcegitcommit: 0dacbe7c80ed387099e3ec21e151f8990b181ede
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/13/2020
ms.locfileid: "41111178"
---
# <a name="runtimes-element"></a><span data-ttu-id="e4baf-102">ランタイム要素</span><span class="sxs-lookup"><span data-stu-id="e4baf-102">Runtimes element</span></span>

<span data-ttu-id="e4baf-103">この機能はプレビュー段階です。</span><span class="sxs-lookup"><span data-stu-id="e4baf-103">This feature is in preview.</span></span> <span data-ttu-id="e4baf-104">アドインのランタイムを指定し、カスタム関数と作業ウィンドウでグローバルデータを共有して、関数呼び出しを相互に行うことができるようにします。</span><span class="sxs-lookup"><span data-stu-id="e4baf-104">Specifies the runtime of your add-in and allows custom functions and the task pane to share global data and make function calls into each other.</span></span> <span data-ttu-id="e4baf-105">マニフェストファイルの`<Host>`要素に従う必要があります。</span><span class="sxs-lookup"><span data-stu-id="e4baf-105">Should follow the `<Host>` element in your manifest file.</span></span>

<span data-ttu-id="e4baf-106">**アドインの種類:** 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="e4baf-106">**Add-in type:** Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="e4baf-107">構文</span><span class="sxs-lookup"><span data-stu-id="e4baf-107">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="TaskPaneAndCustomFunction.Url" lifetime="long" />
</Runtimes>
```

## <a name="child-elements"></a><span data-ttu-id="e4baf-108">子要素</span><span class="sxs-lookup"><span data-stu-id="e4baf-108">Child elements</span></span>

|  <span data-ttu-id="e4baf-109">要素</span><span class="sxs-lookup"><span data-stu-id="e4baf-109">Element</span></span> |  <span data-ttu-id="e4baf-110">必須</span><span class="sxs-lookup"><span data-stu-id="e4baf-110">Required</span></span>  |  <span data-ttu-id="e4baf-111">説明</span><span class="sxs-lookup"><span data-stu-id="e4baf-111">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="e4baf-112">**ランタイム**</span><span class="sxs-lookup"><span data-stu-id="e4baf-112">**Runtime**</span></span>     | <span data-ttu-id="e4baf-113">はい</span><span class="sxs-lookup"><span data-stu-id="e4baf-113">Yes</span></span> |  <span data-ttu-id="e4baf-114">アドインのランタイム。多くの場合、Excel カスタム関数で使用されます。</span><span class="sxs-lookup"><span data-stu-id="e4baf-114">The Runtime for your add-in, often used with Excel custom functions.</span></span>

## <a name="see-also"></a><span data-ttu-id="e4baf-115">関連項目</span><span class="sxs-lookup"><span data-stu-id="e4baf-115">See also</span></span>

<span data-ttu-id="e4baf-116">-[ランタイム](runtimes.md)</span><span class="sxs-lookup"><span data-stu-id="e4baf-116">-[Runtimes](runtimes.md)</span></span>
