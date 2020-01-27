---
title: マニフェストファイル内のランタイム
description: ''
ms.date: 01/24/2020
localization_priority: Normal
ms.openlocfilehash: 6682887935ee6894b5a311ad519408067452bb23
ms.sourcegitcommit: 72d719165cc2b64ac9d3c51fb8be277dfde7d2eb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/25/2020
ms.locfileid: "41554007"
---
# <a name="runtimes-element"></a><span data-ttu-id="1d571-102">ランタイム要素</span><span class="sxs-lookup"><span data-stu-id="1d571-102">Runtimes element</span></span>

<span data-ttu-id="1d571-103">この機能はプレビュー段階です。</span><span class="sxs-lookup"><span data-stu-id="1d571-103">This feature is in preview.</span></span> <span data-ttu-id="1d571-104">アドインのランタイムを指定し、カスタム関数と作業ウィンドウでグローバルデータを共有して、関数呼び出しを相互に行うことができるようにします。</span><span class="sxs-lookup"><span data-stu-id="1d571-104">Specifies the runtime of your add-in and allows custom functions and the task pane to share global data and make function calls into each other.</span></span> <span data-ttu-id="1d571-105">マニフェストファイルの`<Host>`要素に従う必要があります。</span><span class="sxs-lookup"><span data-stu-id="1d571-105">Should follow the `<Host>` element in your manifest file.</span></span>

<span data-ttu-id="1d571-106">**アドインの種類:** 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="1d571-106">**Add-in type:** Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="1d571-107">構文</span><span class="sxs-lookup"><span data-stu-id="1d571-107">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="TaskPaneAndCustomFunction.Url" lifetime="long" />
</Runtimes>
```

## <a name="child-elements"></a><span data-ttu-id="1d571-108">子要素</span><span class="sxs-lookup"><span data-stu-id="1d571-108">Child elements</span></span>

|  <span data-ttu-id="1d571-109">要素</span><span class="sxs-lookup"><span data-stu-id="1d571-109">Element</span></span> |  <span data-ttu-id="1d571-110">必須</span><span class="sxs-lookup"><span data-stu-id="1d571-110">Required</span></span>  |  <span data-ttu-id="1d571-111">説明</span><span class="sxs-lookup"><span data-stu-id="1d571-111">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="1d571-112">**ランタイム**</span><span class="sxs-lookup"><span data-stu-id="1d571-112">**Runtime**</span></span>     | <span data-ttu-id="1d571-113">はい</span><span class="sxs-lookup"><span data-stu-id="1d571-113">Yes</span></span> |  <span data-ttu-id="1d571-114">アドインのランタイム。多くの場合、Excel カスタム関数で使用されます。</span><span class="sxs-lookup"><span data-stu-id="1d571-114">The Runtime for your add-in, often used with Excel custom functions.</span></span>

## <a name="see-also"></a><span data-ttu-id="1d571-115">関連項目</span><span class="sxs-lookup"><span data-stu-id="1d571-115">See also</span></span>

- [<span data-ttu-id="1d571-116">ランタイム</span><span class="sxs-lookup"><span data-stu-id="1d571-116">Runtime</span></span>](runtime.md)
