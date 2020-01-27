---
title: マニフェストファイル内のランタイム
description: ''
ms.date: 01/24/2020
localization_priority: Normal
ms.openlocfilehash: 8fbad8276b3e1d64a6c443cf57d498597d729282
ms.sourcegitcommit: 72d719165cc2b64ac9d3c51fb8be277dfde7d2eb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/25/2020
ms.locfileid: "41554000"
---
# <a name="runtime-element"></a><span data-ttu-id="5f80f-102">Runtime 要素</span><span class="sxs-lookup"><span data-stu-id="5f80f-102">Runtime element</span></span>

<span data-ttu-id="5f80f-103">この機能はプレビュー段階です。</span><span class="sxs-lookup"><span data-stu-id="5f80f-103">This feature is in preview.</span></span> <span data-ttu-id="5f80f-104">[`<Runtimes>`](runtimes.md)要素の子要素。</span><span class="sxs-lookup"><span data-stu-id="5f80f-104">Child element of the [`<Runtimes>`](runtimes.md) element.</span></span> <span data-ttu-id="5f80f-105">この要素を使用すると、Excel カスタム関数とアドインの作業ウィンドウの間でのグローバルデータと関数呼び出しの共有が容易になります。</span><span class="sxs-lookup"><span data-stu-id="5f80f-105">This element facilitates sharing of global data and function calls between Excel custom functions and the task pane of your add-in.</span></span>

<span data-ttu-id="5f80f-106">**アドインの種類:** 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="5f80f-106">**Add-in type:** Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="5f80f-107">構文</span><span class="sxs-lookup"><span data-stu-id="5f80f-107">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="TaskPaneAndCustomFunction.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="5f80f-108">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="5f80f-108">Contained in</span></span>

- [<span data-ttu-id="5f80f-109">ランタイム</span><span class="sxs-lookup"><span data-stu-id="5f80f-109">Runtimes</span></span>](runtimes.md)

## <a name="attributes"></a><span data-ttu-id="5f80f-110">属性</span><span class="sxs-lookup"><span data-stu-id="5f80f-110">Attributes</span></span>

|  <span data-ttu-id="5f80f-111">属性</span><span class="sxs-lookup"><span data-stu-id="5f80f-111">Attribute</span></span>  |  <span data-ttu-id="5f80f-112">必須</span><span class="sxs-lookup"><span data-stu-id="5f80f-112">Required</span></span>  |  <span data-ttu-id="5f80f-113">説明</span><span class="sxs-lookup"><span data-stu-id="5f80f-113">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="5f80f-114">**lifetime = "long"**</span><span class="sxs-lookup"><span data-stu-id="5f80f-114">**lifetime="long"**</span></span>  |  <span data-ttu-id="5f80f-115">はい</span><span class="sxs-lookup"><span data-stu-id="5f80f-115">Yes</span></span>  | <span data-ttu-id="5f80f-116">アドインの作業ウィンドウが閉じているときに Excel カスタム関数が動作するようにする場合は、常に long として表示される必要があります。</span><span class="sxs-lookup"><span data-stu-id="5f80f-116">Should always be listed as long if you want Excel custom functions to work while the task pane of your add-in is closed.</span></span> |
|  <span data-ttu-id="5f80f-117">**resid**</span><span class="sxs-lookup"><span data-stu-id="5f80f-117">**resid**</span></span>  |  <span data-ttu-id="5f80f-118">はい</span><span class="sxs-lookup"><span data-stu-id="5f80f-118">Yes</span></span>  | <span data-ttu-id="5f80f-119">Excel カスタム関数で使用する場合、 `resid`はを`TaskPaneAndCustomFunction.Url`参照する必要があります。</span><span class="sxs-lookup"><span data-stu-id="5f80f-119">If used for Excel custom functions, the `resid` should point to `TaskPaneAndCustomFunction.Url`.</span></span> |

## <a name="see-also"></a><span data-ttu-id="5f80f-120">関連項目</span><span class="sxs-lookup"><span data-stu-id="5f80f-120">See also</span></span>

- [<span data-ttu-id="5f80f-121">ランタイム</span><span class="sxs-lookup"><span data-stu-id="5f80f-121">Runtimes</span></span>](runtimes.md)
