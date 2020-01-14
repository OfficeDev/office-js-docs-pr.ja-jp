---
title: マニフェストファイル内のランタイム
description: ''
ms.date: 01/06/2020
localization_priority: Normal
ms.openlocfilehash: 68def44ba74733934198ac3b32fa1fe649156766
ms.sourcegitcommit: 0dacbe7c80ed387099e3ec21e151f8990b181ede
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/13/2020
ms.locfileid: "41111171"
---
# <a name="runtime-element"></a><span data-ttu-id="540f6-102">Runtime 要素</span><span class="sxs-lookup"><span data-stu-id="540f6-102">Runtime element</span></span>

<span data-ttu-id="540f6-103">この機能はプレビュー段階です。</span><span class="sxs-lookup"><span data-stu-id="540f6-103">This feature is in preview.</span></span> <span data-ttu-id="540f6-104">[`<Runtimes>`](runtime.md)要素の子要素。</span><span class="sxs-lookup"><span data-stu-id="540f6-104">Child element of the [`<Runtimes>`](runtime.md) element.</span></span> <span data-ttu-id="540f6-105">この要素を使用すると、Excel カスタム関数とアドインの作業ウィンドウの間でのグローバルデータと関数呼び出しの共有が容易になります。</span><span class="sxs-lookup"><span data-stu-id="540f6-105">This element facilitates sharing of global data and function calls between Excel custom functions and the task pane of your add-in.</span></span> 

## <a name="contained-in"></a><span data-ttu-id="540f6-106">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="540f6-106">Contained in</span></span>

<span data-ttu-id="540f6-107">-[ランタイム](runtimes.md)</span><span class="sxs-lookup"><span data-stu-id="540f6-107">-[Runtimes](runtimes.md)</span></span>

<span data-ttu-id="540f6-108">**アドインの種類:** 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="540f6-108">**Add-in type:** Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="540f6-109">構文</span><span class="sxs-lookup"><span data-stu-id="540f6-109">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="TaskPaneAndCustomFunction.Url" lifetime="long" />
</Runtimes>
```

## <a name="attributes"></a><span data-ttu-id="540f6-110">属性</span><span class="sxs-lookup"><span data-stu-id="540f6-110">Attributes</span></span>

|  <span data-ttu-id="540f6-111">属性</span><span class="sxs-lookup"><span data-stu-id="540f6-111">Attribute</span></span>  |  <span data-ttu-id="540f6-112">必須</span><span class="sxs-lookup"><span data-stu-id="540f6-112">Required</span></span>  |  <span data-ttu-id="540f6-113">説明</span><span class="sxs-lookup"><span data-stu-id="540f6-113">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="540f6-114">**lifetime = "long"**</span><span class="sxs-lookup"><span data-stu-id="540f6-114">**lifetime="long"**</span></span>  |  <span data-ttu-id="540f6-115">はい</span><span class="sxs-lookup"><span data-stu-id="540f6-115">Yes</span></span>  | <span data-ttu-id="540f6-116">アドインの作業ウィンドウが閉じているときに Excel カスタム関数が動作するようにする場合は、常に long として表示される必要があります。</span><span class="sxs-lookup"><span data-stu-id="540f6-116">Should always be listed as long if you want Excel custom functions to work while the task pane of your add-in is closed.</span></span> |
|  <span data-ttu-id="540f6-117">**resid**</span><span class="sxs-lookup"><span data-stu-id="540f6-117">**resid**</span></span>  |  <span data-ttu-id="540f6-118">はい</span><span class="sxs-lookup"><span data-stu-id="540f6-118">Yes</span></span>  | <span data-ttu-id="540f6-119">Excel カスタム関数で使用する場合、 `resid`はを`TaskPaneAndCustomFunction.Url`参照する必要があります。</span><span class="sxs-lookup"><span data-stu-id="540f6-119">If used for Excel custom functions, the `resid` should point to `TaskPaneAndCustomFunction.Url`.</span></span> |

## <a name="see-also"></a><span data-ttu-id="540f6-120">関連項目</span><span class="sxs-lookup"><span data-stu-id="540f6-120">See also</span></span>

<span data-ttu-id="540f6-121">-[ランタイム](runtime.md)</span><span class="sxs-lookup"><span data-stu-id="540f6-121">-[Runtime](runtime.md)</span></span>
