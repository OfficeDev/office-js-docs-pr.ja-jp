---
title: マニフェスト ファイル内のランタイム
description: Runtime 要素は、リボン、作業ウィンドウ、カスタム関数など、さまざまなコンポーネントに対して共有 JavaScript ランタイムを使用するアドインを構成します。
ms.date: 05/29/2020
localization_priority: Normal
ms.openlocfilehash: 3cabfacc665ccf6c0e4e796cb0e1fbc70c770ee3
ms.sourcegitcommit: 545888b08f57bb1babb05ccfd83b2b3286bdad5c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/08/2021
ms.locfileid: "49789185"
---
# <a name="runtime-element-preview"></a><span data-ttu-id="71e39-103">Runtime 要素 (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="71e39-103">Runtime element (preview)</span></span>

<span data-ttu-id="71e39-104">さまざまなコンポーネントが同じランタイムで実行されるのを確認するために、共有 JavaScript ランタイムを使用するアドインを構成します。</span><span class="sxs-lookup"><span data-stu-id="71e39-104">Configures your add-in to use a shared JavaScript runtime so that various components all run in the same runtime.</span></span> <span data-ttu-id="71e39-105">要素の [`<Runtimes>`](runtimes.md) 子。</span><span class="sxs-lookup"><span data-stu-id="71e39-105">Child of the [`<Runtimes>`](runtimes.md) element.</span></span>

<span data-ttu-id="71e39-106">Excel では、この要素により、リボン、作業ウィンドウ、およびカスタム関数で同じランタイムを使用できます。</span><span class="sxs-lookup"><span data-stu-id="71e39-106">In Excel, this element enables the ribbon, task pane, and custom functions to use the same runtime.</span></span> <span data-ttu-id="71e39-107">詳細については、「共有 [JavaScript ランタイムを使用するために Excel](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)アドインを構成する」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="71e39-107">For more information, see [Configure your Excel add-in to use a shared JavaScript runtime](../../develop/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

<span data-ttu-id="71e39-108">Outlook では、この要素によってイベント ベースのアドインのアクティブ化が有効になります。</span><span class="sxs-lookup"><span data-stu-id="71e39-108">In Outlook, this element enables event-based add-in activation.</span></span> <span data-ttu-id="71e39-109">詳細については、「イベント ベースのアクティブ [化用に Outlook アドインを構成する」を参照してください](../../outlook/autolaunch.md)。</span><span class="sxs-lookup"><span data-stu-id="71e39-109">For more information, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).</span></span>

<span data-ttu-id="71e39-110">**アドインの種類:** 作業ウィンドウ、メール</span><span class="sxs-lookup"><span data-stu-id="71e39-110">**Add-in type:** Task pane, Mail</span></span>

> [!IMPORTANT]
> <span data-ttu-id="71e39-111">**Outlook**: イベント ベースのアクティブ化は現在 [プレビュー中で](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) 、Outlook on the web でのみ使用できます。</span><span class="sxs-lookup"><span data-stu-id="71e39-111">**Outlook**: Event-based activation is currently [in preview](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) and only available in Outlook on the web.</span></span> <span data-ttu-id="71e39-112">詳細については、イベント ベースの [アクティブ化機能をプレビューする方法を参照してください](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature)。</span><span class="sxs-lookup"><span data-stu-id="71e39-112">For more information, see [How to preview the event-based activation feature](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span></span>

## <a name="syntax"></a><span data-ttu-id="71e39-113">構文</span><span class="sxs-lookup"><span data-stu-id="71e39-113">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="71e39-114">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="71e39-114">Contained in</span></span>

- [<span data-ttu-id="71e39-115">ランタイム</span><span class="sxs-lookup"><span data-stu-id="71e39-115">Runtimes</span></span>](runtimes.md)

## <a name="attributes"></a><span data-ttu-id="71e39-116">属性</span><span class="sxs-lookup"><span data-stu-id="71e39-116">Attributes</span></span>

|  <span data-ttu-id="71e39-117">属性</span><span class="sxs-lookup"><span data-stu-id="71e39-117">Attribute</span></span>  |  <span data-ttu-id="71e39-118">必須</span><span class="sxs-lookup"><span data-stu-id="71e39-118">Required</span></span>  |  <span data-ttu-id="71e39-119">説明</span><span class="sxs-lookup"><span data-stu-id="71e39-119">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="71e39-120">**resid**</span><span class="sxs-lookup"><span data-stu-id="71e39-120">**resid**</span></span>  |  <span data-ttu-id="71e39-121">はい</span><span class="sxs-lookup"><span data-stu-id="71e39-121">Yes</span></span>  | <span data-ttu-id="71e39-122">アドインの HTML ページの URL の場所を指定します。</span><span class="sxs-lookup"><span data-stu-id="71e39-122">Specifies the URL location of the HTML page for your add-in.</span></span> <span data-ttu-id="71e39-123">使用できる文字数は 32 文字以内で、要素内の要素の属性と一 `resid` `id` `Url` 致する必要 `Resources` があります。</span><span class="sxs-lookup"><span data-stu-id="71e39-123">The `resid` can be no more than 32 characters and must match an `id` attribute of a `Url` element in the `Resources` element.</span></span> |
|  <span data-ttu-id="71e39-124">**lifetime**</span><span class="sxs-lookup"><span data-stu-id="71e39-124">**lifetime**</span></span>  |  <span data-ttu-id="71e39-125">いいえ</span><span class="sxs-lookup"><span data-stu-id="71e39-125">No</span></span>  | <span data-ttu-id="71e39-126">既定値は次 `lifetime` の `short` 値で、指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="71e39-126">The default value for `lifetime` is `short` and doesn't need to be specified.</span></span> <span data-ttu-id="71e39-127">Outlook アドインは値のみを使用 `short` します。</span><span class="sxs-lookup"><span data-stu-id="71e39-127">Outlook add-ins use only the `short` value.</span></span> <span data-ttu-id="71e39-128">Excel アドインで共有ランタイムを使用する場合は、値を明示的に設定します `long` 。</span><span class="sxs-lookup"><span data-stu-id="71e39-128">If you want to use a shared runtime in an Excel add-in, explicitly set the value to `long`.</span></span> |

## <a name="see-also"></a><span data-ttu-id="71e39-129">関連項目</span><span class="sxs-lookup"><span data-stu-id="71e39-129">See also</span></span>

- [<span data-ttu-id="71e39-130">ランタイム</span><span class="sxs-lookup"><span data-stu-id="71e39-130">Runtimes</span></span>](runtimes.md)
