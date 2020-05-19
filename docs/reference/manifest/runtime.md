---
title: マニフェストファイル内のランタイム
description: Runtime 要素は、アドインが、リボン、作業ウィンドウ、カスタム関数など、さまざまなコンポーネントに対して共有 JavaScript ランタイムを使用するように構成します。
ms.date: 05/18/2020
localization_priority: Normal
ms.openlocfilehash: c2c404bcaad6e24af58f5c0ed8835343abb97e5f
ms.sourcegitcommit: f62d9630de69c5c070e3d4048205f5cc654db7e4
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/18/2020
ms.locfileid: "44278414"
---
# <a name="runtime-element-preview"></a><span data-ttu-id="c37ca-103">Runtime 要素 (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="c37ca-103">Runtime element (preview)</span></span>

<span data-ttu-id="c37ca-104">共有された JavaScript ランタイムを使用するようにアドインを構成し、さまざまなコンポーネントがすべて同じランタイムで実行されるようにします。</span><span class="sxs-lookup"><span data-stu-id="c37ca-104">Configures your add-in to use a shared JavaScript runtime so that various components all run in the same runtime.</span></span> <span data-ttu-id="c37ca-105">要素の子 [`<Runtimes>`](runtimes.md) 。</span><span class="sxs-lookup"><span data-stu-id="c37ca-105">Child of the [`<Runtimes>`](runtimes.md) element.</span></span>

<span data-ttu-id="c37ca-106">Excel では、この要素を使用すると、リボン、作業ウィンドウ、およびカスタム関数が同じランタイムを使用できるようになります。</span><span class="sxs-lookup"><span data-stu-id="c37ca-106">In Excel, this element enables the ribbon, task pane, and custom functions to use the same runtime.</span></span> <span data-ttu-id="c37ca-107">詳細については、「[共有 JavaScript ランタイムを使用するように Excel アドインを構成する](../../excel/configure-your-add-in-to-use-a-shared-runtime.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="c37ca-107">For more information, see [Configure your Excel add-in to use a shared JavaScript runtime](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

<span data-ttu-id="c37ca-108">Outlook では、この要素はイベントベースのアドインのアクティブ化を有効にします。</span><span class="sxs-lookup"><span data-stu-id="c37ca-108">In Outlook, this element enables event-based add-in activation.</span></span> <span data-ttu-id="c37ca-109">詳細については、「[イベントベースのライセンス認証用に Outlook アドインを構成する](../../outlook/autolaunch.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="c37ca-109">For more information, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).</span></span>

<span data-ttu-id="c37ca-110">**アドインの種類:** 作業ウィンドウ、メール</span><span class="sxs-lookup"><span data-stu-id="c37ca-110">**Add-in type:** Task pane, Mail</span></span>

> [!IMPORTANT]
> <span data-ttu-id="c37ca-111">**Excel**: 共有ランタイムは現在プレビュー段階であり、Windows 上の Excel でのみ使用できます。</span><span class="sxs-lookup"><span data-stu-id="c37ca-111">**Excel**: Shared runtime is currently in preview and only available in Excel on Windows.</span></span> <span data-ttu-id="c37ca-112">プレビュー機能を試すには、 [Office Insider](https://insider.office.com/)に参加する必要があります。</span><span class="sxs-lookup"><span data-stu-id="c37ca-112">To try the preview features, you will need to join [Office Insider](https://insider.office.com/).</span></span>
>
> <span data-ttu-id="c37ca-113">**Outlook**: イベントベースのライセンス認証は現在[プレビュー段階で](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md)あり、web 上の Outlook でのみ使用できます。</span><span class="sxs-lookup"><span data-stu-id="c37ca-113">**Outlook**: Event-based activation is currently [in preview](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) and only available in Outlook on the web.</span></span> <span data-ttu-id="c37ca-114">詳細については、「[イベントベースのライセンス認証機能をプレビューする方法](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="c37ca-114">For more information, see [How to preview the event-based activation feature](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span></span>

## <a name="syntax"></a><span data-ttu-id="c37ca-115">構文</span><span class="sxs-lookup"><span data-stu-id="c37ca-115">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="c37ca-116">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="c37ca-116">Contained in</span></span>

- [<span data-ttu-id="c37ca-117">ランタイム</span><span class="sxs-lookup"><span data-stu-id="c37ca-117">Runtimes</span></span>](runtimes.md)

## <a name="attributes"></a><span data-ttu-id="c37ca-118">属性</span><span class="sxs-lookup"><span data-stu-id="c37ca-118">Attributes</span></span>

|  <span data-ttu-id="c37ca-119">属性</span><span class="sxs-lookup"><span data-stu-id="c37ca-119">Attribute</span></span>  |  <span data-ttu-id="c37ca-120">必須</span><span class="sxs-lookup"><span data-stu-id="c37ca-120">Required</span></span>  |  <span data-ttu-id="c37ca-121">説明</span><span class="sxs-lookup"><span data-stu-id="c37ca-121">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="c37ca-122">**resid**</span><span class="sxs-lookup"><span data-stu-id="c37ca-122">**resid**</span></span>  |  <span data-ttu-id="c37ca-123">はい</span><span class="sxs-lookup"><span data-stu-id="c37ca-123">Yes</span></span>  | <span data-ttu-id="c37ca-124">アドインの HTML ページの URL の場所を指定します。</span><span class="sxs-lookup"><span data-stu-id="c37ca-124">Specifies the URL location of the HTML page for your add-in.</span></span> <span data-ttu-id="c37ca-125">は、 `resid` `id` 要素内の要素の属性と一致している必要があり `Url` `Resources` ます。</span><span class="sxs-lookup"><span data-stu-id="c37ca-125">The `resid` must match an `id` attribute of a `Url` element in the `Resources` element.</span></span> |
|  <span data-ttu-id="c37ca-126">**時間**</span><span class="sxs-lookup"><span data-stu-id="c37ca-126">**lifetime**</span></span>  |  <span data-ttu-id="c37ca-127">いいえ</span><span class="sxs-lookup"><span data-stu-id="c37ca-127">No</span></span>  | <span data-ttu-id="c37ca-128">の既定値は、を `lifetime` `short` 指定する必要はありません。</span><span class="sxs-lookup"><span data-stu-id="c37ca-128">The default value for `lifetime` is `short` and doesn't need to be specified.</span></span> <span data-ttu-id="c37ca-129">Outlook アドインは、値のみを使用し `short` ます。</span><span class="sxs-lookup"><span data-stu-id="c37ca-129">Outlook add-ins use only the `short` value.</span></span> <span data-ttu-id="c37ca-130">Excel アドインで共有ランタイムを使用する場合は、の値をに明示的に設定し `long` ます。</span><span class="sxs-lookup"><span data-stu-id="c37ca-130">If you want to use a shared runtime in an Excel add-in, explicitly set the value to `long`.</span></span> |

## <a name="see-also"></a><span data-ttu-id="c37ca-131">関連項目</span><span class="sxs-lookup"><span data-stu-id="c37ca-131">See also</span></span>

- [<span data-ttu-id="c37ca-132">ランタイム</span><span class="sxs-lookup"><span data-stu-id="c37ca-132">Runtimes</span></span>](runtimes.md)
