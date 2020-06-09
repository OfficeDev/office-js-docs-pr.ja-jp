---
title: マニフェストファイル内のランタイム
description: Runtime 要素は、アドインが、リボン、作業ウィンドウ、カスタム関数など、さまざまなコンポーネントに対して共有 JavaScript ランタイムを使用するように構成します。
ms.date: 05/29/2020
localization_priority: Normal
ms.openlocfilehash: e81bd7222585bfa7d5f0f34fe5d9b32e4d45a71e
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608105"
---
# <a name="runtime-element-preview"></a><span data-ttu-id="6acf6-103">Runtime 要素 (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="6acf6-103">Runtime element (preview)</span></span>

<span data-ttu-id="6acf6-104">共有された JavaScript ランタイムを使用するようにアドインを構成し、さまざまなコンポーネントがすべて同じランタイムで実行されるようにします。</span><span class="sxs-lookup"><span data-stu-id="6acf6-104">Configures your add-in to use a shared JavaScript runtime so that various components all run in the same runtime.</span></span> <span data-ttu-id="6acf6-105">要素の子 [`<Runtimes>`](runtimes.md) 。</span><span class="sxs-lookup"><span data-stu-id="6acf6-105">Child of the [`<Runtimes>`](runtimes.md) element.</span></span>

<span data-ttu-id="6acf6-106">Excel では、この要素を使用すると、リボン、作業ウィンドウ、およびカスタム関数が同じランタイムを使用できるようになります。</span><span class="sxs-lookup"><span data-stu-id="6acf6-106">In Excel, this element enables the ribbon, task pane, and custom functions to use the same runtime.</span></span> <span data-ttu-id="6acf6-107">詳細については、「[共有 JavaScript ランタイムを使用するように Excel アドインを構成する](../../excel/configure-your-add-in-to-use-a-shared-runtime.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="6acf6-107">For more information, see [Configure your Excel add-in to use a shared JavaScript runtime](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

<span data-ttu-id="6acf6-108">Outlook では、この要素はイベントベースのアドインのアクティブ化を有効にします。</span><span class="sxs-lookup"><span data-stu-id="6acf6-108">In Outlook, this element enables event-based add-in activation.</span></span> <span data-ttu-id="6acf6-109">詳細については、「[イベントベースのライセンス認証用に Outlook アドインを構成する](../../outlook/autolaunch.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="6acf6-109">For more information, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).</span></span>

<span data-ttu-id="6acf6-110">**アドインの種類:** 作業ウィンドウ、メール</span><span class="sxs-lookup"><span data-stu-id="6acf6-110">**Add-in type:** Task pane, Mail</span></span>

> [!IMPORTANT]
> <span data-ttu-id="6acf6-111">**Excel**: 共有ランタイムは、現在 Windows 上の Excel でのみ使用できます。</span><span class="sxs-lookup"><span data-stu-id="6acf6-111">**Excel**: Shared runtime is currently only available in Excel on Windows.</span></span>
>
> <span data-ttu-id="6acf6-112">**Outlook**: イベントベースのライセンス認証は現在[プレビュー段階で](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md)あり、web 上の Outlook でのみ使用できます。</span><span class="sxs-lookup"><span data-stu-id="6acf6-112">**Outlook**: Event-based activation is currently [in preview](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) and only available in Outlook on the web.</span></span> <span data-ttu-id="6acf6-113">詳細については、「[イベントベースのライセンス認証機能をプレビューする方法](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="6acf6-113">For more information, see [How to preview the event-based activation feature](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span></span>

## <a name="syntax"></a><span data-ttu-id="6acf6-114">構文</span><span class="sxs-lookup"><span data-stu-id="6acf6-114">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="6acf6-115">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="6acf6-115">Contained in</span></span>

- [<span data-ttu-id="6acf6-116">ランタイム</span><span class="sxs-lookup"><span data-stu-id="6acf6-116">Runtimes</span></span>](runtimes.md)

## <a name="attributes"></a><span data-ttu-id="6acf6-117">属性</span><span class="sxs-lookup"><span data-stu-id="6acf6-117">Attributes</span></span>

|  <span data-ttu-id="6acf6-118">属性</span><span class="sxs-lookup"><span data-stu-id="6acf6-118">Attribute</span></span>  |  <span data-ttu-id="6acf6-119">必須</span><span class="sxs-lookup"><span data-stu-id="6acf6-119">Required</span></span>  |  <span data-ttu-id="6acf6-120">説明</span><span class="sxs-lookup"><span data-stu-id="6acf6-120">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="6acf6-121">**resid**</span><span class="sxs-lookup"><span data-stu-id="6acf6-121">**resid**</span></span>  |  <span data-ttu-id="6acf6-122">はい</span><span class="sxs-lookup"><span data-stu-id="6acf6-122">Yes</span></span>  | <span data-ttu-id="6acf6-123">アドインの HTML ページの URL の場所を指定します。</span><span class="sxs-lookup"><span data-stu-id="6acf6-123">Specifies the URL location of the HTML page for your add-in.</span></span> <span data-ttu-id="6acf6-124">は、 `resid` `id` 要素内の要素の属性と一致している必要があり `Url` `Resources` ます。</span><span class="sxs-lookup"><span data-stu-id="6acf6-124">The `resid` must match an `id` attribute of a `Url` element in the `Resources` element.</span></span> |
|  <span data-ttu-id="6acf6-125">**時間**</span><span class="sxs-lookup"><span data-stu-id="6acf6-125">**lifetime**</span></span>  |  <span data-ttu-id="6acf6-126">いいえ</span><span class="sxs-lookup"><span data-stu-id="6acf6-126">No</span></span>  | <span data-ttu-id="6acf6-127">の既定値は、を `lifetime` `short` 指定する必要はありません。</span><span class="sxs-lookup"><span data-stu-id="6acf6-127">The default value for `lifetime` is `short` and doesn't need to be specified.</span></span> <span data-ttu-id="6acf6-128">Outlook アドインは、値のみを使用し `short` ます。</span><span class="sxs-lookup"><span data-stu-id="6acf6-128">Outlook add-ins use only the `short` value.</span></span> <span data-ttu-id="6acf6-129">Excel アドインで共有ランタイムを使用する場合は、の値をに明示的に設定し `long` ます。</span><span class="sxs-lookup"><span data-stu-id="6acf6-129">If you want to use a shared runtime in an Excel add-in, explicitly set the value to `long`.</span></span> |

## <a name="see-also"></a><span data-ttu-id="6acf6-130">関連項目</span><span class="sxs-lookup"><span data-stu-id="6acf6-130">See also</span></span>

- [<span data-ttu-id="6acf6-131">ランタイム</span><span class="sxs-lookup"><span data-stu-id="6acf6-131">Runtimes</span></span>](runtimes.md)
