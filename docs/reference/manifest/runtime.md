---
title: マニフェストファイル内のランタイム
description: Runtime 要素は、アドインが、リボン、作業ウィンドウ、カスタム関数など、さまざまなコンポーネントに対して共有 JavaScript ランタイムを使用するように構成します。
ms.date: 05/29/2020
localization_priority: Normal
ms.openlocfilehash: 9e6e13f83db363fb5485c8d8defbc381c80e32d6
ms.sourcegitcommit: 472b81642e9eb5fb2a55cd98a7b0826d37eb7f73
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/17/2020
ms.locfileid: "45159368"
---
# <a name="runtime-element-preview"></a><span data-ttu-id="1377d-103">Runtime 要素 (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="1377d-103">Runtime element (preview)</span></span>

<span data-ttu-id="1377d-104">共有された JavaScript ランタイムを使用するようにアドインを構成し、さまざまなコンポーネントがすべて同じランタイムで実行されるようにします。</span><span class="sxs-lookup"><span data-stu-id="1377d-104">Configures your add-in to use a shared JavaScript runtime so that various components all run in the same runtime.</span></span> <span data-ttu-id="1377d-105">要素の子 [`<Runtimes>`](runtimes.md) 。</span><span class="sxs-lookup"><span data-stu-id="1377d-105">Child of the [`<Runtimes>`](runtimes.md) element.</span></span>

<span data-ttu-id="1377d-106">Excel では、この要素を使用すると、リボン、作業ウィンドウ、およびカスタム関数が同じランタイムを使用できるようになります。</span><span class="sxs-lookup"><span data-stu-id="1377d-106">In Excel, this element enables the ribbon, task pane, and custom functions to use the same runtime.</span></span> <span data-ttu-id="1377d-107">詳細については、「[共有 JavaScript ランタイムを使用するように Excel アドインを構成する](../../excel/configure-your-add-in-to-use-a-shared-runtime.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="1377d-107">For more information, see [Configure your Excel add-in to use a shared JavaScript runtime](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

<span data-ttu-id="1377d-108">Outlook では、この要素はイベントベースのアドインのアクティブ化を有効にします。</span><span class="sxs-lookup"><span data-stu-id="1377d-108">In Outlook, this element enables event-based add-in activation.</span></span> <span data-ttu-id="1377d-109">詳細については、「[イベントベースのライセンス認証用に Outlook アドインを構成する](../../outlook/autolaunch.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="1377d-109">For more information, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).</span></span>

<span data-ttu-id="1377d-110">**アドインの種類:** 作業ウィンドウ、メール</span><span class="sxs-lookup"><span data-stu-id="1377d-110">**Add-in type:** Task pane, Mail</span></span>

> [!IMPORTANT]
> <span data-ttu-id="1377d-111">**Outlook**: イベントベースのライセンス認証は現在[プレビュー段階で](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md)あり、web 上の Outlook でのみ使用できます。</span><span class="sxs-lookup"><span data-stu-id="1377d-111">**Outlook**: Event-based activation is currently [in preview](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) and only available in Outlook on the web.</span></span> <span data-ttu-id="1377d-112">詳細については、「[イベントベースのライセンス認証機能をプレビューする方法](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="1377d-112">For more information, see [How to preview the event-based activation feature](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span></span>

## <a name="syntax"></a><span data-ttu-id="1377d-113">構文</span><span class="sxs-lookup"><span data-stu-id="1377d-113">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="1377d-114">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="1377d-114">Contained in</span></span>

- [<span data-ttu-id="1377d-115">ランタイム</span><span class="sxs-lookup"><span data-stu-id="1377d-115">Runtimes</span></span>](runtimes.md)

## <a name="attributes"></a><span data-ttu-id="1377d-116">属性</span><span class="sxs-lookup"><span data-stu-id="1377d-116">Attributes</span></span>

|  <span data-ttu-id="1377d-117">属性</span><span class="sxs-lookup"><span data-stu-id="1377d-117">Attribute</span></span>  |  <span data-ttu-id="1377d-118">必須</span><span class="sxs-lookup"><span data-stu-id="1377d-118">Required</span></span>  |  <span data-ttu-id="1377d-119">説明</span><span class="sxs-lookup"><span data-stu-id="1377d-119">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="1377d-120">**resid**</span><span class="sxs-lookup"><span data-stu-id="1377d-120">**resid**</span></span>  |  <span data-ttu-id="1377d-121">はい</span><span class="sxs-lookup"><span data-stu-id="1377d-121">Yes</span></span>  | <span data-ttu-id="1377d-122">アドインの HTML ページの URL の場所を指定します。</span><span class="sxs-lookup"><span data-stu-id="1377d-122">Specifies the URL location of the HTML page for your add-in.</span></span> <span data-ttu-id="1377d-123">は、 `resid` `id` 要素内の要素の属性と一致している必要があり `Url` `Resources` ます。</span><span class="sxs-lookup"><span data-stu-id="1377d-123">The `resid` must match an `id` attribute of a `Url` element in the `Resources` element.</span></span> |
|  <span data-ttu-id="1377d-124">**時間**</span><span class="sxs-lookup"><span data-stu-id="1377d-124">**lifetime**</span></span>  |  <span data-ttu-id="1377d-125">不要</span><span class="sxs-lookup"><span data-stu-id="1377d-125">No</span></span>  | <span data-ttu-id="1377d-126">の既定値は、を `lifetime` `short` 指定する必要はありません。</span><span class="sxs-lookup"><span data-stu-id="1377d-126">The default value for `lifetime` is `short` and doesn't need to be specified.</span></span> <span data-ttu-id="1377d-127">Outlook アドインは、値のみを使用し `short` ます。</span><span class="sxs-lookup"><span data-stu-id="1377d-127">Outlook add-ins use only the `short` value.</span></span> <span data-ttu-id="1377d-128">Excel アドインで共有ランタイムを使用する場合は、の値をに明示的に設定し `long` ます。</span><span class="sxs-lookup"><span data-stu-id="1377d-128">If you want to use a shared runtime in an Excel add-in, explicitly set the value to `long`.</span></span> |

## <a name="see-also"></a><span data-ttu-id="1377d-129">関連項目</span><span class="sxs-lookup"><span data-stu-id="1377d-129">See also</span></span>

- [<span data-ttu-id="1377d-130">ランタイム</span><span class="sxs-lookup"><span data-stu-id="1377d-130">Runtimes</span></span>](runtimes.md)
