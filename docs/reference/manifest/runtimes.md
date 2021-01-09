---
title: マニフェスト ファイル内のランタイム
description: Runtimes 要素は、アドインのランタイムを指定します。
ms.date: 06/01/2020
localization_priority: Normal
ms.openlocfilehash: afbcc6a909c51d2ed56292ef1541193f7f698d28
ms.sourcegitcommit: 545888b08f57bb1babb05ccfd83b2b3286bdad5c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/08/2021
ms.locfileid: "49789164"
---
# <a name="runtimes-element"></a><span data-ttu-id="9f8f9-103">Runtimes 要素</span><span class="sxs-lookup"><span data-stu-id="9f8f9-103">Runtimes element</span></span>

<span data-ttu-id="9f8f9-104">アドインのランタイムを指定します。</span><span class="sxs-lookup"><span data-stu-id="9f8f9-104">Specifies the runtime of your add-in.</span></span> <span data-ttu-id="9f8f9-105">要素の [`<Host>`](host.md) 子。</span><span class="sxs-lookup"><span data-stu-id="9f8f9-105">Child of the [`<Host>`](host.md) element.</span></span>

> [!NOTE]
> <span data-ttu-id="9f8f9-106">Windows Officeで実行する場合、アドインは Internet Explorer 11 ブラウザーを使用します。</span><span class="sxs-lookup"><span data-stu-id="9f8f9-106">When running in Office on Windows, your add-in uses the Internet Explorer 11 browser.</span></span>

<span data-ttu-id="9f8f9-107">Excel では、この要素により、リボン、作業ウィンドウ、およびカスタム関数で同じランタイムを使用できます。</span><span class="sxs-lookup"><span data-stu-id="9f8f9-107">In Excel, this element enables the ribbon, task pane, and custom functions to use the same runtime.</span></span> <span data-ttu-id="9f8f9-108">詳細については、「共有 [JavaScript ランタイムを使用するために Excel](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)アドインを構成する」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="9f8f9-108">For more information, see [Configure your Excel add-in to use a shared JavaScript runtime](../../develop/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

<span data-ttu-id="9f8f9-109">Outlook では、この要素によってイベント ベースのアドインのアクティブ化が有効になります。</span><span class="sxs-lookup"><span data-stu-id="9f8f9-109">In Outlook, this element enables event-based add-in activation.</span></span> <span data-ttu-id="9f8f9-110">詳細については、「イベント ベースのアクティブ [化用に Outlook アドインを構成する」を参照してください](../../outlook/autolaunch.md)。</span><span class="sxs-lookup"><span data-stu-id="9f8f9-110">For more information, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).</span></span>

<span data-ttu-id="9f8f9-111">**アドインの種類:** 作業ウィンドウ、メール</span><span class="sxs-lookup"><span data-stu-id="9f8f9-111">**Add-in type:** Task pane, Mail</span></span>

> [!IMPORTANT]
> <span data-ttu-id="9f8f9-112">**Outlook**: イベント ベースのアクティブ化 [](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md)機能は現在プレビュー中で、Outlook on the web でのみ使用できます。</span><span class="sxs-lookup"><span data-stu-id="9f8f9-112">**Outlook**: The event-based activation feature is currently [in preview](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) and only available in Outlook on the web.</span></span> <span data-ttu-id="9f8f9-113">詳細については、イベント ベースの [アクティブ化機能をプレビューする方法を参照してください](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature)。</span><span class="sxs-lookup"><span data-stu-id="9f8f9-113">For more information, see [How to preview the event-based activation feature](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span></span>

## <a name="syntax"></a><span data-ttu-id="9f8f9-114">構文</span><span class="sxs-lookup"><span data-stu-id="9f8f9-114">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="9f8f9-115">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="9f8f9-115">Contained in</span></span>

[<span data-ttu-id="9f8f9-116">Host</span><span class="sxs-lookup"><span data-stu-id="9f8f9-116">Host</span></span>](host.md)

## <a name="child-elements"></a><span data-ttu-id="9f8f9-117">子要素</span><span class="sxs-lookup"><span data-stu-id="9f8f9-117">Child elements</span></span>

|  <span data-ttu-id="9f8f9-118">要素</span><span class="sxs-lookup"><span data-stu-id="9f8f9-118">Element</span></span> |  <span data-ttu-id="9f8f9-119">必須</span><span class="sxs-lookup"><span data-stu-id="9f8f9-119">Required</span></span>  |  <span data-ttu-id="9f8f9-120">説明</span><span class="sxs-lookup"><span data-stu-id="9f8f9-120">Description</span></span>  |
|:-----|:-----|:-----|
| [<span data-ttu-id="9f8f9-121">ランタイム</span><span class="sxs-lookup"><span data-stu-id="9f8f9-121">Runtime</span></span>](runtime.md) | <span data-ttu-id="9f8f9-122">はい</span><span class="sxs-lookup"><span data-stu-id="9f8f9-122">Yes</span></span> |  <span data-ttu-id="9f8f9-123">アドインのランタイム。</span><span class="sxs-lookup"><span data-stu-id="9f8f9-123">The runtime for your add-in.</span></span> |

## <a name="see-also"></a><span data-ttu-id="9f8f9-124">関連項目</span><span class="sxs-lookup"><span data-stu-id="9f8f9-124">See also</span></span>

- [<span data-ttu-id="9f8f9-125">ランタイム</span><span class="sxs-lookup"><span data-stu-id="9f8f9-125">Runtime</span></span>](runtime.md)
