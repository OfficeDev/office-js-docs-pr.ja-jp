---
title: マニフェストファイル内のランタイム
description: ランタイム要素は、アドインのランタイムを指定します。
ms.date: 06/01/2020
localization_priority: Normal
ms.openlocfilehash: 082491befc6b9dbdc474b0e40f9defd90a4ef75f
ms.sourcegitcommit: 472b81642e9eb5fb2a55cd98a7b0826d37eb7f73
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/17/2020
ms.locfileid: "45159361"
---
# <a name="runtimes-element"></a><span data-ttu-id="4310d-103">ランタイム要素</span><span class="sxs-lookup"><span data-stu-id="4310d-103">Runtimes element</span></span>

<span data-ttu-id="4310d-104">アドインの実行時のランタイムを指定します。</span><span class="sxs-lookup"><span data-stu-id="4310d-104">Specifies the runtime of your add-in.</span></span> <span data-ttu-id="4310d-105">要素の子 [`<Host>`](host.md) 。</span><span class="sxs-lookup"><span data-stu-id="4310d-105">Child of the [`<Host>`](host.md) element.</span></span>

> [!NOTE]
> <span data-ttu-id="4310d-106">Windows で Office を実行している場合、アドインは Internet Explorer 11 ブラウザーを使用します。</span><span class="sxs-lookup"><span data-stu-id="4310d-106">When running in Office on Windows, your add-in uses the Internet Explorer 11 browser.</span></span>

<span data-ttu-id="4310d-107">Excel では、この要素を使用すると、リボン、作業ウィンドウ、およびカスタム関数が同じランタイムを使用できるようになります。</span><span class="sxs-lookup"><span data-stu-id="4310d-107">In Excel, this element enables the ribbon, task pane, and custom functions to use the same runtime.</span></span> <span data-ttu-id="4310d-108">詳細については、「[共有 JavaScript ランタイムを使用するように Excel アドインを構成する](../../excel/configure-your-add-in-to-use-a-shared-runtime.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="4310d-108">For more information, see [Configure your Excel add-in to use a shared JavaScript runtime](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

<span data-ttu-id="4310d-109">Outlook では、この要素はイベントベースのアドインのアクティブ化を有効にします。</span><span class="sxs-lookup"><span data-stu-id="4310d-109">In Outlook, this element enables event-based add-in activation.</span></span> <span data-ttu-id="4310d-110">詳細については、「[イベントベースのライセンス認証用に Outlook アドインを構成する](../../outlook/autolaunch.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="4310d-110">For more information, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).</span></span>

<span data-ttu-id="4310d-111">**アドインの種類:** 作業ウィンドウ、メール</span><span class="sxs-lookup"><span data-stu-id="4310d-111">**Add-in type:** Task pane, Mail</span></span>

> [!IMPORTANT]
> <span data-ttu-id="4310d-112">**Outlook**: イベントベースのライセンス認証機能は現在[プレビュー段階で](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md)あり、web 上の Outlook でのみ使用できます。</span><span class="sxs-lookup"><span data-stu-id="4310d-112">**Outlook**: The event-based activation feature is currently [in preview](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) and only available in Outlook on the web.</span></span> <span data-ttu-id="4310d-113">詳細については、「[イベントベースのライセンス認証機能をプレビューする方法](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="4310d-113">For more information, see [How to preview the event-based activation feature](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span></span>

## <a name="syntax"></a><span data-ttu-id="4310d-114">構文</span><span class="sxs-lookup"><span data-stu-id="4310d-114">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="4310d-115">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="4310d-115">Contained in</span></span>

[<span data-ttu-id="4310d-116">Host</span><span class="sxs-lookup"><span data-stu-id="4310d-116">Host</span></span>](host.md)

## <a name="child-elements"></a><span data-ttu-id="4310d-117">子要素</span><span class="sxs-lookup"><span data-stu-id="4310d-117">Child elements</span></span>

|  <span data-ttu-id="4310d-118">要素</span><span class="sxs-lookup"><span data-stu-id="4310d-118">Element</span></span> |  <span data-ttu-id="4310d-119">必須</span><span class="sxs-lookup"><span data-stu-id="4310d-119">Required</span></span>  |  <span data-ttu-id="4310d-120">説明</span><span class="sxs-lookup"><span data-stu-id="4310d-120">Description</span></span>  |
|:-----|:-----|:-----|
| [<span data-ttu-id="4310d-121">ランタイム</span><span class="sxs-lookup"><span data-stu-id="4310d-121">Runtime</span></span>](runtime.md) | <span data-ttu-id="4310d-122">はい</span><span class="sxs-lookup"><span data-stu-id="4310d-122">Yes</span></span> |  <span data-ttu-id="4310d-123">アドインのランタイム。</span><span class="sxs-lookup"><span data-stu-id="4310d-123">The runtime for your add-in.</span></span> |

## <a name="see-also"></a><span data-ttu-id="4310d-124">関連項目</span><span class="sxs-lookup"><span data-stu-id="4310d-124">See also</span></span>

- [<span data-ttu-id="4310d-125">ランタイム</span><span class="sxs-lookup"><span data-stu-id="4310d-125">Runtime</span></span>](runtime.md)
