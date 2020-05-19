---
title: マニフェストファイル内のランタイム
description: ランタイム要素は、アドインのランタイムを指定します。
ms.date: 05/18/2020
localization_priority: Normal
ms.openlocfilehash: 22156a171ca2f423024efb1b3d2a6fdae07dfef6
ms.sourcegitcommit: f62d9630de69c5c070e3d4048205f5cc654db7e4
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/18/2020
ms.locfileid: "44278365"
---
# <a name="runtimes-element"></a><span data-ttu-id="73195-103">ランタイム要素</span><span class="sxs-lookup"><span data-stu-id="73195-103">Runtimes element</span></span>

<span data-ttu-id="73195-104">アドインの実行時のランタイムを指定します。</span><span class="sxs-lookup"><span data-stu-id="73195-104">Specifies the runtime of your add-in.</span></span> <span data-ttu-id="73195-105">要素の子 [`<Host>`](host.md) 。</span><span class="sxs-lookup"><span data-stu-id="73195-105">Child of the [`<Host>`](host.md) element.</span></span>

<span data-ttu-id="73195-106">Excel では、この要素を使用すると、リボン、作業ウィンドウ、およびカスタム関数が同じランタイムを使用できるようになります。</span><span class="sxs-lookup"><span data-stu-id="73195-106">In Excel, this element enables the ribbon, task pane, and custom functions to use the same runtime.</span></span> <span data-ttu-id="73195-107">詳細については、「[共有 JavaScript ランタイムを使用するように Excel アドインを構成する](../../excel/configure-your-add-in-to-use-a-shared-runtime.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="73195-107">For more information, see [Configure your Excel add-in to use a shared JavaScript runtime](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

<span data-ttu-id="73195-108">Outlook では、この要素はイベントベースのアドインのアクティブ化を有効にします。</span><span class="sxs-lookup"><span data-stu-id="73195-108">In Outlook, this element enables event-based add-in activation.</span></span> <span data-ttu-id="73195-109">詳細については、「[イベントベースのライセンス認証用に Outlook アドインを構成する](../../outlook/autolaunch.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="73195-109">For more information, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).</span></span>

<span data-ttu-id="73195-110">**アドインの種類:** 作業ウィンドウ、メール</span><span class="sxs-lookup"><span data-stu-id="73195-110">**Add-in type:** Task pane, Mail</span></span>

> [!IMPORTANT]
> <span data-ttu-id="73195-111">**Excel**: 共有ランタイムは現在プレビュー段階であり、Windows 上の Excel でのみ使用できます。</span><span class="sxs-lookup"><span data-stu-id="73195-111">**Excel**: Shared runtime is currently in preview and only available in Excel on Windows.</span></span> <span data-ttu-id="73195-112">プレビュー機能を試すには、 [Office Insider](https://insider.office.com/)に参加する必要があります。</span><span class="sxs-lookup"><span data-stu-id="73195-112">To try the preview features, you will need to join [Office Insider](https://insider.office.com/).</span></span>
>
> <span data-ttu-id="73195-113">**Outlook**: イベントベースのライセンス認証機能は現在[プレビュー段階で](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md)あり、web 上の Outlook でのみ使用できます。</span><span class="sxs-lookup"><span data-stu-id="73195-113">**Outlook**: The event-based activation feature is currently [in preview](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) and only available in Outlook on the web.</span></span> <span data-ttu-id="73195-114">詳細については、「[イベントベースのライセンス認証機能をプレビューする方法](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="73195-114">For more information, see [How to preview the event-based activation feature](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span></span>

## <a name="syntax"></a><span data-ttu-id="73195-115">構文</span><span class="sxs-lookup"><span data-stu-id="73195-115">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="73195-116">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="73195-116">Contained in</span></span>

[<span data-ttu-id="73195-117">Host</span><span class="sxs-lookup"><span data-stu-id="73195-117">Host</span></span>](host.md)

## <a name="child-elements"></a><span data-ttu-id="73195-118">子要素</span><span class="sxs-lookup"><span data-stu-id="73195-118">Child elements</span></span>

|  <span data-ttu-id="73195-119">要素</span><span class="sxs-lookup"><span data-stu-id="73195-119">Element</span></span> |  <span data-ttu-id="73195-120">必須</span><span class="sxs-lookup"><span data-stu-id="73195-120">Required</span></span>  |  <span data-ttu-id="73195-121">説明</span><span class="sxs-lookup"><span data-stu-id="73195-121">Description</span></span>  |
|:-----|:-----|:-----|
| [<span data-ttu-id="73195-122">ランタイム</span><span class="sxs-lookup"><span data-stu-id="73195-122">Runtime</span></span>](runtime.md) | <span data-ttu-id="73195-123">はい</span><span class="sxs-lookup"><span data-stu-id="73195-123">Yes</span></span> |  <span data-ttu-id="73195-124">アドインのランタイム。</span><span class="sxs-lookup"><span data-stu-id="73195-124">The runtime for your add-in.</span></span> |

## <a name="see-also"></a><span data-ttu-id="73195-125">関連項目</span><span class="sxs-lookup"><span data-stu-id="73195-125">See also</span></span>

- [<span data-ttu-id="73195-126">ランタイム</span><span class="sxs-lookup"><span data-stu-id="73195-126">Runtime</span></span>](runtime.md)
