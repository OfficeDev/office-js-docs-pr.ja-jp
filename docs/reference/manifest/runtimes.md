---
title: マニフェストファイル内のランタイム
description: ランタイム要素は、アドインのランタイムを指定します。
ms.date: 06/01/2020
localization_priority: Normal
ms.openlocfilehash: 95549d88df24a7d7c54cf27c92c15693491bdf29
ms.sourcegitcommit: 9229102c16a1864e3a8724aaf9b0dc68b1428094
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/03/2020
ms.locfileid: "44520340"
---
# <a name="runtimes-element"></a><span data-ttu-id="2b0fd-103">ランタイム要素</span><span class="sxs-lookup"><span data-stu-id="2b0fd-103">Runtimes element</span></span>

<span data-ttu-id="2b0fd-104">アドインの実行時のランタイムを指定します。</span><span class="sxs-lookup"><span data-stu-id="2b0fd-104">Specifies the runtime of your add-in.</span></span> <span data-ttu-id="2b0fd-105">要素の子 [`<Host>`](host.md) 。</span><span class="sxs-lookup"><span data-stu-id="2b0fd-105">Child of the [`<Host>`](host.md) element.</span></span>

> [!NOTE]
> <span data-ttu-id="2b0fd-106">Windows で Office を実行している場合、アドインは Internet Explorer 11 ブラウザーを使用します。</span><span class="sxs-lookup"><span data-stu-id="2b0fd-106">When running in Office on Windows, your add-in uses the Internet Explorer 11 browser.</span></span>

<span data-ttu-id="2b0fd-107">Excel では、この要素を使用すると、リボン、作業ウィンドウ、およびカスタム関数が同じランタイムを使用できるようになります。</span><span class="sxs-lookup"><span data-stu-id="2b0fd-107">In Excel, this element enables the ribbon, task pane, and custom functions to use the same runtime.</span></span> <span data-ttu-id="2b0fd-108">詳細については、「[共有 JavaScript ランタイムを使用するように Excel アドインを構成する](../../excel/configure-your-add-in-to-use-a-shared-runtime.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="2b0fd-108">For more information, see [Configure your Excel add-in to use a shared JavaScript runtime](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

<span data-ttu-id="2b0fd-109">Outlook では、この要素はイベントベースのアドインのアクティブ化を有効にします。</span><span class="sxs-lookup"><span data-stu-id="2b0fd-109">In Outlook, this element enables event-based add-in activation.</span></span> <span data-ttu-id="2b0fd-110">詳細については、「[イベントベースのライセンス認証用に Outlook アドインを構成する](../../outlook/autolaunch.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="2b0fd-110">For more information, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).</span></span>

<span data-ttu-id="2b0fd-111">**アドインの種類:** 作業ウィンドウ、メール</span><span class="sxs-lookup"><span data-stu-id="2b0fd-111">**Add-in type:** Task pane, Mail</span></span>

> [!IMPORTANT]
> <span data-ttu-id="2b0fd-112">**Excel**: 共有ランタイムは、現在 Windows 上の Excel でのみ使用できます。</span><span class="sxs-lookup"><span data-stu-id="2b0fd-112">**Excel**: Shared runtime is currently only available in Excel on Windows.</span></span>
>
> <span data-ttu-id="2b0fd-113">**Outlook**: イベントベースのライセンス認証機能は現在[プレビュー段階で](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md)あり、web 上の Outlook でのみ使用できます。</span><span class="sxs-lookup"><span data-stu-id="2b0fd-113">**Outlook**: The event-based activation feature is currently [in preview](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) and only available in Outlook on the web.</span></span> <span data-ttu-id="2b0fd-114">詳細については、「[イベントベースのライセンス認証機能をプレビューする方法](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="2b0fd-114">For more information, see [How to preview the event-based activation feature](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span></span>

## <a name="syntax"></a><span data-ttu-id="2b0fd-115">構文</span><span class="sxs-lookup"><span data-stu-id="2b0fd-115">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="2b0fd-116">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="2b0fd-116">Contained in</span></span>

[<span data-ttu-id="2b0fd-117">Host</span><span class="sxs-lookup"><span data-stu-id="2b0fd-117">Host</span></span>](host.md)

## <a name="child-elements"></a><span data-ttu-id="2b0fd-118">子要素</span><span class="sxs-lookup"><span data-stu-id="2b0fd-118">Child elements</span></span>

|  <span data-ttu-id="2b0fd-119">要素</span><span class="sxs-lookup"><span data-stu-id="2b0fd-119">Element</span></span> |  <span data-ttu-id="2b0fd-120">必須</span><span class="sxs-lookup"><span data-stu-id="2b0fd-120">Required</span></span>  |  <span data-ttu-id="2b0fd-121">説明</span><span class="sxs-lookup"><span data-stu-id="2b0fd-121">Description</span></span>  |
|:-----|:-----|:-----|
| [<span data-ttu-id="2b0fd-122">ランタイム</span><span class="sxs-lookup"><span data-stu-id="2b0fd-122">Runtime</span></span>](runtime.md) | <span data-ttu-id="2b0fd-123">はい</span><span class="sxs-lookup"><span data-stu-id="2b0fd-123">Yes</span></span> |  <span data-ttu-id="2b0fd-124">アドインのランタイム。</span><span class="sxs-lookup"><span data-stu-id="2b0fd-124">The runtime for your add-in.</span></span> |

## <a name="see-also"></a><span data-ttu-id="2b0fd-125">関連項目</span><span class="sxs-lookup"><span data-stu-id="2b0fd-125">See also</span></span>

- [<span data-ttu-id="2b0fd-126">ランタイム</span><span class="sxs-lookup"><span data-stu-id="2b0fd-126">Runtime</span></span>](runtime.md)
