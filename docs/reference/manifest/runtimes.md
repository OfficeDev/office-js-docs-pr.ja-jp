---
title: マニフェストファイル内のランタイム (プレビュー)
description: Runtime 要素は、アドインのランタイムを指定します。
ms.date: 02/21/2020
localization_priority: Normal
ms.openlocfilehash: 5797aa78ae3667461de48de481ff44f14c307ced
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720422"
---
# <a name="runtimes-element-preview"></a><span data-ttu-id="4fe0f-103">ランタイム要素 (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="4fe0f-103">Runtimes element (preview)</span></span>

[!include[Running custom functions in browser runtime note](../../includes/excel-shared-runtime-preview-note.md)]

<span data-ttu-id="4fe0f-104">アドインのランタイムを指定し、カスタム関数、リボンボタン、および作業ウィンドウを使用して同じ JavaScript ランタイムを使用できるようにします。</span><span class="sxs-lookup"><span data-stu-id="4fe0f-104">Specifies the runtime of your add-in and enables custom functions, ribbon buttons, and the task pane to use the same JavaScript runtime.</span></span> <span data-ttu-id="4fe0f-105">マニフェストファイル内`<Host>`の要素の子。</span><span class="sxs-lookup"><span data-stu-id="4fe0f-105">Child of the `<Host>` element in your manifest file.</span></span> <span data-ttu-id="4fe0f-106">詳細については、「[共有 JavaScript ランタイムを使用するように Excel アドインを構成する](../../excel/configure-your-add-in-to-use-a-shared-runtime.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="4fe0f-106">For more information, see [Configure your Excel add-in to use a shared JavaScript runtime](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

<span data-ttu-id="4fe0f-107">**アドインの種類:** 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4fe0f-107">**Add-in type:** Task pane</span></span>

> [!IMPORTANT]
> <span data-ttu-id="4fe0f-108">共有ランタイムは現在プレビュー段階であり、Windows 上の Excel でのみ使用できます。</span><span class="sxs-lookup"><span data-stu-id="4fe0f-108">Shared runtime is currently in preview and are only available on Excel on Windows.</span></span> <span data-ttu-id="4fe0f-109">プレビュー機能を試すには、 [Office Insider](https://insider.office.com/)に参加する必要があります。</span><span class="sxs-lookup"><span data-stu-id="4fe0f-109">To try the preview features, you will need to join [Office Insider](https://insider.office.com/).</span></span>

## <a name="syntax"></a><span data-ttu-id="4fe0f-110">構文</span><span class="sxs-lookup"><span data-stu-id="4fe0f-110">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="4fe0f-111">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="4fe0f-111">Contained in</span></span> 
[<span data-ttu-id="4fe0f-112">Host</span><span class="sxs-lookup"><span data-stu-id="4fe0f-112">Host</span></span>](./host.md)

## <a name="child-elements"></a><span data-ttu-id="4fe0f-113">子要素</span><span class="sxs-lookup"><span data-stu-id="4fe0f-113">Child elements</span></span>

|  <span data-ttu-id="4fe0f-114">要素</span><span class="sxs-lookup"><span data-stu-id="4fe0f-114">Element</span></span> |  <span data-ttu-id="4fe0f-115">必須</span><span class="sxs-lookup"><span data-stu-id="4fe0f-115">Required</span></span>  |  <span data-ttu-id="4fe0f-116">説明</span><span class="sxs-lookup"><span data-stu-id="4fe0f-116">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="4fe0f-117">**ランタイム**</span><span class="sxs-lookup"><span data-stu-id="4fe0f-117">**Runtime**</span></span>     | <span data-ttu-id="4fe0f-118">はい</span><span class="sxs-lookup"><span data-stu-id="4fe0f-118">Yes</span></span> |  <span data-ttu-id="4fe0f-119">アドインのランタイム。</span><span class="sxs-lookup"><span data-stu-id="4fe0f-119">The runtime for your add-in.</span></span>

## <a name="see-also"></a><span data-ttu-id="4fe0f-120">関連項目</span><span class="sxs-lookup"><span data-stu-id="4fe0f-120">See also</span></span>

- [<span data-ttu-id="4fe0f-121">ランタイム</span><span class="sxs-lookup"><span data-stu-id="4fe0f-121">Runtime</span></span>](runtime.md)
