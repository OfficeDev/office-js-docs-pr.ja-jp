---
title: マニフェストファイル内のランタイム
description: Runtime 要素は、アドインが、リボン、作業ウィンドウ、およびカスタム関数に対して共有 JavaScript ランタイムを使用するように構成します。
ms.date: 05/11/2020
localization_priority: Normal
ms.openlocfilehash: c5c7356f9985ca7b5972068629b0587f8916348e
ms.sourcegitcommit: 682d18c9149b1153f9c38d28e2a90384e6a261dc
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/13/2020
ms.locfileid: "44217761"
---
# <a name="runtime-element"></a><span data-ttu-id="c19e1-103">Runtime 要素</span><span class="sxs-lookup"><span data-stu-id="c19e1-103">Runtime element</span></span>

<span data-ttu-id="c19e1-104">要素の子要素 [`<Runtimes>`](runtimes.md) 。</span><span class="sxs-lookup"><span data-stu-id="c19e1-104">Child element of the [`<Runtimes>`](runtimes.md) element.</span></span> <span data-ttu-id="c19e1-105">この要素は、リボン、作業ウィンドウ、およびカスタム関数がすべて同じランタイムで実行されるように、共有された JavaScript ランタイムを使用するようにアドインを構成します。</span><span class="sxs-lookup"><span data-stu-id="c19e1-105">This element configures your add-in to use a shared JavaScript runtime so that your ribbon, task pane, and custom functions, all run in the same runtime.</span></span> <span data-ttu-id="c19e1-106">詳細については、「[共有 JavaScript ランタイムを使用するように Excel アドインを構成する](../../excel/configure-your-add-in-to-use-a-shared-runtime.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="c19e1-106">For more information, see [Configure your Excel add-in to use a shared JavaScript runtime](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

<span data-ttu-id="c19e1-107">**アドインの種類:** 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="c19e1-107">**Add-in type:** Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="c19e1-108">構文</span><span class="sxs-lookup"><span data-stu-id="c19e1-108">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="c19e1-109">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="c19e1-109">Contained in</span></span>

- [<span data-ttu-id="c19e1-110">ランタイム</span><span class="sxs-lookup"><span data-stu-id="c19e1-110">Runtimes</span></span>](runtimes.md)

## <a name="attributes"></a><span data-ttu-id="c19e1-111">属性</span><span class="sxs-lookup"><span data-stu-id="c19e1-111">Attributes</span></span>

|  <span data-ttu-id="c19e1-112">属性</span><span class="sxs-lookup"><span data-stu-id="c19e1-112">Attribute</span></span>  |  <span data-ttu-id="c19e1-113">必須</span><span class="sxs-lookup"><span data-stu-id="c19e1-113">Required</span></span>  |  <span data-ttu-id="c19e1-114">説明</span><span class="sxs-lookup"><span data-stu-id="c19e1-114">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="c19e1-115">**lifetime = "long"**</span><span class="sxs-lookup"><span data-stu-id="c19e1-115">**lifetime="long"**</span></span>  |  <span data-ttu-id="c19e1-116">はい</span><span class="sxs-lookup"><span data-stu-id="c19e1-116">Yes</span></span>  | <span data-ttu-id="c19e1-117">Excel アドインの共有ランタイムを常に使用する場合は、必ず指定する必要があり `long` ます。</span><span class="sxs-lookup"><span data-stu-id="c19e1-117">Should always be `long` if you want to use a shared runtime for the Excel add-in.</span></span> |
|  <span data-ttu-id="c19e1-118">**resid**</span><span class="sxs-lookup"><span data-stu-id="c19e1-118">**resid**</span></span>  |  <span data-ttu-id="c19e1-119">はい</span><span class="sxs-lookup"><span data-stu-id="c19e1-119">Yes</span></span>  | <span data-ttu-id="c19e1-120">アドインの HTML ページの URL の場所を指定します。</span><span class="sxs-lookup"><span data-stu-id="c19e1-120">Specifies the URL location of the HTML page for your add-in.</span></span> <span data-ttu-id="c19e1-121">は、 `resid` `id` 要素内の要素の属性と一致している必要があり `Url` `Resources` ます。</span><span class="sxs-lookup"><span data-stu-id="c19e1-121">The `resid` must match an `id` attribute of a `Url` element in the `Resources` element.</span></span> |

## <a name="see-also"></a><span data-ttu-id="c19e1-122">関連項目</span><span class="sxs-lookup"><span data-stu-id="c19e1-122">See also</span></span>

- [<span data-ttu-id="c19e1-123">ランタイム</span><span class="sxs-lookup"><span data-stu-id="c19e1-123">Runtimes</span></span>](runtimes.md)
