---
title: マニフェストファイル内のランタイム (プレビュー)
description: Runtime 要素は、アドインが、リボン、作業ウィンドウ、およびカスタム関数に対して共有 JavaScript ランタイムを使用するように構成します。
ms.date: 02/21/2020
localization_priority: Normal
ms.openlocfilehash: 6237f64fec47ed22b0105bf74c8eb7e2b7c38afe
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42717930"
---
# <a name="runtime-element-preview"></a><span data-ttu-id="05b89-103">Runtime 要素 (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="05b89-103">Runtime element (preview)</span></span>

[!include[Running custom functions in browser runtime note](../../includes/excel-shared-runtime-preview-note.md)]

<span data-ttu-id="05b89-104">[`<Runtimes>`](runtimes.md)要素の子要素。</span><span class="sxs-lookup"><span data-stu-id="05b89-104">Child element of the [`<Runtimes>`](runtimes.md) element.</span></span> <span data-ttu-id="05b89-105">この要素は、リボン、作業ウィンドウ、およびカスタム関数がすべて同じランタイムで実行されるように、共有された JavaScript ランタイムを使用するようにアドインを構成します。</span><span class="sxs-lookup"><span data-stu-id="05b89-105">This element configures your add-in to use a shared JavaScript runtime so that your ribbon, task pane, and custom functions, all run in the same runtime.</span></span> <span data-ttu-id="05b89-106">詳細については、「[共有 JavaScript ランタイムを使用するように Excel アドインを構成する](../../excel/configure-your-add-in-to-use-a-shared-runtime.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="05b89-106">For more information, see [Configure your Excel add-in to use a shared JavaScript runtime](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

<span data-ttu-id="05b89-107">**アドインの種類:** 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="05b89-107">**Add-in type:** Task pane</span></span>

> [!IMPORTANT]
> <span data-ttu-id="05b89-108">共有ランタイムは現在プレビュー段階であり、Windows 上の Excel でのみ使用できます。</span><span class="sxs-lookup"><span data-stu-id="05b89-108">Shared runtime is currently in preview and are only available on Excel on Windows.</span></span> <span data-ttu-id="05b89-109">プレビュー機能を試すには、 [Office Insider](https://insider.office.com/)に参加する必要があります。</span><span class="sxs-lookup"><span data-stu-id="05b89-109">To try the preview features, you will need to join [Office Insider](https://insider.office.com/).</span></span>

## <a name="syntax"></a><span data-ttu-id="05b89-110">構文</span><span class="sxs-lookup"><span data-stu-id="05b89-110">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="05b89-111">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="05b89-111">Contained in</span></span>

- [<span data-ttu-id="05b89-112">ランタイム</span><span class="sxs-lookup"><span data-stu-id="05b89-112">Runtimes</span></span>](runtimes.md)

## <a name="attributes"></a><span data-ttu-id="05b89-113">属性</span><span class="sxs-lookup"><span data-stu-id="05b89-113">Attributes</span></span>

|  <span data-ttu-id="05b89-114">属性</span><span class="sxs-lookup"><span data-stu-id="05b89-114">Attribute</span></span>  |  <span data-ttu-id="05b89-115">必須</span><span class="sxs-lookup"><span data-stu-id="05b89-115">Required</span></span>  |  <span data-ttu-id="05b89-116">説明</span><span class="sxs-lookup"><span data-stu-id="05b89-116">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="05b89-117">**lifetime = "long"**</span><span class="sxs-lookup"><span data-stu-id="05b89-117">**lifetime="long"**</span></span>  |  <span data-ttu-id="05b89-118">はい</span><span class="sxs-lookup"><span data-stu-id="05b89-118">Yes</span></span>  | <span data-ttu-id="05b89-119">Excel アドインの`long`共有ランタイムを常に使用する場合は、必ず指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="05b89-119">Should always be `long` if you want to use a shared runtime for the Excel add-in.</span></span> |
|  <span data-ttu-id="05b89-120">**resid**</span><span class="sxs-lookup"><span data-stu-id="05b89-120">**resid**</span></span>  |  <span data-ttu-id="05b89-121">はい</span><span class="sxs-lookup"><span data-stu-id="05b89-121">Yes</span></span>  | <span data-ttu-id="05b89-122">アドインの HTML ページの URL の場所を指定します。</span><span class="sxs-lookup"><span data-stu-id="05b89-122">Specifies the URL location of the HTML page for your add-in.</span></span> <span data-ttu-id="05b89-123">は`resid` 、 `Resources`要素内`id`の`Url`要素の属性と一致している必要があります。</span><span class="sxs-lookup"><span data-stu-id="05b89-123">The `resid` must match an `id` attribute of a `Url` element in the `Resources` element.</span></span> |

## <a name="see-also"></a><span data-ttu-id="05b89-124">関連項目</span><span class="sxs-lookup"><span data-stu-id="05b89-124">See also</span></span>

- [<span data-ttu-id="05b89-125">ランタイム</span><span class="sxs-lookup"><span data-stu-id="05b89-125">Runtimes</span></span>](runtimes.md)
