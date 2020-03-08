---
title: マニフェストファイル内のランタイム (プレビュー)
description: ''
ms.date: 02/21/2020
localization_priority: Normal
ms.openlocfilehash: dd51c5b317700f92ee74c94835e68523371789f8
ms.sourcegitcommit: 153576b1efd0234c6252433e22db213238573534
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/07/2020
ms.locfileid: "42561829"
---
# <a name="runtime-element-preview"></a><span data-ttu-id="3e88a-102">Runtime 要素 (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="3e88a-102">Runtime element (preview)</span></span>

[!include[Running custom functions in browser runtime note](../../includes/excel-shared-runtime-preview-note.md)]

<span data-ttu-id="3e88a-103">[`<Runtimes>`](runtimes.md)要素の子要素。</span><span class="sxs-lookup"><span data-stu-id="3e88a-103">Child element of the [`<Runtimes>`](runtimes.md) element.</span></span> <span data-ttu-id="3e88a-104">この要素は、リボン、作業ウィンドウ、およびカスタム関数がすべて同じランタイムで実行されるように、共有された JavaScript ランタイムを使用するようにアドインを構成します。</span><span class="sxs-lookup"><span data-stu-id="3e88a-104">This element configures your add-in to use a shared JavaScript runtime so that your ribbon, task pane, and custom functions, all run in the same runtime.</span></span> <span data-ttu-id="3e88a-105">詳細については、「[共有 JavaScript ランタイムを使用するように Excel アドインを構成する](../../excel/configure-your-add-in-to-use-a-shared-runtime.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="3e88a-105">For more information, see [Configure your Excel add-in to use a shared JavaScript runtime](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

<span data-ttu-id="3e88a-106">**アドインの種類:** 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="3e88a-106">**Add-in type:** Task pane</span></span>

> [!IMPORTANT]
> <span data-ttu-id="3e88a-107">共有ランタイムは現在プレビュー段階であり、Windows 上の Excel でのみ使用できます。</span><span class="sxs-lookup"><span data-stu-id="3e88a-107">Shared runtime is currently in preview and are only available on Excel on Windows.</span></span> <span data-ttu-id="3e88a-108">プレビュー機能を試すには、 [Office Insider](https://insider.office.com/)に参加する必要があります。</span><span class="sxs-lookup"><span data-stu-id="3e88a-108">To try the preview features, you will need to join [Office Insider](https://insider.office.com/).</span></span>

## <a name="syntax"></a><span data-ttu-id="3e88a-109">構文</span><span class="sxs-lookup"><span data-stu-id="3e88a-109">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="3e88a-110">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="3e88a-110">Contained in</span></span>

- [<span data-ttu-id="3e88a-111">ランタイム</span><span class="sxs-lookup"><span data-stu-id="3e88a-111">Runtimes</span></span>](runtimes.md)

## <a name="attributes"></a><span data-ttu-id="3e88a-112">属性</span><span class="sxs-lookup"><span data-stu-id="3e88a-112">Attributes</span></span>

|  <span data-ttu-id="3e88a-113">属性</span><span class="sxs-lookup"><span data-stu-id="3e88a-113">Attribute</span></span>  |  <span data-ttu-id="3e88a-114">必須</span><span class="sxs-lookup"><span data-stu-id="3e88a-114">Required</span></span>  |  <span data-ttu-id="3e88a-115">説明</span><span class="sxs-lookup"><span data-stu-id="3e88a-115">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="3e88a-116">**lifetime = "long"**</span><span class="sxs-lookup"><span data-stu-id="3e88a-116">**lifetime="long"**</span></span>  |  <span data-ttu-id="3e88a-117">はい</span><span class="sxs-lookup"><span data-stu-id="3e88a-117">Yes</span></span>  | <span data-ttu-id="3e88a-118">Excel アドインの`long`共有ランタイムを常に使用する場合は、必ず指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="3e88a-118">Should always be `long` if you want to use a shared runtime for the Excel add-in.</span></span> |
|  <span data-ttu-id="3e88a-119">**resid**</span><span class="sxs-lookup"><span data-stu-id="3e88a-119">**resid**</span></span>  |  <span data-ttu-id="3e88a-120">はい</span><span class="sxs-lookup"><span data-stu-id="3e88a-120">Yes</span></span>  | <span data-ttu-id="3e88a-121">アドインの HTML ページの URL の場所を指定します。</span><span class="sxs-lookup"><span data-stu-id="3e88a-121">Specifies the URL location of the HTML page for your add-in.</span></span> <span data-ttu-id="3e88a-122">は`resid` 、 `Resources`要素内`id`の`Url`要素の属性と一致している必要があります。</span><span class="sxs-lookup"><span data-stu-id="3e88a-122">The `resid` must match an `id` attribute of a `Url` element in the `Resources` element.</span></span> |

## <a name="see-also"></a><span data-ttu-id="3e88a-123">関連項目</span><span class="sxs-lookup"><span data-stu-id="3e88a-123">See also</span></span>

- [<span data-ttu-id="3e88a-124">ランタイム</span><span class="sxs-lookup"><span data-stu-id="3e88a-124">Runtimes</span></span>](runtimes.md)
