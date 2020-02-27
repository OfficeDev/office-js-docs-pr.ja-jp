---
title: マニフェストファイル内のランタイム (プレビュー)
description: ''
ms.date: 02/21/2020
localization_priority: Normal
ms.openlocfilehash: 26702896604f9ecf4c69296e5110efe5cdf4218b
ms.sourcegitcommit: dd6d00202f6466c27418247dad7bd136555a6036
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/26/2020
ms.locfileid: "42283885"
---
# <a name="runtime-element-preview"></a><span data-ttu-id="179ab-102">Runtime 要素 (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="179ab-102">Runtime element (preview)</span></span>

[!include[Running custom functions in browser runtime note](../../includes/excel-shared-runtime-preview-note.md)]

<span data-ttu-id="179ab-103">[`<Runtimes>`](runtimes.md)要素の子要素。</span><span class="sxs-lookup"><span data-stu-id="179ab-103">Child element of the [`<Runtimes>`](runtimes.md) element.</span></span> <span data-ttu-id="179ab-104">この要素は、リボン、作業ウィンドウ、およびカスタム関数がすべて同じランタイムで実行されるように、共有された JavaScript ランタイムを使用するようにアドインを構成します。</span><span class="sxs-lookup"><span data-stu-id="179ab-104">This element configures your add-in to use a shared JavaScript runtime so that your ribbon, task pane, and custom functions, all run in the same runtime.</span></span> <span data-ttu-id="179ab-105">詳細については、「[共有 JavaScript ランタイムを使用するように Excel アドインを構成する](../../excel/configure-your-add-in-to-use-a-shared-runtime.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="179ab-105">For more information, see [Configure your Excel add-in to use a shared JavaScript runtime](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

<span data-ttu-id="179ab-106">**アドインの種類:** 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="179ab-106">**Add-in type:** Task pane</span></span>

> [!IMPORTANT]
<span data-ttu-id="179ab-107"><<<<<<< ヘッド共有ランタイムは現在プレビュー段階であり、Windows 上の Excel でのみ使用できます。</span><span class="sxs-lookup"><span data-stu-id="179ab-107"><<<<<<< HEAD Shared runtime is currently in preview and are only available on Excel on Windows.</span></span> <span data-ttu-id="179ab-108">プレビュー機能を試すには、 [Office Insider](https://insider.office.com/)に参加する必要があります。</span><span class="sxs-lookup"><span data-stu-id="179ab-108">To try the preview features, you will need to join [Office Insider](https://insider.office.com/).</span></span>

## <a name="syntax"></a><span data-ttu-id="179ab-109">構文</span><span class="sxs-lookup"><span data-stu-id="179ab-109">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="179ab-110">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="179ab-110">Contained in</span></span>

- [<span data-ttu-id="179ab-111">ランタイム</span><span class="sxs-lookup"><span data-stu-id="179ab-111">Runtimes</span></span>](runtimes.md)

## <a name="attributes"></a><span data-ttu-id="179ab-112">属性</span><span class="sxs-lookup"><span data-stu-id="179ab-112">Attributes</span></span>

|  <span data-ttu-id="179ab-113">属性</span><span class="sxs-lookup"><span data-stu-id="179ab-113">Attribute</span></span>  |  <span data-ttu-id="179ab-114">必須</span><span class="sxs-lookup"><span data-stu-id="179ab-114">Required</span></span>  |  <span data-ttu-id="179ab-115">説明</span><span class="sxs-lookup"><span data-stu-id="179ab-115">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="179ab-116">**lifetime = "long"**</span><span class="sxs-lookup"><span data-stu-id="179ab-116">**lifetime="long"**</span></span>  |  <span data-ttu-id="179ab-117">はい</span><span class="sxs-lookup"><span data-stu-id="179ab-117">Yes</span></span>  | <span data-ttu-id="179ab-118">Excel アドインの`long`共有ランタイムを常に使用する場合は、必ず指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="179ab-118">Should always be `long` if you want to use a shared runtime for the Excel add-in.</span></span> |
|  <span data-ttu-id="179ab-119">**resid**</span><span class="sxs-lookup"><span data-stu-id="179ab-119">**resid**</span></span>  |  <span data-ttu-id="179ab-120">はい</span><span class="sxs-lookup"><span data-stu-id="179ab-120">Yes</span></span>  | <span data-ttu-id="179ab-121">アドインの HTML ページの URL の場所を指定します。</span><span class="sxs-lookup"><span data-stu-id="179ab-121">Specifies the URL location of the HTML page for your add-in.</span></span> <span data-ttu-id="179ab-122">は`resid` 、 `Resources`要素内`id`の`Url`要素の属性と一致している必要があります。</span><span class="sxs-lookup"><span data-stu-id="179ab-122">The `resid` must match an `id` attribute of a `Url` element in the `Resources` element.</span></span> |

## <a name="see-also"></a><span data-ttu-id="179ab-123">関連項目</span><span class="sxs-lookup"><span data-stu-id="179ab-123">See also</span></span>

- [<span data-ttu-id="179ab-124">ランタイム</span><span class="sxs-lookup"><span data-stu-id="179ab-124">Runtimes</span></span>](runtimes.md)
