---
title: マニフェスト ファイル内のランタイム
description: Runtime 要素は、リボン、作業ウィンドウ、カスタム関数など、さまざまなコンポーネントに共有 JavaScript ランタイムを使用するアドインを構成します。
ms.date: 04/08/2021
localization_priority: Normal
ms.openlocfilehash: fa95608d7eff57d68b96ef5b04ec9d33ee63f173
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652245"
---
# <a name="runtime-element"></a><span data-ttu-id="22865-103">Runtime 要素</span><span class="sxs-lookup"><span data-stu-id="22865-103">Runtime element</span></span>

<span data-ttu-id="22865-104">共有 JavaScript ランタイムを使用して、さまざまなコンポーネントすべてが同じランタイムで実行されるアドインを構成します。</span><span class="sxs-lookup"><span data-stu-id="22865-104">Configures your add-in to use a shared JavaScript runtime so that various components all run in the same runtime.</span></span> <span data-ttu-id="22865-105">要素の [`<Runtimes>`](runtimes.md) 子。</span><span class="sxs-lookup"><span data-stu-id="22865-105">Child of the [`<Runtimes>`](runtimes.md) element.</span></span>

<span data-ttu-id="22865-106">**アドインの種類:** 作業ウィンドウ, メール</span><span class="sxs-lookup"><span data-stu-id="22865-106">**Add-in type:** Task pane, Mail</span></span>

[!include[Runtimes support](../../includes/runtimes-note.md)]

## <a name="syntax"></a><span data-ttu-id="22865-107">構文</span><span class="sxs-lookup"><span data-stu-id="22865-107">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="22865-108">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="22865-108">Contained in</span></span>

- [<span data-ttu-id="22865-109">ランタイム</span><span class="sxs-lookup"><span data-stu-id="22865-109">Runtimes</span></span>](runtimes.md)

## <a name="attributes"></a><span data-ttu-id="22865-110">属性</span><span class="sxs-lookup"><span data-stu-id="22865-110">Attributes</span></span>

|  <span data-ttu-id="22865-111">属性</span><span class="sxs-lookup"><span data-stu-id="22865-111">Attribute</span></span>  |  <span data-ttu-id="22865-112">必須</span><span class="sxs-lookup"><span data-stu-id="22865-112">Required</span></span>  |  <span data-ttu-id="22865-113">説明</span><span class="sxs-lookup"><span data-stu-id="22865-113">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="22865-114">**resid**</span><span class="sxs-lookup"><span data-stu-id="22865-114">**resid**</span></span>  |  <span data-ttu-id="22865-115">はい</span><span class="sxs-lookup"><span data-stu-id="22865-115">Yes</span></span>  | <span data-ttu-id="22865-116">アドインの HTML ページの URL の場所を指定します。</span><span class="sxs-lookup"><span data-stu-id="22865-116">Specifies the URL location of the HTML page for your add-in.</span></span> <span data-ttu-id="22865-117">32 文字以内で、要素内の要素の属性と一致 `resid` `id` `Url` する必要 `Resources` があります。</span><span class="sxs-lookup"><span data-stu-id="22865-117">The `resid` can be no more than 32 characters and must match an `id` attribute of a `Url` element in the `Resources` element.</span></span> |
|  <span data-ttu-id="22865-118">**有効期間**</span><span class="sxs-lookup"><span data-stu-id="22865-118">**lifetime**</span></span>  |  <span data-ttu-id="22865-119">いいえ</span><span class="sxs-lookup"><span data-stu-id="22865-119">No</span></span>  | <span data-ttu-id="22865-120">既定値は `lifetime` is `short` であり、指定する必要はない。</span><span class="sxs-lookup"><span data-stu-id="22865-120">The default value for `lifetime` is `short` and doesn't need to be specified.</span></span> <span data-ttu-id="22865-121">Outlook アドインは値のみを使用 `short` します。</span><span class="sxs-lookup"><span data-stu-id="22865-121">Outlook add-ins use only the `short` value.</span></span> <span data-ttu-id="22865-122">Excel アドインで共有ランタイムを使用する場合は、値を明示的にに設定します `long` 。</span><span class="sxs-lookup"><span data-stu-id="22865-122">If you want to use a shared runtime in an Excel add-in, explicitly set the value to `long`.</span></span> |

## <a name="see-also"></a><span data-ttu-id="22865-123">関連項目</span><span class="sxs-lookup"><span data-stu-id="22865-123">See also</span></span>

- [<span data-ttu-id="22865-124">ランタイム</span><span class="sxs-lookup"><span data-stu-id="22865-124">Runtimes</span></span>](runtimes.md)
- [<span data-ttu-id="22865-125">Office アドインを構成して共有 JavaScript ランタイムを使用する</span><span class="sxs-lookup"><span data-stu-id="22865-125">Configure your Office Add-in to use a shared JavaScript runtime</span></span>](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)
- [<span data-ttu-id="22865-126">イベント ベースのライセンス認証用に Outlook アドインを構成する</span><span class="sxs-lookup"><span data-stu-id="22865-126">Configure your Outlook add-in for event-based activation</span></span>](../../outlook/autolaunch.md)
