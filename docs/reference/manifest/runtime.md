---
title: マニフェスト ファイル内のランタイム
description: Runtime 要素は、リボン、作業ウィンドウ、カスタム関数など、さまざまなコンポーネントに共有 JavaScript ランタイムを使用するアドインを構成します。
ms.date: 05/19/2021
localization_priority: Normal
ms.openlocfilehash: cd09abe31ff57eac629c6c61c873c5c886f73f9c
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/21/2021
ms.locfileid: "52590914"
---
# <a name="runtime-element"></a><span data-ttu-id="0e61c-103">Runtime 要素</span><span class="sxs-lookup"><span data-stu-id="0e61c-103">Runtime element</span></span>

<span data-ttu-id="0e61c-104">共有 JavaScript ランタイムを使用して、さまざまなコンポーネントすべてが同じランタイムで実行されるアドインを構成します。</span><span class="sxs-lookup"><span data-stu-id="0e61c-104">Configures your add-in to use a shared JavaScript runtime so that various components all run in the same runtime.</span></span> <span data-ttu-id="0e61c-105">要素の [`<Runtimes>`](runtimes.md) 子。</span><span class="sxs-lookup"><span data-stu-id="0e61c-105">Child of the [`<Runtimes>`](runtimes.md) element.</span></span>

<span data-ttu-id="0e61c-106">**アドインの種類:** 作業ウィンドウ, メール</span><span class="sxs-lookup"><span data-stu-id="0e61c-106">**Add-in type:** Task pane, Mail</span></span>

[!include[Runtimes support](../../includes/runtimes-note.md)]

## <a name="syntax"></a><span data-ttu-id="0e61c-107">構文</span><span class="sxs-lookup"><span data-stu-id="0e61c-107">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="0e61c-108">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="0e61c-108">Contained in</span></span>

- [<span data-ttu-id="0e61c-109">ランタイム</span><span class="sxs-lookup"><span data-stu-id="0e61c-109">Runtimes</span></span>](runtimes.md)

## <a name="child-elements"></a><span data-ttu-id="0e61c-110">子要素</span><span class="sxs-lookup"><span data-stu-id="0e61c-110">Child elements</span></span>

|  <span data-ttu-id="0e61c-111">要素</span><span class="sxs-lookup"><span data-stu-id="0e61c-111">Element</span></span> |  <span data-ttu-id="0e61c-112">必須</span><span class="sxs-lookup"><span data-stu-id="0e61c-112">Required</span></span>  |  <span data-ttu-id="0e61c-113">説明</span><span class="sxs-lookup"><span data-stu-id="0e61c-113">Description</span></span>  |
|:-----|:-----|:-----|
| [<span data-ttu-id="0e61c-114">Override</span><span class="sxs-lookup"><span data-stu-id="0e61c-114">Override</span></span>](override.md) | <span data-ttu-id="0e61c-115">なし</span><span class="sxs-lookup"><span data-stu-id="0e61c-115">No</span></span> | <span data-ttu-id="0e61c-116">**Outlook**: LaunchEvent 拡張ポイント ハンドラーにデスクトップで必要Outlook JavaScript ファイルの URL [の場所を指定](../../reference/manifest/extensionpoint.md#launchevent)します。</span><span class="sxs-lookup"><span data-stu-id="0e61c-116">**Outlook**: Specifies the URL location of the JavaScript file that Outlook Desktop requires for [LaunchEvent extension point](../../reference/manifest/extensionpoint.md#launchevent) handlers.</span></span> <span data-ttu-id="0e61c-117">**重要**: 現時点では、定義できる要素は 1 つで `<Override>` 、型である必要があります `javascript` 。</span><span class="sxs-lookup"><span data-stu-id="0e61c-117">**Important**: At present, you can only define one `<Override>` element and it must be of type `javascript`.</span></span>|

## <a name="attributes"></a><span data-ttu-id="0e61c-118">属性</span><span class="sxs-lookup"><span data-stu-id="0e61c-118">Attributes</span></span>

|  <span data-ttu-id="0e61c-119">属性</span><span class="sxs-lookup"><span data-stu-id="0e61c-119">Attribute</span></span>  |  <span data-ttu-id="0e61c-120">必須</span><span class="sxs-lookup"><span data-stu-id="0e61c-120">Required</span></span>  |  <span data-ttu-id="0e61c-121">説明</span><span class="sxs-lookup"><span data-stu-id="0e61c-121">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="0e61c-122">**resid**</span><span class="sxs-lookup"><span data-stu-id="0e61c-122">**resid**</span></span>  |  <span data-ttu-id="0e61c-123">必要</span><span class="sxs-lookup"><span data-stu-id="0e61c-123">Yes</span></span>  | <span data-ttu-id="0e61c-124">アドインの HTML ページの URL の場所を指定します。</span><span class="sxs-lookup"><span data-stu-id="0e61c-124">Specifies the URL location of the HTML page for your add-in.</span></span> <span data-ttu-id="0e61c-125">32 文字以内で、要素内の要素の属性と一致 `resid` `id` `Url` する必要 `Resources` があります。</span><span class="sxs-lookup"><span data-stu-id="0e61c-125">The `resid` can be no more than 32 characters and must match an `id` attribute of a `Url` element in the `Resources` element.</span></span> |
|  <span data-ttu-id="0e61c-126">**有効期間**</span><span class="sxs-lookup"><span data-stu-id="0e61c-126">**lifetime**</span></span>  |  <span data-ttu-id="0e61c-127">いいえ</span><span class="sxs-lookup"><span data-stu-id="0e61c-127">No</span></span>  | <span data-ttu-id="0e61c-128">既定値は `lifetime` is `short` であり、指定する必要はない。</span><span class="sxs-lookup"><span data-stu-id="0e61c-128">The default value for `lifetime` is `short` and doesn't need to be specified.</span></span> <span data-ttu-id="0e61c-129">Outlookは値のみを使用 `short` します。</span><span class="sxs-lookup"><span data-stu-id="0e61c-129">Outlook add-ins use only the `short` value.</span></span> <span data-ttu-id="0e61c-130">アドインで共有ランタイムを使用する場合Excelに値を明示的に設定します `long` 。</span><span class="sxs-lookup"><span data-stu-id="0e61c-130">If you want to use a shared runtime in an Excel add-in, explicitly set the value to `long`.</span></span> |

## <a name="see-also"></a><span data-ttu-id="0e61c-131">関連項目</span><span class="sxs-lookup"><span data-stu-id="0e61c-131">See also</span></span>

- [<span data-ttu-id="0e61c-132">ランタイム</span><span class="sxs-lookup"><span data-stu-id="0e61c-132">Runtimes</span></span>](runtimes.md)
- [<span data-ttu-id="0e61c-133">Office アドインを構成して共有 JavaScript ランタイムを使用する</span><span class="sxs-lookup"><span data-stu-id="0e61c-133">Configure your Office Add-in to use a shared JavaScript runtime</span></span>](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)
- [<span data-ttu-id="0e61c-134">イベント ベースのOutlook用にアドインを構成する</span><span class="sxs-lookup"><span data-stu-id="0e61c-134">Configure your Outlook add-in for event-based activation</span></span>](../../outlook/autolaunch.md)
