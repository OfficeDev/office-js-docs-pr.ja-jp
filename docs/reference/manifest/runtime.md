---
title: マニフェスト ファイルのランタイム
description: Runtime 要素は、リボン、作業ウィンドウ、カスタム関数など、さまざまなコンポーネントに共有 JavaScript ランタイムを使用するようにアドインを構成します。
ms.date: 05/14/2021
localization_priority: Normal
ms.openlocfilehash: c59e5a23e53940aea46c758d710b4a455cb5c0cc
ms.sourcegitcommit: 693d364616b42eea66977eef47530adabc51a40f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/19/2021
ms.locfileid: "52555305"
---
# <a name="runtime-element"></a><span data-ttu-id="f3c84-103">ランタイム要素</span><span class="sxs-lookup"><span data-stu-id="f3c84-103">Runtime element</span></span>

<span data-ttu-id="f3c84-104">共有 JavaScript ランタイムを使用して、さまざまなコンポーネントがすべて同じランタイムで実行されるようにアドインを構成します。</span><span class="sxs-lookup"><span data-stu-id="f3c84-104">Configures your add-in to use a shared JavaScript runtime so that various components all run in the same runtime.</span></span> <span data-ttu-id="f3c84-105">要素の子 [`<Runtimes>`](runtimes.md) 。</span><span class="sxs-lookup"><span data-stu-id="f3c84-105">Child of the [`<Runtimes>`](runtimes.md) element.</span></span>

<span data-ttu-id="f3c84-106">**アドインの種類:** 作業ウィンドウ,メール</span><span class="sxs-lookup"><span data-stu-id="f3c84-106">**Add-in type:** Task pane, Mail</span></span>

[!include[Runtimes support](../../includes/runtimes-note.md)]

## <a name="syntax"></a><span data-ttu-id="f3c84-107">構文</span><span class="sxs-lookup"><span data-stu-id="f3c84-107">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="f3c84-108">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="f3c84-108">Contained in</span></span>

- [<span data-ttu-id="f3c84-109">ランタイム</span><span class="sxs-lookup"><span data-stu-id="f3c84-109">Runtimes</span></span>](runtimes.md)

## <a name="child-elements"></a><span data-ttu-id="f3c84-110">子要素</span><span class="sxs-lookup"><span data-stu-id="f3c84-110">Child elements</span></span>

|  <span data-ttu-id="f3c84-111">要素</span><span class="sxs-lookup"><span data-stu-id="f3c84-111">Element</span></span> |  <span data-ttu-id="f3c84-112">必須</span><span class="sxs-lookup"><span data-stu-id="f3c84-112">Required</span></span>  |  <span data-ttu-id="f3c84-113">説明</span><span class="sxs-lookup"><span data-stu-id="f3c84-113">Description</span></span>  |
|:-----|:-----|:-----|
| <span data-ttu-id="f3c84-114">[上書き](override.md) (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="f3c84-114">[Override](override.md) (preview)</span></span> | <span data-ttu-id="f3c84-115">いいえ</span><span class="sxs-lookup"><span data-stu-id="f3c84-115">No</span></span> | <span data-ttu-id="f3c84-116">**Outlook**:[デスクトップで起動イベント拡張ポイント](../../reference/manifest/extensionpoint.md#launchevent-preview)ハンドラーに必要な JavaScript ファイルの URL の場所 Outlookを指定します。</span><span class="sxs-lookup"><span data-stu-id="f3c84-116">**Outlook**: Specifies the URL location of the JavaScript file that Outlook Desktop requires for [LaunchEvent extension point](../../reference/manifest/extensionpoint.md#launchevent-preview) handlers.</span></span> <span data-ttu-id="f3c84-117">**重要**: 現在、定義できる要素は 1 つのみ `<Override>` で、型が必要です `javascript` 。</span><span class="sxs-lookup"><span data-stu-id="f3c84-117">**Important**: At present, you can only define one `<Override>` element and it must be of type `javascript`.</span></span>|

## <a name="attributes"></a><span data-ttu-id="f3c84-118">属性</span><span class="sxs-lookup"><span data-stu-id="f3c84-118">Attributes</span></span>

|  <span data-ttu-id="f3c84-119">属性</span><span class="sxs-lookup"><span data-stu-id="f3c84-119">Attribute</span></span>  |  <span data-ttu-id="f3c84-120">必須</span><span class="sxs-lookup"><span data-stu-id="f3c84-120">Required</span></span>  |  <span data-ttu-id="f3c84-121">説明</span><span class="sxs-lookup"><span data-stu-id="f3c84-121">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="f3c84-122">**resid**</span><span class="sxs-lookup"><span data-stu-id="f3c84-122">**resid**</span></span>  |  <span data-ttu-id="f3c84-123">はい</span><span class="sxs-lookup"><span data-stu-id="f3c84-123">Yes</span></span>  | <span data-ttu-id="f3c84-124">アドインの HTML ページの URL の場所を指定します。</span><span class="sxs-lookup"><span data-stu-id="f3c84-124">Specifies the URL location of the HTML page for your add-in.</span></span> <span data-ttu-id="f3c84-125">は `resid` 32 文字以内 `id` で、要素の属性と一致する必要があります `Url` `Resources` 。</span><span class="sxs-lookup"><span data-stu-id="f3c84-125">The `resid` can be no more than 32 characters and must match an `id` attribute of a `Url` element in the `Resources` element.</span></span> |
|  <span data-ttu-id="f3c84-126">**一生**</span><span class="sxs-lookup"><span data-stu-id="f3c84-126">**lifetime**</span></span>  |  <span data-ttu-id="f3c84-127">いいえ</span><span class="sxs-lookup"><span data-stu-id="f3c84-127">No</span></span>  | <span data-ttu-id="f3c84-128">デフォルト値 `lifetime` は `short` 、指定する必要はありません。</span><span class="sxs-lookup"><span data-stu-id="f3c84-128">The default value for `lifetime` is `short` and doesn't need to be specified.</span></span> <span data-ttu-id="f3c84-129">Outlookアドインでは、値のみを使用します `short` 。</span><span class="sxs-lookup"><span data-stu-id="f3c84-129">Outlook add-ins use only the `short` value.</span></span> <span data-ttu-id="f3c84-130">Excel アドインで共有ランタイムを使用する場合は、値を明示的に に 設定 `long` します。</span><span class="sxs-lookup"><span data-stu-id="f3c84-130">If you want to use a shared runtime in an Excel add-in, explicitly set the value to `long`.</span></span> |

## <a name="see-also"></a><span data-ttu-id="f3c84-131">関連項目</span><span class="sxs-lookup"><span data-stu-id="f3c84-131">See also</span></span>

- [<span data-ttu-id="f3c84-132">ランタイム</span><span class="sxs-lookup"><span data-stu-id="f3c84-132">Runtimes</span></span>](runtimes.md)
- [<span data-ttu-id="f3c84-133">Office アドインを構成して共有 JavaScript ランタイムを使用する</span><span class="sxs-lookup"><span data-stu-id="f3c84-133">Configure your Office Add-in to use a shared JavaScript runtime</span></span>](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)
- [<span data-ttu-id="f3c84-134">イベント ベースのアクティブ化用にOutlook アドインを構成する</span><span class="sxs-lookup"><span data-stu-id="f3c84-134">Configure your Outlook add-in for event-based activation</span></span>](../../outlook/autolaunch.md)
