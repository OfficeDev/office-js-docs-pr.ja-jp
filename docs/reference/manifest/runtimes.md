---
title: マニフェストファイル内のランタイム (プレビュー)
description: ''
ms.date: 02/21/2020
localization_priority: Normal
ms.openlocfilehash: 17e53b53d55ea9547cdfc5c4f89f8f4c3a7ab75e
ms.sourcegitcommit: dd6d00202f6466c27418247dad7bd136555a6036
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/26/2020
ms.locfileid: "42283880"
---
# <a name="runtimes-element-preview"></a><span data-ttu-id="41b4b-102">ランタイム要素 (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="41b4b-102">Runtimes element (preview)</span></span>

[!include[Running custom functions in browser runtime note](../../includes/excel-shared-runtime-preview-note.md)]

<span data-ttu-id="41b4b-103">アドインのランタイムを指定し、カスタム関数、リボンボタン、および作業ウィンドウを使用して同じ JavaScript ランタイムを使用できるようにします。</span><span class="sxs-lookup"><span data-stu-id="41b4b-103">Specifies the runtime of your add-in and enables custom functions, ribbon buttons, and the task pane to use the same JavaScript runtime.</span></span> <span data-ttu-id="41b4b-104">マニフェストファイル内`<Host>`の要素の子。</span><span class="sxs-lookup"><span data-stu-id="41b4b-104">Child of the `<Host>` element in your manifest file.</span></span> <span data-ttu-id="41b4b-105">詳細については、「[共有 JavaScript ランタイムを使用するように Excel アドインを構成する](../../excel/configure-your-add-in-to-use-a-shared-runtime.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="41b4b-105">For more information, see [Configure your Excel add-in to use a shared JavaScript runtime](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

<span data-ttu-id="41b4b-106">**アドインの種類:** 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="41b4b-106">**Add-in type:** Task pane</span></span>

> [!IMPORTANT]
> <span data-ttu-id="41b4b-107">共有ランタイムは現在プレビュー段階であり、Windows 上の Excel でのみ使用できます。</span><span class="sxs-lookup"><span data-stu-id="41b4b-107">Shared runtime is currently in preview and are only available on Excel on Windows.</span></span> <span data-ttu-id="41b4b-108">プレビュー機能を試すには、 [Office Insider](https://insider.office.com/)に参加する必要があります。</span><span class="sxs-lookup"><span data-stu-id="41b4b-108">To try the preview features, you will need to join [Office Insider](https://insider.office.com/).</span></span>

## <a name="syntax"></a><span data-ttu-id="41b4b-109">構文</span><span class="sxs-lookup"><span data-stu-id="41b4b-109">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="41b4b-110">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="41b4b-110">Contained in</span></span> 
[<span data-ttu-id="41b4b-111">Host</span><span class="sxs-lookup"><span data-stu-id="41b4b-111">Host</span></span>](./host.md)

## <a name="child-elements"></a><span data-ttu-id="41b4b-112">子要素</span><span class="sxs-lookup"><span data-stu-id="41b4b-112">Child elements</span></span>

|  <span data-ttu-id="41b4b-113">要素</span><span class="sxs-lookup"><span data-stu-id="41b4b-113">Element</span></span> |  <span data-ttu-id="41b4b-114">必須</span><span class="sxs-lookup"><span data-stu-id="41b4b-114">Required</span></span>  |  <span data-ttu-id="41b4b-115">説明</span><span class="sxs-lookup"><span data-stu-id="41b4b-115">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="41b4b-116">**ランタイム**</span><span class="sxs-lookup"><span data-stu-id="41b4b-116">**Runtime**</span></span>     | <span data-ttu-id="41b4b-117">はい</span><span class="sxs-lookup"><span data-stu-id="41b4b-117">Yes</span></span> |  <span data-ttu-id="41b4b-118">アドインのランタイム。</span><span class="sxs-lookup"><span data-stu-id="41b4b-118">The runtime for your add-in.</span></span>

## <a name="see-also"></a><span data-ttu-id="41b4b-119">関連項目</span><span class="sxs-lookup"><span data-stu-id="41b4b-119">See also</span></span>

- [<span data-ttu-id="41b4b-120">ランタイム</span><span class="sxs-lookup"><span data-stu-id="41b4b-120">Runtime</span></span>](runtime.md)
