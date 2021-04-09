---
title: マニフェスト ファイル内のランタイム
description: Runtimes 要素は、アドインのランタイムを指定します。
ms.date: 04/08/2021
localization_priority: Normal
ms.openlocfilehash: a5cd05a0890615375bf3466caf70d22f9912d951
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652236"
---
# <a name="runtimes-element"></a><span data-ttu-id="2139a-103">Runtimes 要素</span><span class="sxs-lookup"><span data-stu-id="2139a-103">Runtimes element</span></span>

<span data-ttu-id="2139a-104">アドインのランタイムを指定します。</span><span class="sxs-lookup"><span data-stu-id="2139a-104">Specifies the runtime of your add-in.</span></span> <span data-ttu-id="2139a-105">要素の [`<Host>`](host.md) 子。</span><span class="sxs-lookup"><span data-stu-id="2139a-105">Child of the [`<Host>`](host.md) element.</span></span>

> [!NOTE]
> <span data-ttu-id="2139a-106">Windows でOffice実行すると、アドインは 11 ブラウザー Internet Explorer使用します。</span><span class="sxs-lookup"><span data-stu-id="2139a-106">When running in Office on Windows, your add-in uses the Internet Explorer 11 browser.</span></span>

<span data-ttu-id="2139a-107">**アドインの種類:** 作業ウィンドウ, メール</span><span class="sxs-lookup"><span data-stu-id="2139a-107">**Add-in type:** Task pane, Mail</span></span>

[!include[Runtimes support](../../includes/runtimes-note.md)]

## <a name="syntax"></a><span data-ttu-id="2139a-108">構文</span><span class="sxs-lookup"><span data-stu-id="2139a-108">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="2139a-109">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="2139a-109">Contained in</span></span>

[<span data-ttu-id="2139a-110">Host</span><span class="sxs-lookup"><span data-stu-id="2139a-110">Host</span></span>](host.md)

## <a name="child-elements"></a><span data-ttu-id="2139a-111">子要素</span><span class="sxs-lookup"><span data-stu-id="2139a-111">Child elements</span></span>

|  <span data-ttu-id="2139a-112">要素</span><span class="sxs-lookup"><span data-stu-id="2139a-112">Element</span></span> |  <span data-ttu-id="2139a-113">必須</span><span class="sxs-lookup"><span data-stu-id="2139a-113">Required</span></span>  |  <span data-ttu-id="2139a-114">説明</span><span class="sxs-lookup"><span data-stu-id="2139a-114">Description</span></span>  |
|:-----|:-----|:-----|
| [<span data-ttu-id="2139a-115">Runtime</span><span class="sxs-lookup"><span data-stu-id="2139a-115">Runtime</span></span>](runtime.md) | <span data-ttu-id="2139a-116">はい</span><span class="sxs-lookup"><span data-stu-id="2139a-116">Yes</span></span> |  <span data-ttu-id="2139a-117">アドインのランタイム。</span><span class="sxs-lookup"><span data-stu-id="2139a-117">The runtime for your add-in.</span></span> |

## <a name="see-also"></a><span data-ttu-id="2139a-118">関連項目</span><span class="sxs-lookup"><span data-stu-id="2139a-118">See also</span></span>

- [<span data-ttu-id="2139a-119">Runtime</span><span class="sxs-lookup"><span data-stu-id="2139a-119">Runtime</span></span>](runtime.md)
- [<span data-ttu-id="2139a-120">Office アドインを構成して共有 JavaScript ランタイムを使用する</span><span class="sxs-lookup"><span data-stu-id="2139a-120">Configure your Office Add-in to use a shared JavaScript runtime</span></span>](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)
- [<span data-ttu-id="2139a-121">イベント ベースのライセンス認証用に Outlook アドインを構成する</span><span class="sxs-lookup"><span data-stu-id="2139a-121">Configure your Outlook add-in for event-based activation</span></span>](../../outlook/autolaunch.md)
