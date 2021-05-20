---
title: マニフェスト ファイルのランタイム
description: ランタイム要素は、アドインのランタイムを指定します。
ms.date: 05/14/2021
localization_priority: Normal
ms.openlocfilehash: 80336674c6d954bb9e0c6892feb41cb2f03c5859
ms.sourcegitcommit: 693d364616b42eea66977eef47530adabc51a40f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/19/2021
ms.locfileid: "52555298"
---
# <a name="runtimes-element"></a><span data-ttu-id="6dfc6-103">ランタイム要素</span><span class="sxs-lookup"><span data-stu-id="6dfc6-103">Runtimes element</span></span>

<span data-ttu-id="6dfc6-104">アドインのランタイムを指定します。</span><span class="sxs-lookup"><span data-stu-id="6dfc6-104">Specifies the runtime of your add-in.</span></span> <span data-ttu-id="6dfc6-105">要素の子 [`<Host>`](host.md) 。</span><span class="sxs-lookup"><span data-stu-id="6dfc6-105">Child of the [`<Host>`](host.md) element.</span></span>

> [!NOTE]
> <span data-ttu-id="6dfc6-106">WindowsでOfficeで実行する場合、マニフェストに要素を持つアドイン `<Runtimes>` は、必ずしも他の方法と同じ WebView コントロールで実行されるとは限りません。</span><span class="sxs-lookup"><span data-stu-id="6dfc6-106">When running in Office on Windows, an add-in that has a `<Runtimes>` element in its manifest does not necessarily run in the same webview control as it otherwise would.</span></span> <span data-ttu-id="6dfc6-107">WindowsとOfficeのバージョンが通常使用される webview コントロールを決定する方法の詳細については、「Office[アドインで使用されるブラウザー](../../concepts/browsers-used-by-office-web-add-ins.md)」を参照してください。WebView2 (Chromium ベース) でMicrosoft Edgeを使用する場合、アドインは要素を持つかどうかにかかわらず、そのブラウザーを使用します `<Runtimes>` 。</span><span class="sxs-lookup"><span data-stu-id="6dfc6-107">For more information about how the versions of Windows and Office determine what webview control is normally used, see [Browsers used by Office Add-ins](../../concepts/browsers-used-by-office-web-add-ins.md). If the conditions described there for using Microsoft Edge with WebView2 (Chromium-based) are met, then the add-in uses that browser whether or not it has a `<Runtimes>` element.</span></span> <span data-ttu-id="6dfc6-108">ただし、これらの条件が満たされない場合、要素を含むアドインでは `<Runtimes>` 、WindowsやMicrosoft 365のバージョンに関係なく、常に Internet Explorer 11 が使用されます。</span><span class="sxs-lookup"><span data-stu-id="6dfc6-108">However, when those conditions are not met, an add-in with a `<Runtimes>` element always uses Internet Explorer 11 regardless of the Windows or Microsoft 365 version.</span></span>

<span data-ttu-id="6dfc6-109">**アドインの種類:** 作業ウィンドウ,メール</span><span class="sxs-lookup"><span data-stu-id="6dfc6-109">**Add-in type:** Task pane, Mail</span></span>

[!include[Runtimes support](../../includes/runtimes-note.md)]

## <a name="syntax"></a><span data-ttu-id="6dfc6-110">構文</span><span class="sxs-lookup"><span data-stu-id="6dfc6-110">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="6dfc6-111">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="6dfc6-111">Contained in</span></span>

[<span data-ttu-id="6dfc6-112">Host</span><span class="sxs-lookup"><span data-stu-id="6dfc6-112">Host</span></span>](host.md)

## <a name="child-elements"></a><span data-ttu-id="6dfc6-113">子要素</span><span class="sxs-lookup"><span data-stu-id="6dfc6-113">Child elements</span></span>

|  <span data-ttu-id="6dfc6-114">要素</span><span class="sxs-lookup"><span data-stu-id="6dfc6-114">Element</span></span> |  <span data-ttu-id="6dfc6-115">必須</span><span class="sxs-lookup"><span data-stu-id="6dfc6-115">Required</span></span>  |  <span data-ttu-id="6dfc6-116">説明</span><span class="sxs-lookup"><span data-stu-id="6dfc6-116">Description</span></span>  |
|:-----|:-----|:-----|
| [<span data-ttu-id="6dfc6-117">ランタイム</span><span class="sxs-lookup"><span data-stu-id="6dfc6-117">Runtime</span></span>](runtime.md) | <span data-ttu-id="6dfc6-118">はい</span><span class="sxs-lookup"><span data-stu-id="6dfc6-118">Yes</span></span> |  <span data-ttu-id="6dfc6-119">アドインのランタイム。</span><span class="sxs-lookup"><span data-stu-id="6dfc6-119">The runtime for your add-in.</span></span> <span data-ttu-id="6dfc6-120">**重要**: 現在、定義できる要素は 1 つだけです `<Runtime>` 。</span><span class="sxs-lookup"><span data-stu-id="6dfc6-120">**Important**: At present, you can only define one `<Runtime>` element.</span></span> |

## <a name="see-also"></a><span data-ttu-id="6dfc6-121">関連項目</span><span class="sxs-lookup"><span data-stu-id="6dfc6-121">See also</span></span>

- [<span data-ttu-id="6dfc6-122">ランタイム</span><span class="sxs-lookup"><span data-stu-id="6dfc6-122">Runtime</span></span>](runtime.md)
- [<span data-ttu-id="6dfc6-123">Office アドインを構成して共有 JavaScript ランタイムを使用する</span><span class="sxs-lookup"><span data-stu-id="6dfc6-123">Configure your Office Add-in to use a shared JavaScript runtime</span></span>](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)
- [<span data-ttu-id="6dfc6-124">イベント ベースのアクティブ化用にOutlook アドインを構成する</span><span class="sxs-lookup"><span data-stu-id="6dfc6-124">Configure your Outlook add-in for event-based activation</span></span>](../../outlook/autolaunch.md)
