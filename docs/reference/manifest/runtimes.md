---
title: マニフェスト ファイル内のランタイム
description: Runtimes 要素は、アドインのランタイムを指定します。
ms.date: 05/14/2021
localization_priority: Normal
ms.openlocfilehash: 80336674c6d954bb9e0c6892feb41cb2f03c5859
ms.sourcegitcommit: 693d364616b42eea66977eef47530adabc51a40f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/19/2021
ms.locfileid: "52555298"
---
# <a name="runtimes-element"></a><span data-ttu-id="f29ca-103">Runtimes 要素</span><span class="sxs-lookup"><span data-stu-id="f29ca-103">Runtimes element</span></span>

<span data-ttu-id="f29ca-104">アドインのランタイムを指定します。</span><span class="sxs-lookup"><span data-stu-id="f29ca-104">Specifies the runtime of your add-in.</span></span> <span data-ttu-id="f29ca-105">要素の [`<Host>`](host.md) 子。</span><span class="sxs-lookup"><span data-stu-id="f29ca-105">Child of the [`<Host>`](host.md) element.</span></span>

> [!NOTE]
> <span data-ttu-id="f29ca-106">Windows で Office で実行する場合、マニフェスト内に要素を持つアドインは、それ以外の場合と同じ Webview コントロールで必ずしも `<Runtimes>` 実行されるとは限りません。</span><span class="sxs-lookup"><span data-stu-id="f29ca-106">When running in Office on Windows, an add-in that has a `<Runtimes>` element in its manifest does not necessarily run in the same webview control as it otherwise would.</span></span> <span data-ttu-id="f29ca-107">Windows および Office のバージョンでどの webview コントロールが通常使用されるのかを決定する方法の詳細については、「Office アドインで使用されるブラウザー」を[参照](../../concepts/browsers-used-by-office-web-add-ins.md)してください。webView2 (Chromium ベース) で Microsoft Edge を使用する場合に説明されている条件が満たされている場合、アドインは要素を持っているかどうかに応じ、そのブラウザーを使用します `<Runtimes>` 。</span><span class="sxs-lookup"><span data-stu-id="f29ca-107">For more information about how the versions of Windows and Office determine what webview control is normally used, see [Browsers used by Office Add-ins](../../concepts/browsers-used-by-office-web-add-ins.md). If the conditions described there for using Microsoft Edge with WebView2 (Chromium-based) are met, then the add-in uses that browser whether or not it has a `<Runtimes>` element.</span></span> <span data-ttu-id="f29ca-108">ただし、これらの条件が満たされない場合、要素を持つアドインは、Windows またはバージョンに関係なく、常に Internet Explorer 11 Windows を `<Runtimes>` Microsoft 365します。</span><span class="sxs-lookup"><span data-stu-id="f29ca-108">However, when those conditions are not met, an add-in with a `<Runtimes>` element always uses Internet Explorer 11 regardless of the Windows or Microsoft 365 version.</span></span>

<span data-ttu-id="f29ca-109">**アドインの種類:** 作業ウィンドウ, メール</span><span class="sxs-lookup"><span data-stu-id="f29ca-109">**Add-in type:** Task pane, Mail</span></span>

[!include[Runtimes support](../../includes/runtimes-note.md)]

## <a name="syntax"></a><span data-ttu-id="f29ca-110">構文</span><span class="sxs-lookup"><span data-stu-id="f29ca-110">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="f29ca-111">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="f29ca-111">Contained in</span></span>

[<span data-ttu-id="f29ca-112">Host</span><span class="sxs-lookup"><span data-stu-id="f29ca-112">Host</span></span>](host.md)

## <a name="child-elements"></a><span data-ttu-id="f29ca-113">子要素</span><span class="sxs-lookup"><span data-stu-id="f29ca-113">Child elements</span></span>

|  <span data-ttu-id="f29ca-114">要素</span><span class="sxs-lookup"><span data-stu-id="f29ca-114">Element</span></span> |  <span data-ttu-id="f29ca-115">必須</span><span class="sxs-lookup"><span data-stu-id="f29ca-115">Required</span></span>  |  <span data-ttu-id="f29ca-116">説明</span><span class="sxs-lookup"><span data-stu-id="f29ca-116">Description</span></span>  |
|:-----|:-----|:-----|
| [<span data-ttu-id="f29ca-117">ランタイム</span><span class="sxs-lookup"><span data-stu-id="f29ca-117">Runtime</span></span>](runtime.md) | <span data-ttu-id="f29ca-118">必要</span><span class="sxs-lookup"><span data-stu-id="f29ca-118">Yes</span></span> |  <span data-ttu-id="f29ca-119">アドインのランタイム。</span><span class="sxs-lookup"><span data-stu-id="f29ca-119">The runtime for your add-in.</span></span> <span data-ttu-id="f29ca-120">**重要**: 現時点では、1 つの要素のみを定義 `<Runtime>` できます。</span><span class="sxs-lookup"><span data-stu-id="f29ca-120">**Important**: At present, you can only define one `<Runtime>` element.</span></span> |

## <a name="see-also"></a><span data-ttu-id="f29ca-121">関連項目</span><span class="sxs-lookup"><span data-stu-id="f29ca-121">See also</span></span>

- [<span data-ttu-id="f29ca-122">ランタイム</span><span class="sxs-lookup"><span data-stu-id="f29ca-122">Runtime</span></span>](runtime.md)
- [<span data-ttu-id="f29ca-123">Office アドインを構成して共有 JavaScript ランタイムを使用する</span><span class="sxs-lookup"><span data-stu-id="f29ca-123">Configure your Office Add-in to use a shared JavaScript runtime</span></span>](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)
- [<span data-ttu-id="f29ca-124">イベント ベースのOutlook用にアドインを構成する</span><span class="sxs-lookup"><span data-stu-id="f29ca-124">Configure your Outlook add-in for event-based activation</span></span>](../../outlook/autolaunch.md)
