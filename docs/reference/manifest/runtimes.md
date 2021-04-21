---
title: マニフェスト ファイル内のランタイム
description: Runtimes 要素は、アドインのランタイムを指定します。
ms.date: 04/16/2021
localization_priority: Normal
ms.openlocfilehash: 8f4a602c05b9af7bde9f644ef40b61a214e66cd5
ms.sourcegitcommit: da8ad214406f2e1cd80982af8a13090e76187dbd
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/21/2021
ms.locfileid: "51917087"
---
# <a name="runtimes-element"></a><span data-ttu-id="8cb82-103">Runtimes 要素</span><span class="sxs-lookup"><span data-stu-id="8cb82-103">Runtimes element</span></span>

<span data-ttu-id="8cb82-104">アドインのランタイムを指定します。</span><span class="sxs-lookup"><span data-stu-id="8cb82-104">Specifies the runtime of your add-in.</span></span> <span data-ttu-id="8cb82-105">要素の [`<Host>`](host.md) 子。</span><span class="sxs-lookup"><span data-stu-id="8cb82-105">Child of the [`<Host>`](host.md) element.</span></span>

> [!NOTE]
> <span data-ttu-id="8cb82-106">Windows 上Officeで実行する場合、マニフェストに要素を持つアドインは、それ以外の場合と同じ Web ビュー コントロールで必ずしも `<Runtimes>` 実行されるとは限りません。</span><span class="sxs-lookup"><span data-stu-id="8cb82-106">When running in Office on Windows, an add-in that has a `<Runtimes>` element in its manifest does not necessarily run in the same webview control as it otherwise would.</span></span> <span data-ttu-id="8cb82-107">Windows および Officeのバージョンでどの Web ビュー コントロールが通常使用されるのかを決定する方法の詳細については、「Office アドインで使用されるブラウザー」を [参照してください](../../concepts/browsers-used-by-office-web-add-ins.md)。WebView2 で Microsoft Edge を使用する条件 (クロムベース) が満たされている場合、アドインはそのブラウザーが要素を持っているかどうかに応じ、そのブラウザーを使用 `<Runtimes>` します。</span><span class="sxs-lookup"><span data-stu-id="8cb82-107">For more information about how the versions of Windows and Office determine what webview control is normally used, see [Browsers used by Office Add-ins](../../concepts/browsers-used-by-office-web-add-ins.md). If the conditions described there for using Microsoft Edge with WebView2 (Chromium-based) are met, then the add-in uses that browser whether or not it has a `<Runtimes>` element.</span></span> <span data-ttu-id="8cb82-108">ただし、これらの条件が満たされない場合、要素を持つアドインは、Windows または Microsoft 365 バージョンに関係なく、常に Internet Explorer `<Runtimes>` 11 を使用します。</span><span class="sxs-lookup"><span data-stu-id="8cb82-108">However, when those conditions are not met, an add-in with a `<Runtimes>` element always uses Internet Explorer 11 regardless of the Windows or Microsoft 365 version.</span></span>

<span data-ttu-id="8cb82-109">**アドインの種類:** 作業ウィンドウ, メール</span><span class="sxs-lookup"><span data-stu-id="8cb82-109">**Add-in type:** Task pane, Mail</span></span>

[!include[Runtimes support](../../includes/runtimes-note.md)]

## <a name="syntax"></a><span data-ttu-id="8cb82-110">構文</span><span class="sxs-lookup"><span data-stu-id="8cb82-110">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="8cb82-111">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="8cb82-111">Contained in</span></span>

[<span data-ttu-id="8cb82-112">Host</span><span class="sxs-lookup"><span data-stu-id="8cb82-112">Host</span></span>](host.md)

## <a name="child-elements"></a><span data-ttu-id="8cb82-113">子要素</span><span class="sxs-lookup"><span data-stu-id="8cb82-113">Child elements</span></span>

|  <span data-ttu-id="8cb82-114">要素</span><span class="sxs-lookup"><span data-stu-id="8cb82-114">Element</span></span> |  <span data-ttu-id="8cb82-115">必須</span><span class="sxs-lookup"><span data-stu-id="8cb82-115">Required</span></span>  |  <span data-ttu-id="8cb82-116">説明</span><span class="sxs-lookup"><span data-stu-id="8cb82-116">Description</span></span>  |
|:-----|:-----|:-----|
| [<span data-ttu-id="8cb82-117">ランタイム</span><span class="sxs-lookup"><span data-stu-id="8cb82-117">Runtime</span></span>](runtime.md) | <span data-ttu-id="8cb82-118">はい</span><span class="sxs-lookup"><span data-stu-id="8cb82-118">Yes</span></span> |  <span data-ttu-id="8cb82-119">アドインのランタイム。</span><span class="sxs-lookup"><span data-stu-id="8cb82-119">The runtime for your add-in.</span></span> |

## <a name="see-also"></a><span data-ttu-id="8cb82-120">関連項目</span><span class="sxs-lookup"><span data-stu-id="8cb82-120">See also</span></span>

- [<span data-ttu-id="8cb82-121">ランタイム</span><span class="sxs-lookup"><span data-stu-id="8cb82-121">Runtime</span></span>](runtime.md)
- [<span data-ttu-id="8cb82-122">Office アドインを構成して共有 JavaScript ランタイムを使用する</span><span class="sxs-lookup"><span data-stu-id="8cb82-122">Configure your Office Add-in to use a shared JavaScript runtime</span></span>](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)
- [<span data-ttu-id="8cb82-123">イベント ベースのライセンス認証用に Outlook アドインを構成する</span><span class="sxs-lookup"><span data-stu-id="8cb82-123">Configure your Outlook add-in for event-based activation</span></span>](../../outlook/autolaunch.md)
