---
title: Office ダイアログ ボックスを使用してビデオを再生する
description: '[ビデオの再生] ダイアログ ボックスでビデオを開いて再生するOffice説明します。'
ms.date: 01/29/2020
localization_priority: Normal
ms.openlocfilehash: 2519b2f105503a0479eee07d885a1543f5455343
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349883"
---
# <a name="use-the-office-dialog-box-to-show-a-video"></a><span data-ttu-id="28bd3-103">[ビデオをOffice]ダイアログ ボックスを使用してビデオを表示する</span><span class="sxs-lookup"><span data-stu-id="28bd3-103">Use the Office dialog box to show a video</span></span>

<span data-ttu-id="28bd3-104">この記事では、[アドイン] ダイアログ ボックスでビデオを再生Office説明します。</span><span class="sxs-lookup"><span data-stu-id="28bd3-104">This article explains how to play a video in an Office Add-in dialog box.</span></span>

> [!NOTE]
> <span data-ttu-id="28bd3-105">この記事では、「Office アドインで Office ダイアログ API を使用する」の説明に従って[、Office](dialog-api-in-office-add-ins.md)ダイアログ ボックスを使用する基本について理解している必要があります。</span><span class="sxs-lookup"><span data-stu-id="28bd3-105">This article presumes you're familiar with the basics of using the Office dialog box as described in [Use the Office dialog API in your Office Add-ins](dialog-api-in-office-add-ins.md).</span></span>

<span data-ttu-id="28bd3-106">ダイアログ API を使用してダイアログ ボックスでビデオを再生するにはOffice手順を実行します。</span><span class="sxs-lookup"><span data-stu-id="28bd3-106">To play a video in a dialog box with the Office dialog API, follow these steps:</span></span>

1. <span data-ttu-id="28bd3-107">iframe と他のコンテンツを含むページを作成します。</span><span class="sxs-lookup"><span data-stu-id="28bd3-107">Create a page containing an iframe and no other content.</span></span> <span data-ttu-id="28bd3-108">ページはホスト ページと同じドメインにある必要があります。</span><span class="sxs-lookup"><span data-stu-id="28bd3-108">The page must be in the same domain as the host page.</span></span> <span data-ttu-id="28bd3-109">ホスト ページの種類を確認するには、「ホスト ページからダイアログ ボックスを開 [く」を参照してください](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page)。</span><span class="sxs-lookup"><span data-stu-id="28bd3-109">For a reminder of what a host page is, see [Open a dialog box from a host page](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page).</span></span> <span data-ttu-id="28bd3-110">`src`iframe の属性で、オンライン ビデオの URL をポイントします。</span><span class="sxs-lookup"><span data-stu-id="28bd3-110">In the `src` attribute of the iframe, point to the URL of an online video.</span></span> <span data-ttu-id="28bd3-111">ビデオの URL のプロトコルは HTTPS である必要があります。</span><span class="sxs-lookup"><span data-stu-id="28bd3-111">The protocol of the video's URL must be HTTPS.</span></span> <span data-ttu-id="28bd3-112">この記事では、このページを "l" と呼video.dialogbox.htmします。</span><span class="sxs-lookup"><span data-stu-id="28bd3-112">In this article, we'll call this page "video.dialogbox.html".</span></span> <span data-ttu-id="28bd3-113">マークアップの例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="28bd3-113">The following is an example of the markup.</span></span>

    ```HTML
    <iframe class="ms-firstrun-video__player"  width="640" height="360"
        src="https://www.youtube.com/embed/XVfOe5mFbAE?rel=0&autoplay=1"
        frameborder="0" allowfullscreen>
    </iframe>
    ```

2. <span data-ttu-id="28bd3-114">ホスト ページで `displayDialogAsync` の呼び出しを使用して、video.dialogbox.html を開きます。</span><span class="sxs-lookup"><span data-stu-id="28bd3-114">Use a call of `displayDialogAsync` in the host page to open video.dialogbox.html.</span></span>
3. <span data-ttu-id="28bd3-115">ユーザーがダイアログ ボックスを閉じたときに、アドインに通知する必要がある場合は、`DialogEventReceived` イベントのハンドラーを登録して、12006 イベントを処理します。</span><span class="sxs-lookup"><span data-stu-id="28bd3-115">If your add-in needs to know when the user closes the dialog box, register a handler for the `DialogEventReceived` event and handle the 12006 event.</span></span> <span data-ttu-id="28bd3-116">詳細については、「エラーと[イベント」ダイアログ ボックスOffice参照してください](dialog-handle-errors-events.md)。</span><span class="sxs-lookup"><span data-stu-id="28bd3-116">For details, see [Errors and events in the Office dialog box](dialog-handle-errors-events.md).</span></span>

<span data-ttu-id="28bd3-117">ダイアログ ボックスで再生するビデオのサンプルについては、ビデオの配置パターン [を参照してください](../design/first-run-experience-patterns.md#video-placemat)。</span><span class="sxs-lookup"><span data-stu-id="28bd3-117">For a sample of a video playing in a dialog box, see the [video placemat design pattern](../design/first-run-experience-patterns.md#video-placemat).</span></span>

![アプリの前にあるアドイン ダイアログ ボックスで再生されているビデオを示すExcel。](../images/video-placemats-dialog-open.png)
