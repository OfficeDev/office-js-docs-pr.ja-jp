---
title: Office ダイアログ ボックスを使用してビデオを再生する
description: Office ダイアログボックスでビデオを開いて再生する方法について説明します。
ms.date: 01/29/2020
localization_priority: Normal
ms.openlocfilehash: f0d524996b105061b8e5d1b584d8b3e0d44eec7c
ms.sourcegitcommit: ceb8dd66f3fb9c963fce8446c2f6c65ead56fbc1
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/18/2020
ms.locfileid: "49131774"
---
# <a name="use-the-office-dialog-box-to-show-a-video"></a><span data-ttu-id="5586e-103">Office ダイアログボックスを使用してビデオを表示する</span><span class="sxs-lookup"><span data-stu-id="5586e-103">Use the Office dialog box to show a video</span></span>

<span data-ttu-id="5586e-104">この記事では、Office アドインダイアログボックスでビデオを再生する方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="5586e-104">This article explains how to play a video in an Office Add-in dialog box.</span></span>

> [!NOTE]
> <span data-ttu-id="5586e-105">この記事では、「office [アドインで office ダイアログ API を使用](dialog-api-in-office-add-ins.md)する」で説明されているように、office ダイアログボックスの使用に関する基本事項を理解していることを前提としています。</span><span class="sxs-lookup"><span data-stu-id="5586e-105">This article presumes you're familiar with the basics of using the Office dialog box as described in [Use the Office dialog API in your Office Add-ins](dialog-api-in-office-add-ins.md).</span></span>

<span data-ttu-id="5586e-106">Office ダイアログ API を使用してダイアログボックス内のビデオを再生するには、次の手順を実行します。</span><span class="sxs-lookup"><span data-stu-id="5586e-106">To play a video in a dialog box with the Office dialog API, follow these steps:</span></span>

1. <span data-ttu-id="5586e-107">Iframe を含むページを作成し、その他のコンテンツは作成しません。</span><span class="sxs-lookup"><span data-stu-id="5586e-107">Create a page containing an iframe and no other content.</span></span> <span data-ttu-id="5586e-108">このページは、ホストページと同じドメインにある必要があります。</span><span class="sxs-lookup"><span data-stu-id="5586e-108">The page must be in the same domain as the host page.</span></span> <span data-ttu-id="5586e-109">ホストページについての通知については、「 [ホストページからダイアログボックスを開く](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="5586e-109">For a reminder of what a host page is, see [Open a dialog box from a host page](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page).</span></span> <span data-ttu-id="5586e-110">`src`Iframe の属性で、オンラインビデオの URL をポイントします。</span><span class="sxs-lookup"><span data-stu-id="5586e-110">In the `src` attribute of the iframe, point to the URL of an online video.</span></span> <span data-ttu-id="5586e-111">ビデオの URL のプロトコルは HTTPS である必要があります。</span><span class="sxs-lookup"><span data-stu-id="5586e-111">The protocol of the video's URL must be HTTPS.</span></span> <span data-ttu-id="5586e-112">この記事では、このページ "video.dialogbox.html" を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="5586e-112">In this article, we'll call this page "video.dialogbox.html".</span></span> <span data-ttu-id="5586e-113">マークアップの例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="5586e-113">The following is an example of the markup:</span></span>

    ```HTML
    <iframe class="ms-firstrun-video__player"  width="640" height="360"
        src="https://www.youtube.com/embed/XVfOe5mFbAE?rel=0&autoplay=1"
        frameborder="0" allowfullscreen>
    </iframe>
    ```

2. <span data-ttu-id="5586e-114">ホスト ページで `displayDialogAsync` の呼び出しを使用して、video.dialogbox.html を開きます。</span><span class="sxs-lookup"><span data-stu-id="5586e-114">Use a call of `displayDialogAsync` in the host page to open video.dialogbox.html.</span></span>
3. <span data-ttu-id="5586e-115">ユーザーがダイアログ ボックスを閉じたときに、アドインに通知する必要がある場合は、`DialogEventReceived` イベントのハンドラーを登録して、12006 イベントを処理します。</span><span class="sxs-lookup"><span data-stu-id="5586e-115">If your add-in needs to know when the user closes the dialog box, register a handler for the `DialogEventReceived` event and handle the 12006 event.</span></span> <span data-ttu-id="5586e-116">詳細については、「 [Office ダイアログボックスでのエラーとイベント](dialog-handle-errors-events.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="5586e-116">For details, see [Errors and events in the Office dialog box](dialog-handle-errors-events.md).</span></span>

<span data-ttu-id="5586e-117">ダイアログボックスでビデオを再生する例については、「 [ビデオプレイスマット設計パターン](../design/first-run-experience-patterns.md#video-placemat)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="5586e-117">For a sample of a video playing in a dialog box, see the [video placemat design pattern](../design/first-run-experience-patterns.md#video-placemat).</span></span>

![Excel の前面にあるアドインダイアログボックスでビデオを再生していることを示すスクリーンショット](../images/video-placemats-dialog-open.png)
