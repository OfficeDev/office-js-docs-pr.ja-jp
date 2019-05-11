---
title: スクリプトラボを使用して Office JavaScript API を探索する
description: スクリプトラボを使用して、Office JS API とプロトタイプ機能を調査します。
ms.topic: article
ms.date: 05/08/2019
localization_priority: Normal
ms.openlocfilehash: 968f1c748558ee0d14898a3ad0dbe44bdc7a532a
ms.sourcegitcommit: a99be9c4771c45f3e07e781646e0e649aa47213f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/11/2019
ms.locfileid: "33952229"
---
# <a name="explore-office-javascript-api-using-script-lab"></a><span data-ttu-id="7b4e2-103">スクリプトラボを使用して Office JavaScript API を探索する</span><span class="sxs-lookup"><span data-stu-id="7b4e2-103">Explore Office JavaScript API using Script Lab</span></span>

<span data-ttu-id="7b4e2-104">[スクリプトラボアドイン](https://store.office.com/app.aspx?assetid=WA104380862)は office ストアから無料で利用できます。これにより、Excel や Word などの office プログラムで作業しているときに OFFICE JavaScript API を調べることができます。</span><span class="sxs-lookup"><span data-stu-id="7b4e2-104">The [Script Lab add-in](https://store.office.com/app.aspx?assetid=WA104380862), which is available free from the Office store, enables you to explore the Office JavaScript API while you are working in an Office program such as Excel or Word.</span></span> <span data-ttu-id="7b4e2-105">スクリプトラボは、アドインに必要な機能を試作して検証する際に開発ツールキットに追加する便利なツールです。</span><span class="sxs-lookup"><span data-stu-id="7b4e2-105">Script Lab is a convenient tool to add to your development toolkit as you prototype and verify functionality you want in your add-in.</span></span>

## <a name="what-is-script-lab"></a><span data-ttu-id="7b4e2-106">スクリプトラボとは</span><span class="sxs-lookup"><span data-stu-id="7b4e2-106">What is Script Lab?</span></span>

<span data-ttu-id="7b4e2-107">スクリプトラボは、Excel、Word、または PowerPoint で Office JavaScript API を使用して Office アドインを開発する方法について学習する必要があるユーザーのためのツールです。</span><span class="sxs-lookup"><span data-stu-id="7b4e2-107">Script Lab is a tool for anyone who wants to learn how to develop Office Add-ins using the Office JavaScript API in Excel, Word, or PowerPoint.</span></span> <span data-ttu-id="7b4e2-108">これにより IntelliSense が提供され、Visual Studio Code で使用されるのと同じフレームワークである、使用可能なものと、モナコフレームワークに基づいて構築されているものがわかります。</span><span class="sxs-lookup"><span data-stu-id="7b4e2-108">It provides IntelliSense so you can see what's available and is built on the Monaco framework, the same framework used by Visual Studio Code.</span></span> <span data-ttu-id="7b4e2-109">スクリプトラボを使用すると、サンプルのライブラリにアクセスして、機能をすばやく試すことができます。また、独自のコードのベースとしてサンプルを選択することもできます。</span><span class="sxs-lookup"><span data-stu-id="7b4e2-109">Through Script Lab, you can access a library of samples to quickly try out features or you can choose a sample as the base for your own code.</span></span> <span data-ttu-id="7b4e2-110">また、 [office js](https://github.com/OfficeDev/office-js-snippets#office-js-snippets)にスニペットを追加してサンプルライブラリを拡張することも歓迎します。</span><span class="sxs-lookup"><span data-stu-id="7b4e2-110">You are also welcome to expand the sample library by adding snippets to the [office-js-snippets repo](https://github.com/OfficeDev/office-js-snippets#office-js-snippets).</span></span> <span data-ttu-id="7b4e2-111">スクリプトラボのもう1つの魅力的な機能は、[ユーザー](/office/dev/add-ins/excel/custom-functions-overview)が試すことができるように、ベータ版またはプレビュー機能です。</span><span class="sxs-lookup"><span data-stu-id="7b4e2-111">Another exciting feature of Script Lab is beta or preview functionality like [custom functions](/office/dev/add-ins/excel/custom-functions-overview) is available for you to try.</span></span>

> [!TIP]
> <span data-ttu-id="7b4e2-112">ベータ版またはプレビューに参加するには、 [Office Insider プログラム](https://products.office.com/office-insider)にサインアップする必要がある場合があります。</span><span class="sxs-lookup"><span data-stu-id="7b4e2-112">To participate in beta or preview, you may have to sign up for the [Office Insider program](https://products.office.com/office-insider).</span></span>

<span data-ttu-id="7b4e2-113">これまでに良好なことがありますか?</span><span class="sxs-lookup"><span data-stu-id="7b4e2-113">Sounds good so far?</span></span> <span data-ttu-id="7b4e2-114">この1分間のビデオを見て、実行中のスクリプトラボを確認してください。</span><span class="sxs-lookup"><span data-stu-id="7b4e2-114">Take a look at this one-minute video to see Script Lab in action.</span></span>

<span data-ttu-id="7b4e2-115">[![Excel、Word、PowerPoint Online で実行されているスクリプトラボを示すビデオをプレビューします。](../images/screenshot-wide-youtube.png 'スクリプトラボプレビューのビデオ')](https://aka.ms/scriptlabvideo)</span><span class="sxs-lookup"><span data-stu-id="7b4e2-115">[![Preview video showing Script Lab running in Excel, Word, and PowerPoint Online.](../images/screenshot-wide-youtube.png 'Script Lab preview video')](https://aka.ms/scriptlabvideo)</span></span>

## <a name="script-lab-supported-clients"></a><span data-ttu-id="7b4e2-116">スクリプトラボでサポートされているクライアント</span><span class="sxs-lookup"><span data-stu-id="7b4e2-116">Script Lab supported clients</span></span>

<span data-ttu-id="7b4e2-117">スクリプトラボは、Excel、Word、および PowerPoint の次のクライアントでサポートされています。</span><span class="sxs-lookup"><span data-stu-id="7b4e2-117">Script Lab is supported for Excel, Word, and PowerPoint on the following clients.</span></span>

- <span data-ttu-id="7b4e2-118">Windows 上の Office (Office 365 に接続)</span><span class="sxs-lookup"><span data-stu-id="7b4e2-118">Office on Windows (connected to Office 365)</span></span>
- <span data-ttu-id="7b4e2-119">Office for Mac (Office 365 に接続)</span><span class="sxs-lookup"><span data-stu-id="7b4e2-119">Office for Mac (connected to Office 365)</span></span>
- <span data-ttu-id="7b4e2-120">Office Online</span><span class="sxs-lookup"><span data-stu-id="7b4e2-120">Office Online</span></span>
- <span data-ttu-id="7b4e2-121">Office 2013 以降 (Windows)</span><span class="sxs-lookup"><span data-stu-id="7b4e2-121">Office 2013 or later on Windows</span></span>
- <span data-ttu-id="7b4e2-122">Office 2016 以降 (Mac 版)</span><span class="sxs-lookup"><span data-stu-id="7b4e2-122">Office 2016 or later for Mac</span></span>

## <a name="next-steps"></a><span data-ttu-id="7b4e2-123">次のステップ</span><span class="sxs-lookup"><span data-stu-id="7b4e2-123">Next steps</span></span>

<span data-ttu-id="7b4e2-124">Office アドインを作成する準備ができたら、推奨されている Office アプリケーションの[5 分間のクイックスタート](/office/dev/add-ins/#5-minute-quick-starts)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="7b4e2-124">When you're ready to create your Office Add-in, see the [5-minute quick start](/office/dev/add-ins/#5-minute-quick-starts) for your preferred Office application.</span></span>

## <a name="see-also"></a><span data-ttu-id="7b4e2-125">関連項目</span><span class="sxs-lookup"><span data-stu-id="7b4e2-125">See also</span></span>

- [<span data-ttu-id="7b4e2-126">スクリプトラボの取得</span><span class="sxs-lookup"><span data-stu-id="7b4e2-126">Get Script Lab</span></span>](https://store.office.com/app.aspx?assetid=WA104380862)
- [<span data-ttu-id="7b4e2-127">スクリプトラボの詳細情報</span><span class="sxs-lookup"><span data-stu-id="7b4e2-127">Learn more about Script Lab</span></span>](https://github.com/OfficeDev/script-lab#script-lab-a-microsoft-garage-project)
- [<span data-ttu-id="7b4e2-128">開発者プログラムにサインアップする</span><span class="sxs-lookup"><span data-stu-id="7b4e2-128">Sign up for the dev program</span></span>](https://developer.microsoft.com/office/dev-program)
