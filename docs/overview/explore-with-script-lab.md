---
title: スクリプトラボを使用して Office JavaScript API を探索する
description: スクリプトラボを使用して、Office JS API とプロトタイプ機能を調査します。
ms.topic: article
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: 9ef81443fade450a7bea519a99cb607c320dd4f6
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/21/2019
ms.locfileid: "35128656"
---
# <a name="explore-office-javascript-api-using-script-lab"></a><span data-ttu-id="8c12c-103">スクリプトラボを使用して Office JavaScript API を探索する</span><span class="sxs-lookup"><span data-stu-id="8c12c-103">Explore Office JavaScript API using Script Lab</span></span>

<span data-ttu-id="8c12c-104">[スクリプトラボアドイン](https://store.office.com/app.aspx?assetid=WA104380862)は office ストアから無料で利用できます。これにより、Excel や Word などの office プログラムで作業しているときに OFFICE JavaScript API を調べることができます。</span><span class="sxs-lookup"><span data-stu-id="8c12c-104">The [Script Lab add-in](https://store.office.com/app.aspx?assetid=WA104380862), which is available free from the Office store, enables you to explore the Office JavaScript API while you're working in an Office program such as Excel or Word.</span></span> <span data-ttu-id="8c12c-105">スクリプトラボは、アドインに必要な機能を試作して検証する際に開発ツールキットに追加する便利なツールです。</span><span class="sxs-lookup"><span data-stu-id="8c12c-105">Script Lab is a convenient tool to add to your development toolkit as you prototype and verify functionality you want in your add-in.</span></span>

## <a name="what-is-script-lab"></a><span data-ttu-id="8c12c-106">スクリプトラボとは</span><span class="sxs-lookup"><span data-stu-id="8c12c-106">What is Script Lab?</span></span>

<span data-ttu-id="8c12c-107">スクリプトラボは、Excel、Word、または PowerPoint で Office JavaScript API を使用して Office アドインを開発する方法について学習する必要があるユーザーのためのツールです。</span><span class="sxs-lookup"><span data-stu-id="8c12c-107">Script Lab is a tool for anyone who wants to learn how to develop Office Add-ins using the Office JavaScript API in Excel, Word, or PowerPoint.</span></span> <span data-ttu-id="8c12c-108">これにより IntelliSense が提供され、Visual Studio Code で使用されるのと同じフレームワークである、使用可能なものと、モナコフレームワークに基づいて構築されているものがわかります。</span><span class="sxs-lookup"><span data-stu-id="8c12c-108">It provides IntelliSense so you can see what's available and is built on the Monaco framework, the same framework used by Visual Studio Code.</span></span> <span data-ttu-id="8c12c-109">スクリプトラボを使用すると、サンプルのライブラリにアクセスして、機能をすばやく試すことができます。また、サンプルを独自のコードの開始点として使用することもできます。</span><span class="sxs-lookup"><span data-stu-id="8c12c-109">Through Script Lab, you can access a library of samples to quickly try out features or you can use a sample as the starting point for your own code.</span></span> <span data-ttu-id="8c12c-110">スクリプトラボを使用してプレビュー Api を試すこともできます。</span><span class="sxs-lookup"><span data-stu-id="8c12c-110">You can even use Script Lab to try preview APIs.</span></span>

<span data-ttu-id="8c12c-111">これまでに良好なことがありますか?</span><span class="sxs-lookup"><span data-stu-id="8c12c-111">Sounds good so far?</span></span> <span data-ttu-id="8c12c-112">この1分間のビデオを見て、実行中のスクリプトラボを確認してください。</span><span class="sxs-lookup"><span data-stu-id="8c12c-112">Take a look at this one-minute video to see Script Lab in action.</span></span>

<span data-ttu-id="8c12c-113">[![Excel、Word、および PowerPoint で実行されているスクリプトラボを示すビデオをプレビューします。](../images/screenshot-wide-youtube.png 'スクリプトラボプレビューのビデオ')](https://aka.ms/scriptlabvideo)</span><span class="sxs-lookup"><span data-stu-id="8c12c-113">[![Preview video showing Script Lab running in Excel, Word, and PowerPoint.](../images/screenshot-wide-youtube.png 'Script Lab preview video')](https://aka.ms/scriptlabvideo)</span></span>

## <a name="key-features"></a><span data-ttu-id="8c12c-114">主な機能</span><span class="sxs-lookup"><span data-stu-id="8c12c-114">Key features</span></span>

<span data-ttu-id="8c12c-115">スクリプトラボ Office JavaScript API と prototype アドインの機能について調べるのに役立つさまざまな機能が用意されています。</span><span class="sxs-lookup"><span data-stu-id="8c12c-115">Script Lab offers a number of features to help you explore the Office JavaScript API and prototype add-in functionality.</span></span>

### <a name="explore-samples"></a><span data-ttu-id="8c12c-116">サンプルを検索する</span><span class="sxs-lookup"><span data-stu-id="8c12c-116">Explore samples</span></span>

<span data-ttu-id="8c12c-117">API を使用して特定のタスクを実行する方法を示す組み込みのサンプルスニペットのコレクションを使用して、すぐに作業を開始できます。</span><span class="sxs-lookup"><span data-stu-id="8c12c-117">Get started quickly with a collection of built-in sample snippets that show how to complete certain tasks with the API.</span></span> <span data-ttu-id="8c12c-118">サンプルを実行すると、作業ウィンドウまたはドキュメントの結果をすぐに確認したり、サンプルを調べて API のしくみを確認したり、サンプルスニペットを使用して、独自のアドインのプロトタイプ機能の基礎として使用したりすることもできます。</span><span class="sxs-lookup"><span data-stu-id="8c12c-118">You can run the samples to instantly see the result in the task pane or document, examine the samples to learn how the API works, and even use sample snippets as the basis for prototyping functionality of your own add-in.</span></span>

![サンプル](../images/script-lab-samples.jpg)

### <a name="code-and-style"></a><span data-ttu-id="8c12c-120">コードとスタイル</span><span class="sxs-lookup"><span data-stu-id="8c12c-120">Code and style</span></span>

<span data-ttu-id="8c12c-121">Office JS API を呼び出す JavaScript または TypeScript コードに加えて、各スニペットには、作業ウィンドウの外観を定義する、作業ウィンドウと CSS のコンテンツを定義する HTML マークアップも含まれています。</span><span class="sxs-lookup"><span data-stu-id="8c12c-121">In addition to JavaScript or TypeScript code that calls the Office JS API, each snippet also contains HTML markup that defines content of the task pane and CSS that defines the appearance of the task pane.</span></span> <span data-ttu-id="8c12c-122">HTML マークアップと CSS をカスタマイズして、独自のアドインの作業ウィンドウデザインを試作する際に、要素の配置とスタイル設定を試すことができます。</span><span class="sxs-lookup"><span data-stu-id="8c12c-122">You can customize the HTML markup and CSS to experiment with element placement and styling as you prototype task pane design for your own add-in.</span></span>

> [!TIP]
> <span data-ttu-id="8c12c-123">スニペット内でプレビュー Api を呼び出すには、スニペットのライブラリを更新して、ベータ CDN (`https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`) とプレビューの種類の定義`@types/office-js-preview`を使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="8c12c-123">To call preview APIs within a snippet, you'll need to update the snippet's libraries to use the beta CDN (`https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`) and the preview type definitions `@types/office-js-preview`.</span></span> <span data-ttu-id="8c12c-124">また、一部のプレビュー Api は、 [Office insider プログラム](https://products.office.com/office-insider)にサインアップし、Office の insider ビルドを実行している場合にのみアクセスできます。</span><span class="sxs-lookup"><span data-stu-id="8c12c-124">Additionally, some preview APIs are only accessible if you've signed up for the [Office Insider program](https://products.office.com/office-insider) and are running an Insider build of Office.</span></span>

### <a name="save-and-share-snippets"></a><span data-ttu-id="8c12c-125">スニペットの保存と共有</span><span class="sxs-lookup"><span data-stu-id="8c12c-125">Save and share snippets</span></span>

<span data-ttu-id="8c12c-126">既定では、スクリプトラボで開いたスニペットはブラウザーのキャッシュに保存されます。</span><span class="sxs-lookup"><span data-stu-id="8c12c-126">By default, snippets that you open in Script Lab will be saved to your browser cache.</span></span> <span data-ttu-id="8c12c-127">スニペットを完全に保存するには、それを[GitHub gist](https://gist.github.com)にエクスポートします。</span><span class="sxs-lookup"><span data-stu-id="8c12c-127">To save a snippet permanently, you can export it to a [GitHub gist](https://gist.github.com).</span></span> <span data-ttu-id="8c12c-128">独自にスニペットを保存するための secret gist を作成したり、他のユーザーと共有したりする予定がある場合は、共有された gist を作成します。</span><span class="sxs-lookup"><span data-stu-id="8c12c-128">Create a secret gist to save a snippet exclusively for your own use, or create a shared gist if you plan to share it with others.</span></span>

![共有オプション](../images/script-lab-share.jpg)

### <a name="import-snippets"></a><span data-ttu-id="8c12c-130">スニペットのインポート</span><span class="sxs-lookup"><span data-stu-id="8c12c-130">Import snippets</span></span>

<span data-ttu-id="8c12c-131">スニペットをスクリプトラボにインポートするには、スニペット YAML が格納されているパブリック[GitHub gist](https://gist.github.com)への URL を指定するか、スニペットの完全な yaml に貼り付けます。</span><span class="sxs-lookup"><span data-stu-id="8c12c-131">You can import a snippet into Script Lab either by specifying the URL to the public [GitHub gist](https://gist.github.com) where the snippet YAML is stored or by pasting in the complete YAML for the snippet.</span></span> <span data-ttu-id="8c12c-132">この機能は、他のユーザーが自分のスニペットを GitHub gist に発行するか、スニペットの YAML を提供することによって、自分のスニペットを共有しているシナリオで役立つことがあります。</span><span class="sxs-lookup"><span data-stu-id="8c12c-132">This feature may be useful in scenarios where someone else has shared their snippet with you by either publishing it to a GitHub gist or providing their snippet's YAML.</span></span>

![スニペットのインポートオプション](../images/script-lab-import-snippet.jpg)

## <a name="supported-clients"></a><span data-ttu-id="8c12c-134">サポートされるクライアント</span><span class="sxs-lookup"><span data-stu-id="8c12c-134">Supported clients</span></span>

<span data-ttu-id="8c12c-135">スクリプトラボは、Excel、Word、および PowerPoint の次のクライアントでサポートされています。</span><span class="sxs-lookup"><span data-stu-id="8c12c-135">Script Lab is supported for Excel, Word, and PowerPoint on the following clients.</span></span>

- <span data-ttu-id="8c12c-136">Office 2013 以降 (Windows)</span><span class="sxs-lookup"><span data-stu-id="8c12c-136">Office 2013 or later on Windows</span></span>
- <span data-ttu-id="8c12c-137">Office 2016 以降の Mac</span><span class="sxs-lookup"><span data-stu-id="8c12c-137">Office 2016 or later on Mac</span></span>
- <span data-ttu-id="8c12c-138">Web 上の Office</span><span class="sxs-lookup"><span data-stu-id="8c12c-138">Office on the web</span></span>

## <a name="next-steps"></a><span data-ttu-id="8c12c-139">次のステップ</span><span class="sxs-lookup"><span data-stu-id="8c12c-139">Next steps</span></span>

<span data-ttu-id="8c12c-140">[Office js](https://github.com/OfficeDev/office-js-snippets#office-js-snippets)の GitHub リポジトリに新しいスニペットを投稿することによって、スクリプトラボのサンプルライブラリを拡張することをお歓迎します。</span><span class="sxs-lookup"><span data-stu-id="8c12c-140">You're welcome to expand the sample library in Script Lab by contributing new snippets to the [office-js-snippets](https://github.com/OfficeDev/office-js-snippets#office-js-snippets) GitHub repository.</span></span>

<span data-ttu-id="8c12c-141">Office アドインを作成する準備ができたら、推奨されている Office アプリケーションの[5 分間のクイックスタート](/office/dev/add-ins/#5-minute-quick-starts)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="8c12c-141">When you're ready to create your Office Add-in, see the [5-minute quick start](/office/dev/add-ins/#5-minute-quick-starts) for your preferred Office application.</span></span>

## <a name="see-also"></a><span data-ttu-id="8c12c-142">関連項目</span><span class="sxs-lookup"><span data-stu-id="8c12c-142">See also</span></span>

- [<span data-ttu-id="8c12c-143">スクリプトラボの取得</span><span class="sxs-lookup"><span data-stu-id="8c12c-143">Get Script Lab</span></span>](https://store.office.com/app.aspx?assetid=WA104380862)
- [<span data-ttu-id="8c12c-144">スクリプトラボの詳細情報</span><span class="sxs-lookup"><span data-stu-id="8c12c-144">Learn more about Script Lab</span></span>](https://github.com/OfficeDev/script-lab#script-lab-a-microsoft-garage-project)
- [<span data-ttu-id="8c12c-145">開発者プログラムにサインアップする</span><span class="sxs-lookup"><span data-stu-id="8c12c-145">Sign up for the dev program</span></span>](https://developer.microsoft.com/office/dev-program)
