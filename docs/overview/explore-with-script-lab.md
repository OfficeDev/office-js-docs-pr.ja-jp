---
title: Script Lab を使用して Office JavaScript API を探索する
description: Script Lab を使用して、Office JS API およびプロトタイプの機能を調べます。
ms.date: 06/18/2021
ms.topic: conceptual
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: 7f4b67dd2369181e5d7b2b92496c8259ffd5c120
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/23/2021
ms.locfileid: "53077009"
---
# <a name="explore-office-javascript-api-using-script-lab"></a><span data-ttu-id="2e0ad-103">Script Lab を使用して Office JavaScript API を探索する</span><span class="sxs-lookup"><span data-stu-id="2e0ad-103">Explore Office JavaScript API using Script Lab</span></span>

<span data-ttu-id="2e0ad-104">AppSource から無料で入手できる [Script Lab](https://appsource.microsoft.com/product/office/WA104380862) および [Outlook 用 Script Lab](https://appsource.microsoft.com/product/office/wa200001603) アドインを使用すると、Excel や Outlook などの Office プログラムでの作業中に Office JavaScript API を調査できます。</span><span class="sxs-lookup"><span data-stu-id="2e0ad-104">The [Script Lab](https://appsource.microsoft.com/product/office/WA104380862) and [Script Lab for Outlook](https://appsource.microsoft.com/product/office/wa200001603) add-ins, available free from AppSource, enable you to explore the Office JavaScript API while you're working in an Office program such as Excel or Outlook.</span></span> <span data-ttu-id="2e0ad-105">Script Lab は、独自のアドインで必要な機能のプロトタイプを作成して検証するときに、開発ツールキットに追加する便利なツールです。</span><span class="sxs-lookup"><span data-stu-id="2e0ad-105">Script Lab is a convenient tool to add to your development toolkit as you prototype and verify functionality you want in your own add-in.</span></span>

## <a name="what-is-script-lab"></a><span data-ttu-id="2e0ad-106">Script Lab とは</span><span class="sxs-lookup"><span data-stu-id="2e0ad-106">What is Script Lab?</span></span>

<span data-ttu-id="2e0ad-107">Script Lab は、Excel、Outlook、Word、PowerPoint で Office JavaScript API を使用して Office アドインを開発する方法を学習したい人のためのツールです。</span><span class="sxs-lookup"><span data-stu-id="2e0ad-107">Script Lab is a tool for anyone who wants to learn how to develop Office Add-ins using the Office JavaScript API in Excel, Outlook, Word, and PowerPoint.</span></span> <span data-ttu-id="2e0ad-108">IntelliSense を提供しているので、何が利用できるのかを見ることができ、Visual Studio Code で使用されているのと同じフレームワークである Monaco フレームワークの上に構築されています。</span><span class="sxs-lookup"><span data-stu-id="2e0ad-108">It provides IntelliSense so you can see what's available and is built on the Monaco framework, the same framework used by Visual Studio Code.</span></span> <span data-ttu-id="2e0ad-109">Script Lab では、サンプルのライブラリにアクセスして、簡単に機能を試すことができます。また、独自のコードの開始点としてサンプルを使用することもできます。</span><span class="sxs-lookup"><span data-stu-id="2e0ad-109">Through Script Lab, you can access a library of samples to quickly try out features or you can use a sample as the starting point for your own code.</span></span> <span data-ttu-id="2e0ad-110">Script Lab を使用して、プレビュー API を試すこともできます。</span><span class="sxs-lookup"><span data-stu-id="2e0ad-110">You can even use Script Lab to try preview APIs.</span></span>

<span data-ttu-id="2e0ad-p103">いかがでしたか? この 1 分間のビデオを見て、Script Lab の動作を確認します。</span><span class="sxs-lookup"><span data-stu-id="2e0ad-p103">Sounds good so far? Take a look at this one-minute video to see Script Lab in action.</span></span>

<span data-ttu-id="2e0ad-113">[![Excel、Word、PowerPoint での Script Lab の実行を紹介するプレビュー ビデオ。](../images/screenshot-wide-youtube.png 'Script Lab のプレビュー ビデオ。')](https://aka.ms/scriptlabvideo)</span><span class="sxs-lookup"><span data-stu-id="2e0ad-113">[![Preview video showing Script Lab running in Excel, Word, and PowerPoint.](../images/screenshot-wide-youtube.png 'Script Lab preview video.')](https://aka.ms/scriptlabvideo)</span></span>

## <a name="key-features"></a><span data-ttu-id="2e0ad-114">主な機能</span><span class="sxs-lookup"><span data-stu-id="2e0ad-114">Key features</span></span>

<span data-ttu-id="2e0ad-115">Script Lab には、Office JavaScript API およびプロトタイプ アドインの機能の調査に役立つ機能が多数用意されています。</span><span class="sxs-lookup"><span data-stu-id="2e0ad-115">Script Lab offers a number of features to help you explore the Office JavaScript API and prototype add-in functionality.</span></span>

### <a name="explore-samples"></a><span data-ttu-id="2e0ad-116">サンプルの確認</span><span class="sxs-lookup"><span data-stu-id="2e0ad-116">Explore samples</span></span>

<span data-ttu-id="2e0ad-117">API を使用してタスクを完了する方法を示す組み込みのサンプル スニペットのコレクションを使用してすぐに開始できます。</span><span class="sxs-lookup"><span data-stu-id="2e0ad-117">Get started quickly with a collection of built-in sample snippets that show how to complete tasks with the API.</span></span> <span data-ttu-id="2e0ad-118">サンプルを実行すると、作業ウィンドウまたはドキュメントですばやく結果を表示したり、API のしくみをサンプルで確認して学んだり、独自のアドインのプロトタイプにサンプルを使用したりもできます。</span><span class="sxs-lookup"><span data-stu-id="2e0ad-118">You can run the samples to instantly see the result in the task pane or document, examine the samples to learn how the API works, and even use samples to prototype your own add-in.</span></span>

![サンプル。](../images/script-lab-samples.jpg)

### <a name="code-and-style"></a><span data-ttu-id="2e0ad-120">コードとスタイル</span><span class="sxs-lookup"><span data-stu-id="2e0ad-120">Code and style</span></span>

<span data-ttu-id="2e0ad-121">Office JS API を呼び出す JavaScript または TypeScript コードに加えて、各スニペットには、作業ウィンドウのコンテンツを定義する HTML マークアップと、作業ウィンドウの外観を定義する CSS も含まれています。</span><span class="sxs-lookup"><span data-stu-id="2e0ad-121">In addition to JavaScript or TypeScript code that calls the Office JS API, each snippet also contains HTML markup that defines content of the task pane and CSS that defines the appearance of the task pane.</span></span> <span data-ttu-id="2e0ad-122">HTML マークアップと CSS をカスタマイズして、独自のアドインの作業ウィンドウ デザインのプロトタイプを作成する際に、要素の配置とスタイル設定を試すことができます。</span><span class="sxs-lookup"><span data-stu-id="2e0ad-122">You can customize the HTML markup and CSS to experiment with element placement and styling as you prototype task pane design for your own add-in.</span></span>

> [!TIP]
> <span data-ttu-id="2e0ad-123">スニペット内でプレビュー API を呼び出すには、スニペットのライブラリを更新して、ベータ CDN (`https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`) とプレビューの種類の定義 `@types/office-js-preview` を使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="2e0ad-123">To call preview APIs within a snippet, you'll need to update the snippet's libraries to use the beta CDN (`https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`) and the preview type definitions `@types/office-js-preview`.</span></span> <span data-ttu-id="2e0ad-124">また、一部のプレビュー API は、[Office Insider プログラム](https://insider.office.com)にサインアップして、Insider ビルドの Office を実行している場合にのみアクセスできます。</span><span class="sxs-lookup"><span data-stu-id="2e0ad-124">Additionally, some preview APIs are only accessible if you've signed up for the [Office Insider program](https://insider.office.com) and are running an Insider build of Office.</span></span>

### <a name="save-and-share-snippets"></a><span data-ttu-id="2e0ad-125">スニペットの保存と共有</span><span class="sxs-lookup"><span data-stu-id="2e0ad-125">Save and share snippets</span></span>

<span data-ttu-id="2e0ad-126">既定では、Script Lab で開いたスニペットはブラウザーのキャッシュに保存されます。</span><span class="sxs-lookup"><span data-stu-id="2e0ad-126">By default, snippets that you open in Script Lab will be saved to your browser cache.</span></span> <span data-ttu-id="2e0ad-127">スニペットを完全に保存するには、そのスニペットを [GitHub の Gist](https://gist.github.com) にエクスポートします。</span><span class="sxs-lookup"><span data-stu-id="2e0ad-127">To save a snippet permanently, you can export it to a [GitHub gist](https://gist.github.com).</span></span> <span data-ttu-id="2e0ad-128">自分専用にスニペットを保存するには、秘密の Gist を作成するか、他のユーザーと共有する予定がある場合はパブリックの Gist を作成します。</span><span class="sxs-lookup"><span data-stu-id="2e0ad-128">Create a secret gist to save a snippet exclusively for your own use, or create a public gist if you plan to share it with others.</span></span>

![共有オプション。](../images/script-lab-share.jpg)

### <a name="import-snippets"></a><span data-ttu-id="2e0ad-130">スニペットのインポート</span><span class="sxs-lookup"><span data-stu-id="2e0ad-130">Import snippets</span></span>

<span data-ttu-id="2e0ad-131">スニペット YAML が保存されているパブリック [ GitHub の Gist ](https://gist.github.com) に URL を指定するか、スニペットの完全な YAML を貼り付けて、スニペットを Script Lab にインポートできます。</span><span class="sxs-lookup"><span data-stu-id="2e0ad-131">You can import a snippet into Script Lab either by specifying the URL to the public [GitHub gist](https://gist.github.com) where the snippet YAML is stored or by pasting in the complete YAML for the snippet.</span></span> <span data-ttu-id="2e0ad-132">この機能は、GitHub の Gist にスニペットを公開するか、スニペットの YAML を提供すると、他のユーザーがスニペットを自分と共有しているシナリオで役立ちます。</span><span class="sxs-lookup"><span data-stu-id="2e0ad-132">This feature may be useful in scenarios where someone else has shared their snippet with you by either publishing it to a GitHub gist or providing their snippet's YAML.</span></span>

![スニペットのインポート オプション。](../images/script-lab-import-snippet.jpg)

## <a name="supported-clients"></a><span data-ttu-id="2e0ad-134">サポートされるクライアント</span><span class="sxs-lookup"><span data-stu-id="2e0ad-134">Supported clients</span></span>

<span data-ttu-id="2e0ad-135">Script Lab は、次のクライアント上の Excel、Word、PowerPoint でサポートされています。</span><span class="sxs-lookup"><span data-stu-id="2e0ad-135">Script Lab is supported for Excel, Word, and PowerPoint on the following clients.</span></span>

- <span data-ttu-id="2e0ad-136">Microsoft 365 Office サブスクリプション</span><span class="sxs-lookup"><span data-stu-id="2e0ad-136">Microsoft 365 Office subscription</span></span>
- <span data-ttu-id="2e0ad-137">Mac での Office 2016 以降</span><span class="sxs-lookup"><span data-stu-id="2e0ad-137">Office 2016 or later on Mac</span></span>
- <span data-ttu-id="2e0ad-138">Office on the web</span><span class="sxs-lookup"><span data-stu-id="2e0ad-138">Office on the web</span></span>

<span data-ttu-id="2e0ad-139">Outlook 用 Script Lab は、次のクライアントで使用できます。</span><span class="sxs-lookup"><span data-stu-id="2e0ad-139">Script Lab for Outlook is available on the following clients.</span></span>

- <span data-ttu-id="2e0ad-140">Microsoft 365 Office サブスクリプション</span><span class="sxs-lookup"><span data-stu-id="2e0ad-140">Microsoft 365 Office subscription</span></span>
- <span data-ttu-id="2e0ad-141">Mac 用 Outlook 2016 以降</span><span class="sxs-lookup"><span data-stu-id="2e0ad-141">Outlook 2016 or later on Mac</span></span>
- <span data-ttu-id="2e0ad-142">Chrome、Microsoft Edge、または Safari ブラウザーを使用している場合の Outlook on the web</span><span class="sxs-lookup"><span data-stu-id="2e0ad-142">Outlook on the web when using Chrome, Microsoft Edge, or Safari browsers</span></span>

<span data-ttu-id="2e0ad-143">Outlook 用 Script Lab の詳細については、関連する[ブログの投稿](https://developer.microsoft.com/outlook/blogs/script-lab-now-supports-outlook/)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="2e0ad-143">For more details on Script Lab for Outlook, see the related [blog post](https://developer.microsoft.com/outlook/blogs/script-lab-now-supports-outlook/).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="2e0ad-144">Script Lab は、2021 年のある時点で、Internet Explorer を使ってアドインをホストするプラットフォームと Office バージョンの組み合わせでの動作を停止します。これには、Office 2019 までの 1 回限りの購入の Office と、Microsoft 365 (サブスクリプション) Office の一部の旧バージョンが含まれます。</span><span class="sxs-lookup"><span data-stu-id="2e0ad-144">Sometime in 2021, Script Lab will stop working on the combinations of platform and Office version that use Internet Explorer to host add-ins. This includes one-time purchase Office versions through Office 2019, and some older versions of Microsoft 365 (subscription) Office.</span></span> <span data-ttu-id="2e0ad-145">(詳細については、「[Office アドインによって使用されるブラウザー](../concepts/browsers-used-by-office-web-add-ins.md)」を参照してください。) Script Lab で Office JavaScript Library API の調査およびテストを行うには、他のプラットフォームとバージョンの組み合わせが必要です。</span><span class="sxs-lookup"><span data-stu-id="2e0ad-145">(For more information, see [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md).) You'll need other platform and version combinations to explore and test the Office JavaScript Library APIs with Script Lab.</span></span> <span data-ttu-id="2e0ad-146">ただし、これらの API の動作は Internet Explorer でも変わらないので、これは実際には Script Lab の短所とは言えません。</span><span class="sxs-lookup"><span data-stu-id="2e0ad-146">But the behavior of these APIs isn't different in Internet Explorer, so this isn't really a weakness of Script Lab.</span></span> <span data-ttu-id="2e0ad-147">[AppSource](/office/dev/store/submit-to-appsource-via-partner-center) に送信する Office アドイン は、Internet Explorer を使用してアドインをホストするプラットフォームとバージョンの組み合わせに対応している必要があります。</span><span class="sxs-lookup"><span data-stu-id="2e0ad-147">Note that Office Add-ins submitted to [AppSource](/office/dev/store/submit-to-appsource-via-partner-center) must support the platform and version combinations that use Internet Explorer to host add-ins.</span></span>

## <a name="next-steps"></a><span data-ttu-id="2e0ad-148">次の手順</span><span class="sxs-lookup"><span data-stu-id="2e0ad-148">Next steps</span></span>

<span data-ttu-id="2e0ad-149">Excel、Word、または PowerPoint で Script Lab を使用するには、AppSource から [Script Lab アドイン](https://appsource.microsoft.com/product/office/WA104380862)をインストールします。</span><span class="sxs-lookup"><span data-stu-id="2e0ad-149">To use Script Lab in Excel, Word, or PowerPoint, install the [Script Lab add-in](https://appsource.microsoft.com/product/office/WA104380862) from AppSource.</span></span> 

<span data-ttu-id="2e0ad-150">Outlook 用 Script Lab を使用するには、AppSource から [Outlook 用 Script Lab のアドイン](https://appsource.microsoft.com/product/office/wa200001603)をインストールします。</span><span class="sxs-lookup"><span data-stu-id="2e0ad-150">To use Script Lab for Outlook, install the [Script Lab for Outlook add-in](https://appsource.microsoft.com/product/office/wa200001603) from AppSource.</span></span>

<span data-ttu-id="2e0ad-151">新しいスニペットを [office-js-snippets](https://github.com/OfficeDev/office-js-snippets#office-js-snippets) GitHub リポジトリに投稿し、Script Lab のサンプル ライブラリを拡張してください。</span><span class="sxs-lookup"><span data-stu-id="2e0ad-151">You're welcome to expand the sample library in Script Lab by contributing new snippets to the [office-js-snippets](https://github.com/OfficeDev/office-js-snippets#office-js-snippets) GitHub repository.</span></span>

<span data-ttu-id="2e0ad-152">最初の Office アドインを作成する準備ができたら、[Excel](../quickstarts/excel-quickstart-jquery.md)、[Outlook](../quickstarts/outlook-quickstart.md)、[Word](../quickstarts/word-quickstart.md)、[OneNote ](../quickstarts/onenote-quickstart.md)、[PowerPoint](../quickstarts/powerpoint-quickstart.md)、または [Project](../quickstarts/project-quickstart.md) のクイック スタートを試してください。</span><span class="sxs-lookup"><span data-stu-id="2e0ad-152">When you're ready to create your first Office Add-in, try out the quick start for [Excel](../quickstarts/excel-quickstart-jquery.md), [Outlook](../quickstarts/outlook-quickstart.md), [Word](../quickstarts/word-quickstart.md), [OneNote](../quickstarts/onenote-quickstart.md), [PowerPoint](../quickstarts/powerpoint-quickstart.md), or [Project](../quickstarts/project-quickstart.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="2e0ad-153">関連項目</span><span class="sxs-lookup"><span data-stu-id="2e0ad-153">See also</span></span>

- [<span data-ttu-id="2e0ad-154">Excel、Word、または Powerpoint 用 Script Lab を取得する</span><span class="sxs-lookup"><span data-stu-id="2e0ad-154">Get Script Lab for Excel, Word, or Powerpoint</span></span>](https://appsource.microsoft.com/product/office/WA104380862)
- [<span data-ttu-id="2e0ad-155">Outlook 用 Script Lab を取得する</span><span class="sxs-lookup"><span data-stu-id="2e0ad-155">Get Script Lab for Outlook</span></span>](https://appsource.microsoft.com/product/office/wa200001603)
- [<span data-ttu-id="2e0ad-156">Script Lab の詳細情報</span><span class="sxs-lookup"><span data-stu-id="2e0ad-156">Learn more about Script Lab</span></span>](https://github.com/OfficeDev/script-lab#script-lab-a-microsoft-garage-project)
- [<span data-ttu-id="2e0ad-157">Microsoft 365 開発者プログラムに参加する</span><span class="sxs-lookup"><span data-stu-id="2e0ad-157">Join the Microsoft 365 developer program</span></span>](https://developer.microsoft.com/office/dev-program)
- [<span data-ttu-id="2e0ad-158">Office アドインを開発する</span><span class="sxs-lookup"><span data-stu-id="2e0ad-158">Developing Office Add-ins</span></span>](../develop/develop-overview.md)
- [<span data-ttu-id="2e0ad-159">Microsoft 365 開発者プログラムについて</span><span class="sxs-lookup"><span data-stu-id="2e0ad-159">Learn about the Microsoft 365 Developer Program</span></span>](https://developer.microsoft.com/microsoft-365/dev-program)
