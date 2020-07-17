---
title: Script Lab を使用して Office JavaScript API を探索する
description: Script Lab を使用して、Office JS API およびプロトタイプの機能を調べます。
ms.date: 06/10/2020
ms.topic: conceptual
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: ab2d086551dbfa5063615f505d8cb8aa5a210b7a
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/10/2020
ms.locfileid: "45094135"
---
# <a name="explore-office-javascript-api-using-script-lab"></a><span data-ttu-id="d71b6-103">Script Lab を使用して Office JavaScript API を探索する</span><span class="sxs-lookup"><span data-stu-id="d71b6-103">Explore Office JavaScript API using Script Lab</span></span>

<span data-ttu-id="d71b6-104">AppSource から無料で入手できる [Script Lab](https://appsource.microsoft.com/product/office/WA104380862) および [Outlook 用 Script Lab](https://appsource.microsoft.com/product/office/wa200001603) アドインを使用すると、Excel や Outlook などの Office プログラムでの作業中に Office JavaScript API を調査できます。</span><span class="sxs-lookup"><span data-stu-id="d71b6-104">The [Script Lab](https://appsource.microsoft.com/product/office/WA104380862) and [Script Lab for Outlook](https://appsource.microsoft.com/product/office/wa200001603) add-ins, available free from AppSource, enable you to explore the Office JavaScript API while you're working in an Office program such as Excel or Outlook.</span></span> <span data-ttu-id="d71b6-105">Script Lab は、独自のアドインで必要な機能のプロトタイプを作成して検証するときに、開発ツールキットに追加する便利なツールです。</span><span class="sxs-lookup"><span data-stu-id="d71b6-105">Script Lab is a convenient tool to add to your development toolkit as you prototype and verify functionality you want in your own add-in.</span></span>

## <a name="what-is-script-lab"></a><span data-ttu-id="d71b6-106">Script Lab とは</span><span class="sxs-lookup"><span data-stu-id="d71b6-106">What is Script Lab?</span></span>

<span data-ttu-id="d71b6-107">Script Lab は、Excel、Outlook、Word、PowerPoint で Office JavaScript API を使用して Office アドインを開発する方法を学習したい人のためのツールです。</span><span class="sxs-lookup"><span data-stu-id="d71b6-107">Script Lab is a tool for anyone who wants to learn how to develop Office Add-ins using the Office JavaScript API in Excel, Outlook, Word, and PowerPoint.</span></span> <span data-ttu-id="d71b6-108">IntelliSense を提供しているので、何が利用できるのかを見ることができ、Visual Studio Code で使用されているのと同じフレームワークである Monaco フレームワークの上に構築されています。</span><span class="sxs-lookup"><span data-stu-id="d71b6-108">It provides IntelliSense so you can see what's available and is built on the Monaco framework, the same framework used by Visual Studio Code.</span></span> <span data-ttu-id="d71b6-109">Script Lab では、サンプルのライブラリにアクセスして、簡単に機能を試すことができます。また、独自のコードの開始点としてサンプルを使用することもできます。</span><span class="sxs-lookup"><span data-stu-id="d71b6-109">Through Script Lab, you can access a library of samples to quickly try out features or you can use a sample as the starting point for your own code.</span></span> <span data-ttu-id="d71b6-110">Script Lab を使用して、プレビュー API を試すこともできます。</span><span class="sxs-lookup"><span data-stu-id="d71b6-110">You can even use Script Lab to try preview APIs.</span></span>

<span data-ttu-id="d71b6-111">今のところいいですか?</span><span class="sxs-lookup"><span data-stu-id="d71b6-111">Sounds good so far?</span></span> <span data-ttu-id="d71b6-112">この 1 分間のビデオを見て、Script Lab の動作を確認します。</span><span class="sxs-lookup"><span data-stu-id="d71b6-112">Take a look at this one-minute video to see Script Lab in action.</span></span>

<span data-ttu-id="d71b6-113">[![Excel、Word、PowerPoint での Script Lab の実行を紹介するプレビュー ビデオ。](../images/screenshot-wide-youtube.png 'Script Lab のプレビュー ビデオ')](https://aka.ms/scriptlabvideo)</span><span class="sxs-lookup"><span data-stu-id="d71b6-113">[![Preview video showing Script Lab running in Excel, Word, and PowerPoint.](../images/screenshot-wide-youtube.png 'Script Lab preview video')](https://aka.ms/scriptlabvideo)</span></span>

## <a name="key-features"></a><span data-ttu-id="d71b6-114">主な機能</span><span class="sxs-lookup"><span data-stu-id="d71b6-114">Key features</span></span>

<span data-ttu-id="d71b6-115">Script Lab には、Office JavaScript API およびプロトタイプ アドインの機能の調査に役立つ機能が多数用意されています。</span><span class="sxs-lookup"><span data-stu-id="d71b6-115">Script Lab offers a number of features to help you explore the Office JavaScript API and prototype add-in functionality.</span></span>

### <a name="explore-samples"></a><span data-ttu-id="d71b6-116">サンプルの確認</span><span class="sxs-lookup"><span data-stu-id="d71b6-116">Explore samples</span></span>

<span data-ttu-id="d71b6-117">API を使用してタスクを完了する方法を示す組み込みのサンプル スニペットのコレクションを使用してすぐに開始できます。</span><span class="sxs-lookup"><span data-stu-id="d71b6-117">Get started quickly with a collection of built-in sample snippets that show how to complete tasks with the API.</span></span> <span data-ttu-id="d71b6-118">サンプルを実行すると、作業ウィンドウまたはドキュメントですばやく結果を表示したり、API のしくみをサンプルで確認して学んだり、独自のアドインのプロトタイプにサンプルを使用したりもできます。</span><span class="sxs-lookup"><span data-stu-id="d71b6-118">You can run the samples to instantly see the result in the task pane or document, examine the samples to learn how the API works, and even use samples to prototype your own add-in.</span></span>

![サンプル](../images/script-lab-samples.jpg)

### <a name="code-and-style"></a><span data-ttu-id="d71b6-120">コードとスタイル</span><span class="sxs-lookup"><span data-stu-id="d71b6-120">Code and style</span></span>

<span data-ttu-id="d71b6-121">Office JS API を呼び出す JavaScript または TypeScript コードに加えて、各スニペットには、作業ウィンドウのコンテンツを定義する HTML マークアップと、作業ウィンドウの外観を定義する CSS も含まれています。</span><span class="sxs-lookup"><span data-stu-id="d71b6-121">In addition to JavaScript or TypeScript code that calls the Office JS API, each snippet also contains HTML markup that defines content of the task pane and CSS that defines the appearance of the task pane.</span></span> <span data-ttu-id="d71b6-122">HTML マークアップと CSS をカスタマイズして、独自のアドインの作業ウィンドウ デザインのプロトタイプを作成する際に、要素の配置とスタイル設定を試すことができます。</span><span class="sxs-lookup"><span data-stu-id="d71b6-122">You can customize the HTML markup and CSS to experiment with element placement and styling as you prototype task pane design for your own add-in.</span></span>

> [!TIP]
> <span data-ttu-id="d71b6-123">スニペット内でプレビュー API を呼び出すには、スニペットのライブラリを更新して、ベータ CDN (`https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`) とプレビューの種類の定義 `@types/office-js-preview` を使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="d71b6-123">To call preview APIs within a snippet, you'll need to update the snippet's libraries to use the beta CDN (`https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`) and the preview type definitions `@types/office-js-preview`.</span></span> <span data-ttu-id="d71b6-124">また、一部のプレビュー API は、[Office Insider プログラム](https://insider.office.com)にサインアップして、Insider ビルドの Office を実行している場合にのみアクセスできます。</span><span class="sxs-lookup"><span data-stu-id="d71b6-124">Additionally, some preview APIs are only accessible if you've signed up for the [Office Insider program](https://insider.office.com) and are running an Insider build of Office.</span></span>

### <a name="save-and-share-snippets"></a><span data-ttu-id="d71b6-125">スニペットの保存と共有</span><span class="sxs-lookup"><span data-stu-id="d71b6-125">Save and share snippets</span></span>

<span data-ttu-id="d71b6-126">既定では、Script Lab で開いたスニペットはブラウザーのキャッシュに保存されます。</span><span class="sxs-lookup"><span data-stu-id="d71b6-126">By default, snippets that you open in Script Lab will be saved to your browser cache.</span></span> <span data-ttu-id="d71b6-127">スニペットを完全に保存するには、そのスニペットを [GitHub の Gist](https://gist.github.com) にエクスポートします。</span><span class="sxs-lookup"><span data-stu-id="d71b6-127">To save a snippet permanently, you can export it to a [GitHub gist](https://gist.github.com).</span></span> <span data-ttu-id="d71b6-128">自分専用にスニペットを保存するには、秘密の Gist を作成するか、他のユーザーと共有する予定がある場合はパブリックの Gist を作成します。</span><span class="sxs-lookup"><span data-stu-id="d71b6-128">Create a secret gist to save a snippet exclusively for your own use, or create a public gist if you plan to share it with others.</span></span>

![共有オプション](../images/script-lab-share.jpg)

### <a name="import-snippets"></a><span data-ttu-id="d71b6-130">スニペットのインポート</span><span class="sxs-lookup"><span data-stu-id="d71b6-130">Import snippets</span></span>

<span data-ttu-id="d71b6-131">スニペット YAML が保存されているパブリック [ GitHub の Gist ](https://gist.github.com) に URL を指定するか、スニペットの完全な YAML を貼り付けて、スニペットを Script Lab にインポートできます。</span><span class="sxs-lookup"><span data-stu-id="d71b6-131">You can import a snippet into Script Lab either by specifying the URL to the public [GitHub gist](https://gist.github.com) where the snippet YAML is stored or by pasting in the complete YAML for the snippet.</span></span> <span data-ttu-id="d71b6-132">この機能は、GitHub の Gist にスニペットを公開するか、スニペットの YAML を提供すると、他のユーザーがスニペットを自分と共有しているシナリオで役立ちます。</span><span class="sxs-lookup"><span data-stu-id="d71b6-132">This feature may be useful in scenarios where someone else has shared their snippet with you by either publishing it to a GitHub gist or providing their snippet's YAML.</span></span>

![スニペットのインポート オプション](../images/script-lab-import-snippet.jpg)

## <a name="supported-clients"></a><span data-ttu-id="d71b6-134">サポートされるクライアント</span><span class="sxs-lookup"><span data-stu-id="d71b6-134">Supported clients</span></span>

<span data-ttu-id="d71b6-135">Script Lab は、次のクライアント上の Excel、Word、PowerPoint でサポートされています。</span><span class="sxs-lookup"><span data-stu-id="d71b6-135">Script Lab is supported for Excel, Word, and PowerPoint on the following clients.</span></span>

- <span data-ttu-id="d71b6-136">Windows での Office 2013 以降</span><span class="sxs-lookup"><span data-stu-id="d71b6-136">Office 2013 or later on Windows</span></span>
- <span data-ttu-id="d71b6-137">Mac での Office 2016 以降</span><span class="sxs-lookup"><span data-stu-id="d71b6-137">Office 2016 or later on Mac</span></span>
- <span data-ttu-id="d71b6-138">Office on the web</span><span class="sxs-lookup"><span data-stu-id="d71b6-138">Office on the web</span></span>

<span data-ttu-id="d71b6-139">Outlook 用 Script Lab は、次のクライアントで使用できます。</span><span class="sxs-lookup"><span data-stu-id="d71b6-139">Script Lab for Outlook is available on the following clients.</span></span>

- <span data-ttu-id="d71b6-140">Windows 用 Outlook 2013 以降</span><span class="sxs-lookup"><span data-stu-id="d71b6-140">Outlook 2013 or later on Windows</span></span>
- <span data-ttu-id="d71b6-141">Mac 用 Outlook 2016 以降</span><span class="sxs-lookup"><span data-stu-id="d71b6-141">Outlook 2016 or later on Mac</span></span>
- <span data-ttu-id="d71b6-142">Chrome、Microsoft Edge、または Safari ブラウザーを使用している場合の Outlook on the web</span><span class="sxs-lookup"><span data-stu-id="d71b6-142">Outlook on the web when using Chrome, Microsoft Edge, or Safari browsers</span></span>

<span data-ttu-id="d71b6-143">Outlook 用 Script Lab の詳細については、関連する[ブログの投稿](https://developer.microsoft.com/outlook/blogs/script-lab-now-supports-outlook/)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="d71b6-143">For more details on Script Lab for Outlook, see the related [blog post](https://developer.microsoft.com/outlook/blogs/script-lab-now-supports-outlook/).</span></span>

## <a name="next-steps"></a><span data-ttu-id="d71b6-144">次の手順</span><span class="sxs-lookup"><span data-stu-id="d71b6-144">Next steps</span></span>

<span data-ttu-id="d71b6-145">Excel、Word、または PowerPoint で Script Lab を使用するには、AppSource から [Script Lab アドイン](https://appsource.microsoft.com/product/office/WA104380862)をインストールします。</span><span class="sxs-lookup"><span data-stu-id="d71b6-145">To use Script Lab in Excel, Word, or PowerPoint, install the [Script Lab add-in](https://appsource.microsoft.com/product/office/WA104380862) from AppSource.</span></span> 

<span data-ttu-id="d71b6-146">Outlook 用 Script Lab を使用するには、AppSource から [Outlook 用 Script Lab のアドイン](https://appsource.microsoft.com/product/office/wa200001603)をインストールします。</span><span class="sxs-lookup"><span data-stu-id="d71b6-146">To use Script Lab for Outlook, install the [Script Lab for Outlook add-in](https://appsource.microsoft.com/product/office/wa200001603) from AppSource.</span></span>

<span data-ttu-id="d71b6-147">新しいスニペットを [office-js-snippets](https://github.com/OfficeDev/office-js-snippets#office-js-snippets) GitHub リポジトリに投稿し、Script Lab のサンプル ライブラリを拡張してください。</span><span class="sxs-lookup"><span data-stu-id="d71b6-147">You're welcome to expand the sample library in Script Lab by contributing new snippets to the [office-js-snippets](https://github.com/OfficeDev/office-js-snippets#office-js-snippets) GitHub repository.</span></span>

<span data-ttu-id="d71b6-148">最初の Office アドインを作成する準備ができたら、[Excel](../quickstarts/excel-quickstart-jquery.md)、[Outlook](../quickstarts/outlook-quickstart.md)、[Word](../quickstarts/word-quickstart.md)、[OneNote ](../quickstarts/onenote-quickstart.md)、[PowerPoint](../quickstarts/powerpoint-quickstart.md)、または [Project](../quickstarts/project-quickstart.md) のクイック スタートを試してください。</span><span class="sxs-lookup"><span data-stu-id="d71b6-148">When you're ready to create your first Office Add-in, try out the quick start for [Excel](../quickstarts/excel-quickstart-jquery.md), [Outlook](../quickstarts/outlook-quickstart.md), [Word](../quickstarts/word-quickstart.md), [OneNote](../quickstarts/onenote-quickstart.md), [PowerPoint](../quickstarts/powerpoint-quickstart.md), or [Project](../quickstarts/project-quickstart.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="d71b6-149">関連項目</span><span class="sxs-lookup"><span data-stu-id="d71b6-149">See also</span></span>

- [<span data-ttu-id="d71b6-150">Excel、Word、または Powerpoint 用 Script Lab を取得する</span><span class="sxs-lookup"><span data-stu-id="d71b6-150">Get Script Lab for Excel, Word, or Powerpoint</span></span>](https://appsource.microsoft.com/product/office/WA104380862)
- [<span data-ttu-id="d71b6-151">Outlook 用 Script Lab を取得する</span><span class="sxs-lookup"><span data-stu-id="d71b6-151">Get Script Lab for Outlook</span></span>](https://appsource.microsoft.com/product/office/wa200001603)
- [<span data-ttu-id="d71b6-152">Script Lab の詳細情報</span><span class="sxs-lookup"><span data-stu-id="d71b6-152">Learn more about Script Lab</span></span>](https://github.com/OfficeDev/script-lab#script-lab-a-microsoft-garage-project)
- [<span data-ttu-id="d71b6-153">Microsoft 365 開発者プログラムに参加する</span><span class="sxs-lookup"><span data-stu-id="d71b6-153">Join the Microsoft 365 developer program</span></span>](https://developer.microsoft.com/office/dev-program)
- [<span data-ttu-id="d71b6-154">Office アドインを構築する</span><span class="sxs-lookup"><span data-stu-id="d71b6-154">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)
