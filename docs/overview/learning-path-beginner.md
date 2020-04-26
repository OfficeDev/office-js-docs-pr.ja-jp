---
title: ここから開始! 初心者向け Office アドイン開発ガイド
description: Office アドインの学習リソースを使用する初心者向け推奨パス。
ms.date: 04/16/2020
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: 026f90ea62960cbbf5ab4420d40a4a9165139cae
ms.sourcegitcommit: 803587b324fc8038721709d7db5664025cf03c6b
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/17/2020
ms.locfileid: "43547622"
---
# <a name="start-here-a-guide-for-beginners-making-office-add-ins"></a><span data-ttu-id="ae88b-104">ここから開始!</span><span class="sxs-lookup"><span data-stu-id="ae88b-104">Start Here!</span></span> <span data-ttu-id="ae88b-105">初心者向け Office アドイン開発ガイド</span><span class="sxs-lookup"><span data-stu-id="ae88b-105">A guide for beginners making Office Add-ins</span></span>

<span data-ttu-id="ae88b-106">独自のクロスプラットフォーム Office 拡張機能を構築する必要がありますか?</span><span class="sxs-lookup"><span data-stu-id="ae88b-106">Want to get started building your own cross-platform Office extensions?</span></span> <span data-ttu-id="ae88b-107">次の手順では、最初に読むべきこと、インストールするツール、完了すべき推奨チュートリアルを示します。</span><span class="sxs-lookup"><span data-stu-id="ae88b-107">The following steps show you what to read first, what tools to install, and recommended tutorials to complete.</span></span>

## <a name="step-0-prerequisites"></a><span data-ttu-id="ae88b-108">手順 0: 前提条件</span><span class="sxs-lookup"><span data-stu-id="ae88b-108">Step 0: Prerequisites</span></span>

- <span data-ttu-id="ae88b-109">Office アドインは、Office に組み込まれている基本 Web アプリケーションです。</span><span class="sxs-lookup"><span data-stu-id="ae88b-109">Office Add-ins are essentially web applications embedded in Office.</span></span> <span data-ttu-id="ae88b-110">まず、Web アプリケーションの基本について説明し、次に、Web でのホスト方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="ae88b-110">So, you should first have a basic understanding of web applications and how they are hosted on the web.</span></span> <span data-ttu-id="ae88b-111">インターネット、書籍、オンライン コースにこれに関する膨大な情報があります。</span><span class="sxs-lookup"><span data-stu-id="ae88b-111">There is an enormous amount of information about this on the Internet, in books, and in online courses.</span></span> <span data-ttu-id="ae88b-112">Web アプリケーションに関する事前知識がまったくない場合、Bing で "Web アプリとは?" と検索することから始めることを</span><span class="sxs-lookup"><span data-stu-id="ae88b-112">A good way to start if you have no prior knowledge of web applications at all is to search for "What is a web app?"</span></span> <span data-ttu-id="ae88b-113">お勧めします。</span><span class="sxs-lookup"><span data-stu-id="ae88b-113">on Bing.</span></span>
- <span data-ttu-id="ae88b-114">Office アドインの作成に使用する主要なプログラミング言語は、JavaScript または TypeScript です。</span><span class="sxs-lookup"><span data-stu-id="ae88b-114">The primary programming language you will use in creating Office Add-ins is JavaScript or TypeScript.</span></span> <span data-ttu-id="ae88b-115">TypeScript は、JavaScript の厳密に型指定されたバージョンと考えることができます。</span><span class="sxs-lookup"><span data-stu-id="ae88b-115">You can think of TypeScript as a strongly-typed version of JavaScript.</span></span> <span data-ttu-id="ae88b-116">これらの言語のいずれにも慣れておらず、VBA、VB.Net、C# の経験がある場合、TypeScript から学習することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="ae88b-116">If you are not familiar with either of these languages, but you have experience with VBA, VB.Net, C#, you will probably find TypeScript easier to learn.</span></span> <span data-ttu-id="ae88b-117">繰り返しになりますが、インターネット、書籍、オンライン コースに、これらの言語に関する豊富な情報があります。</span><span class="sxs-lookup"><span data-stu-id="ae88b-117">Again, there is a wealth of information about these languages on the Internet, in books, and in online courses.</span></span>

## <a name="step-1-begin-with-fundamentals"></a><span data-ttu-id="ae88b-118">手順 1: 基本から始める</span><span class="sxs-lookup"><span data-stu-id="ae88b-118">Step 1: Begin with fundamentals</span></span>

<span data-ttu-id="ae88b-119">今にもコーディングを始めたいと考えておられるかもしれませんが、IDE やコード エディターを開く前に、Office アドインについて、以下をお読みください。</span><span class="sxs-lookup"><span data-stu-id="ae88b-119">We know you're eager to start coding, but there are some things about Office Add-ins that you should read before you open your IDE or code editor.</span></span>

- <span data-ttu-id="ae88b-120">[Office アドイン プラットフォームの概要](office-add-ins.md): Office Web アドインとは何であるか、VSTO アドインなどの Office を拡張する以前の方法との違いを説明します。</span><span class="sxs-lookup"><span data-stu-id="ae88b-120">[Office Add-ins Platform Overview](office-add-ins.md): Find out what Office Web Add-ins are and how they differ from older ways of extending Office, such as VSTO add-ins.</span></span>
- <span data-ttu-id="ae88b-121">[Office アドインの構築](office-add-ins-fundamentals.md): ツール、アドイン UI の作成、JavaScript API を使用する Office ドキュメントの操作を含む、Office アドインの開発とライフサイクルの概要を説明します。</span><span class="sxs-lookup"><span data-stu-id="ae88b-121">[Building Office Add-ins](office-add-ins-fundamentals.md): Get an overview of Office add-in development and lifecycle including tooling, creating an add-in UI, and using the JavaScript APIs to interact with the Office document.</span></span>

<span data-ttu-id="ae88b-122">これらの記事には多くのリンクが含まれていますが、初心者が Office アドインを使用している場合は、これらを読み、次のセクションに進むときに、ここに戻ることをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="ae88b-122">There are a lot of links in those articles, but if you're a beginner with Office Add-ins, we recommend that you come back here when you've read them and continue with the next section.</span></span>

## <a name="step-2-install-tools-and-create-your-first-add-in"></a><span data-ttu-id="ae88b-123">手順 2: ツールをインストールし、最初のアドインを作成する</span><span class="sxs-lookup"><span data-stu-id="ae88b-123">Step 2: Install tools and create your first add-in</span></span>

<span data-ttu-id="ae88b-124">全体像を把握できたので、クイック スタートのいずれかを行います。</span><span class="sxs-lookup"><span data-stu-id="ae88b-124">You've got the big picture now, so dive in with one of our quick starts.</span></span> <span data-ttu-id="ae88b-125">プラットフォームについて学習する場合は、Excel クイック スタートをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="ae88b-125">For purposes of learning the platform, we recommend the Excel quick start.</span></span> <span data-ttu-id="ae88b-126">Visual Studio に基づくバージョンがあります。また、node.js と Visual Studio Code に基づくバージョンがあります。</span><span class="sxs-lookup"><span data-stu-id="ae88b-126">There is a version that is based on Visual Studio and a version that is based in Node.js and Visual Studio Code.</span></span>

- [<span data-ttu-id="ae88b-127">Visual Studio</span><span class="sxs-lookup"><span data-stu-id="ae88b-127">Visual Studio</span></span>](../quickstarts/excel-quickstart-jquery.md?tabs=visualstudio)
- [<span data-ttu-id="ae88b-128">Node.js および Visual Studio Code</span><span class="sxs-lookup"><span data-stu-id="ae88b-128">Node.js and Visual Studio Code</span></span>](../quickstarts/excel-quickstart-jquery.md?tabs=yeomangenerator)

## <a name="step-3-code"></a><span data-ttu-id="ae88b-129">手順 3: コーディング</span><span class="sxs-lookup"><span data-stu-id="ae88b-129">Step 3: Code</span></span>

<span data-ttu-id="ae88b-130">オーナーズ マニュアルを読んでも、理解することはできません。この [ Excel チュートリアル](../tutorials/excel-tutorial.md)を使用してコーディングを開始してください。</span><span class="sxs-lookup"><span data-stu-id="ae88b-130">You can't learn to drive by reading the owner's manual, so start coding with this [Excel tutorial](../tutorials/excel-tutorial.md).</span></span> <span data-ttu-id="ae88b-131">Office JavaScript ライブラリとアドインのマニフェストにある一部の XML を使用します。</span><span class="sxs-lookup"><span data-stu-id="ae88b-131">You'll be using the Office JavaScript library and some XML in the add-in's manifest.</span></span> <span data-ttu-id="ae88b-132">後の手順において、両方の背景がわかりやすくなっているため、何も記憶する必要はありません。</span><span class="sxs-lookup"><span data-stu-id="ae88b-132">There's no need to memorize anything, because you'll be getting more background about both in a later steps.</span></span>

## <a name="step-4-understand-the-javascript-library"></a><span data-ttu-id="ae88b-133">手順 4: JavaScript ライブラリを理解する</span><span class="sxs-lookup"><span data-stu-id="ae88b-133">Step 4: Understand the JavaScript library</span></span>

<span data-ttu-id="ae88b-134">最初に、Microsoft Learn「[Office JavaScript API について理解する](https://docs.microsoft.com/learn/modules/understand-office-javascript-apis/index)」のこのチュートリアルで、Microsoft Learn ライブラリの全体像を把握します。</span><span class="sxs-lookup"><span data-stu-id="ae88b-134">First, get the big picture of the Office JavaScript library with this tutorial from Microsoft Learn: [Understand the Office JavaScript APIs](https://docs.microsoft.com/learn/modules/understand-office-javascript-apis/index).</span></span>

<span data-ttu-id="ae88b-135">次に、API を実行して調査するサンドボックスである [Script Lab ツール](explore-with-script-lab.md)を使用して、Office JavaScript API を学習します。</span><span class="sxs-lookup"><span data-stu-id="ae88b-135">Then explore the Office JavaScript APIs with our [the Script Lab tool](explore-with-script-lab.md) -- a sandbox for running and exploring the APIs.</span></span>

## <a name="step-5-understand-the-manifest"></a><span data-ttu-id="ae88b-136">手順 5: マニフェストを理解する</span><span class="sxs-lookup"><span data-stu-id="ae88b-136">Step 5: Understand the manifest</span></span>

<span data-ttu-id="ae88b-137">アドイン マニフェストの目的を理解し、[Office アドイン XML マニフェスト](../develop/add-in-manifests.md)の XML マークアップの概要を理解します。</span><span class="sxs-lookup"><span data-stu-id="ae88b-137">Get an understanding of the purposes of the add-in manifest and an introduction to its XML markup in [Office Add-ins XML manifest](../develop/add-in-manifests.md).</span></span>

## <a name="next-steps"></a><span data-ttu-id="ae88b-138">次の手順</span><span class="sxs-lookup"><span data-stu-id="ae88b-138">Next Steps</span></span>

<span data-ttu-id="ae88b-139">おめでとうございます。 Office アドインの初心向けラーニング パスを完了しました!</span><span class="sxs-lookup"><span data-stu-id="ae88b-139">Congratulations on finishing the beginner's learning path for Office Add-ins!</span></span> <span data-ttu-id="ae88b-140">ドキュメントの詳細については、以下をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="ae88b-140">Here are some suggestions for further exploration of our documentation:</span></span>

- <span data-ttu-id="ae88b-141">その他の Office アプリケーション向けのチュートリアルおよびクイック スタート:</span><span class="sxs-lookup"><span data-stu-id="ae88b-141">Tutorials or quick starts for other Office applications:</span></span>

  - [<span data-ttu-id="ae88b-142">OneNote クイック スタート</span><span class="sxs-lookup"><span data-stu-id="ae88b-142">OneNote quick start</span></span>](../quickstarts/onenote-quickstart.md)
  - [<span data-ttu-id="ae88b-143">Outlook チュートリアル</span><span class="sxs-lookup"><span data-stu-id="ae88b-143">Outlook tutorial</span></span>](/outlook/add-ins/addin-tutorial)
  - [<span data-ttu-id="ae88b-144">PowerPoint チュートリアル</span><span class="sxs-lookup"><span data-stu-id="ae88b-144">PowerPoint tutorial</span></span>](../tutorials/powerpoint-tutorial.md)
  - [<span data-ttu-id="ae88b-145">Project クイック スタート</span><span class="sxs-lookup"><span data-stu-id="ae88b-145">Project quick start</span></span>](../quickstarts/project-quickstart.md)
  - [<span data-ttu-id="ae88b-146">Word チュートリアル</span><span class="sxs-lookup"><span data-stu-id="ae88b-146">Word tutorial</span></span>](../tutorials/word-tutorial.md)

- <span data-ttu-id="ae88b-147">その他の重要な主題:</span><span class="sxs-lookup"><span data-stu-id="ae88b-147">Other important subjects:</span></span>

  - [<span data-ttu-id="ae88b-148">Office アドインを開発する</span><span class="sxs-lookup"><span data-stu-id="ae88b-148">Develop Office Add-ins</span></span>](../develop/develop-overview.md)
  - [<span data-ttu-id="ae88b-149">Office アドイン開発のベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="ae88b-149">Best practices for developing Office Add-ins</span></span>](../concepts/add-in-development-best-practices.md)
  - [<span data-ttu-id="ae88b-150">Office アドインを設計する</span><span class="sxs-lookup"><span data-stu-id="ae88b-150">Design Office Add-ins</span></span>](../design/add-in-design.md)
  - [<span data-ttu-id="ae88b-151">Office アドインのテストとデバッグ</span><span class="sxs-lookup"><span data-stu-id="ae88b-151">Test and debug Office Add-ins</span></span>](../testing/test-debug-office-add-ins.md)
  - [<span data-ttu-id="ae88b-152">Office アドインを展開し、発行する</span><span class="sxs-lookup"><span data-stu-id="ae88b-152">Deploy and publish Office Add-ins</span></span>](../publish/publish.md)
  - [<span data-ttu-id="ae88b-153">リソース</span><span class="sxs-lookup"><span data-stu-id="ae88b-153">Resources</span></span>](../resources/resources-links-help.md)
