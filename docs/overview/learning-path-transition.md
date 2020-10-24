---
title: VSTO アドイン開発者向けガイド
description: 熟練した VSTO アドイン開発者に向けた Office Web アドイン学習リソースへの推奨パス。
ms.date: 10/14/2020
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: 1dca15a4d286e3bfa5b7ba4a502bb9161bf3257f
ms.sourcegitcommit: 42e6cfe51d99d4f3f05a3245829d764b28c46bbb
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/23/2020
ms.locfileid: "48741065"
---
# <a name="vsto-add-in-developers-guide"></a><span data-ttu-id="42169-103">VSTO アドイン開発者向けガイド</span><span class="sxs-lookup"><span data-stu-id="42169-103">VSTO add-in developer's guide</span></span>

<span data-ttu-id="42169-104">Windows で動作する Office アプリケーション用の VSTO アドインを作成しました。そしてここからは、Office を Windows、Mac、オンライン バージョンの Office スイートで動作するように拡張するための新しい方法である Office Web アドインについて説明します。</span><span class="sxs-lookup"><span data-stu-id="42169-104">So, you've made some VSTO add-ins for Office applications that run on Windows and now you're exploring the new way of extending Office that will run on Windows, Mac, and the online version of the Office suite: Office Web Add-ins.</span></span>

<span data-ttu-id="42169-105">Office Web アドインのオブジェクト モデルは Excel、Word、その他の Office アプリケーションのオブジェクト モデルと似たようなパターンをたどるので、それらのオブジェクト モデルへの理解が大きな助けとなるでしょう。</span><span class="sxs-lookup"><span data-stu-id="42169-105">Your understanding of the object models for the Excel, Word, and the other Office applications will be a huge help because the object models in Office Web Add-ins follow similar patterns.</span></span> <span data-ttu-id="42169-106">ただし、いくつか課題があります。</span><span class="sxs-lookup"><span data-stu-id="42169-106">But there are going to be some challenges:</span></span>

- <span data-ttu-id="42169-107">C# や Visual Basic .NET ではなく、別の言語 (JavaScript または TypeScript のいずれか) を使用して作業することになります。</span><span class="sxs-lookup"><span data-stu-id="42169-107">You will be working with a different language (either JavaScript or TypeScript) instead of C# or Visual Basic .NET.</span></span> <span data-ttu-id="42169-108">(後述するように、既存のコードの一部を Web アドインで再利用する方法もあります)。</span><span class="sxs-lookup"><span data-stu-id="42169-108">(There is also a way, described below, to reuse some of your existing code in a web add-in.)</span></span>
- <span data-ttu-id="42169-109">Office Web アドインは、VSTO アドインとは別に展開されます。</span><span class="sxs-lookup"><span data-stu-id="42169-109">Office Web Add-ins are deployed differently from VSTO add-ins.</span></span>
- <span data-ttu-id="42169-110">Office Web アドインは、Office アプリケーションに組み込まれた簡易ブラウザー ウィンドウで動作する Web アプリケーションなので、Web アプリケーションの基本的な理解と、それらがどのように Web サーバーやクラウド アカウントでホストされるかについてを理解しておく必要があります。</span><span class="sxs-lookup"><span data-stu-id="42169-110">Office Web Add-ins are web applications that run in a simplified browser window that is embedded in the Office application, so you need to gain a basic understanding of web applications and how they are hosted on web servers or cloud accounts.</span></span> 

<span data-ttu-id="42169-111">これらの理由により、この記事の多くは、Office 拡張機能の完全な初心者向けの学習パスである、[初心者向けガイド](learning-path-beginner.md) と重複しています。</span><span class="sxs-lookup"><span data-stu-id="42169-111">For these reasons, much of this article duplicates our learning path for complete beginners to Office extensions: [Beginner's guide](learning-path-beginner.md).</span></span> <span data-ttu-id="42169-112">VSTO アドイン開発者が経験を活用し、既存のコードを再利用できるようにするための追加の学習リソースを追加しました。</span><span class="sxs-lookup"><span data-stu-id="42169-112">What we have added are some additional learning resources to help VSTO add-in developers leverage their experience, and also help them reuse their existing code.</span></span>

## <a name="step-0-prerequisites"></a><span data-ttu-id="42169-113">手順 0: 前提条件</span><span class="sxs-lookup"><span data-stu-id="42169-113">Step 0: Prerequisites</span></span>

- <span data-ttu-id="42169-114">Office Web アドイン (Office アドインとも呼ばれる) は、Office に組み込まれている基本 Web アプリケーションです。</span><span class="sxs-lookup"><span data-stu-id="42169-114">Office Web Add-ins (also referred to as Office Add-ins) are essentially web applications embedded in Office.</span></span> <span data-ttu-id="42169-115">まず、Web アプリケーションの基本について説明し、次に、Web でのホスト方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="42169-115">So, you should first have a basic understanding of web applications and how they are hosted on the web.</span></span> <span data-ttu-id="42169-116">インターネット、書籍、オンライン コース上にこれについての膨大な情報があります。</span><span class="sxs-lookup"><span data-stu-id="42169-116">There's an enormous amount of information about this on the Internet, in books, and in online courses.</span></span> <span data-ttu-id="42169-117">Web アプリケーションに関する事前知識がまったくない場合、Bing で "Web アプリとは?" と検索することから始めることを</span><span class="sxs-lookup"><span data-stu-id="42169-117">A good way to start if you have no prior knowledge of web applications at all is to search for "What is a web app?"</span></span> <span data-ttu-id="42169-118">お勧めします。</span><span class="sxs-lookup"><span data-stu-id="42169-118">on Bing.</span></span>
- <span data-ttu-id="42169-119">Office アドインの作成に使用する主要なプログラミング言語は、JavaScript または TypeScript です。</span><span class="sxs-lookup"><span data-stu-id="42169-119">The primary programming language you'll use to create Office Add-ins is JavaScript or TypeScript.</span></span> <span data-ttu-id="42169-120">TypeScript は、JavaScript の厳密に型指定されたバージョンと考えることができます。</span><span class="sxs-lookup"><span data-stu-id="42169-120">You can think of TypeScript as a strongly-typed version of JavaScript.</span></span> <span data-ttu-id="42169-121">これらの言語のいずれにも慣れておらず、VBA、VB.Net、C# の経験がある場合、TypeScript の方が学習しやすいかもしれません。</span><span class="sxs-lookup"><span data-stu-id="42169-121">If you're not familiar with either of these languages, but you have experience with VBA, VB.Net, C#, you'll probably find TypeScript easier to learn.</span></span> <span data-ttu-id="42169-122">繰り返しになりますが、インターネット、書籍、オンライン コース上に、これらの言語に関する豊富な情報があります。</span><span class="sxs-lookup"><span data-stu-id="42169-122">Again, there's a wealth of information about these languages on the Internet, in books, and in online courses.</span></span>

## <a name="step-1-begin-with-fundamentals"></a><span data-ttu-id="42169-123">手順 1: 基本から始める</span><span class="sxs-lookup"><span data-stu-id="42169-123">Step 1: Begin with fundamentals</span></span>

<span data-ttu-id="42169-124">今にもコーディングを始めたいと考えておられるかもしれませんが、IDE やコード エディターを開く前に、Office アドインについて、以下をお読みください。</span><span class="sxs-lookup"><span data-stu-id="42169-124">We know you're eager to start coding, but there are some things about Office Add-ins that you should read before you open your IDE or code editor.</span></span>

- <span data-ttu-id="42169-125">[Office アドイン プラットフォームの概要](office-add-ins.md): Office Web アドインとは何であるか、VSTO アドインなどの Office を拡張する以前の方法との違いを説明します。</span><span class="sxs-lookup"><span data-stu-id="42169-125">[Office Add-ins Platform Overview](office-add-ins.md): Find out what Office Web Add-ins are and how they differ from older ways of extending Office, such as VSTO add-ins.</span></span>
- <span data-ttu-id="42169-126">[Office アドインの開発](../develop/develop-overview.md): ツール、アドイン UI の作成、JavaScript API を使用する Office ドキュメントの操作を含む、Office アドインの開発とライフサイクルの概要を説明します。</span><span class="sxs-lookup"><span data-stu-id="42169-126">[Develop Office Add-ins](../develop/develop-overview.md): Get an overview of Office add-in development and lifecycle including tooling, creating an add-in UI, and using the JavaScript APIs to interact with the Office document.</span></span>

<span data-ttu-id="42169-127">これらの記事には多くのリンクが含まれていますが、Office アドインに移行している場合は、これらを読み、次のセクションに進むときに、ここに戻ることをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="42169-127">There are a lot of links in those articles, but if you're transitioning to Office Web Add-ins, we recommend that you come back here when you've read them and continue with the next section.</span></span>

## <a name="step-2-install-tools-and-create-your-first-add-in"></a><span data-ttu-id="42169-128">手順 2: ツールをインストールし、最初のアドインを作成する</span><span class="sxs-lookup"><span data-stu-id="42169-128">Step 2: Install tools and create your first add-in</span></span>

<span data-ttu-id="42169-129">全体像を把握できたので、クイック スタートのいずれかを行います。</span><span class="sxs-lookup"><span data-stu-id="42169-129">You've got the big picture now, so dive in with one of our quick starts.</span></span> <span data-ttu-id="42169-130">プラットフォームについて学習する場合は、Excel クイック スタートをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="42169-130">For purposes of learning the platform, we recommend the Excel quick start.</span></span> <span data-ttu-id="42169-131">Visual Studio をベースにしたバージョンと、Node.js と Visual Studio Code をベースにしたバージョンがあります。</span><span class="sxs-lookup"><span data-stu-id="42169-131">There's a version based on Visual Studio and another based on Node.js and Visual Studio Code.</span></span> <span data-ttu-id="42169-132">VSTO アドインから移行している場合は、Visual Studio バージョンの方が作業がしやすいかもしれません。</span><span class="sxs-lookup"><span data-stu-id="42169-132">If you're transitioning from VSTO add-ins, you'll probably find the Visual Studio version easier to work with.</span></span>

- [<span data-ttu-id="42169-133">Visual Studio</span><span class="sxs-lookup"><span data-stu-id="42169-133">Visual Studio</span></span>](../quickstarts/excel-quickstart-jquery.md?tabs=visualstudio)
- [<span data-ttu-id="42169-134">Node.js および Visual Studio Code</span><span class="sxs-lookup"><span data-stu-id="42169-134">Node.js and Visual Studio Code</span></span>](../quickstarts/excel-quickstart-jquery.md?tabs=yeomangenerator)

## <a name="step-3-code"></a><span data-ttu-id="42169-135">手順 3: コーディング</span><span class="sxs-lookup"><span data-stu-id="42169-135">Step 3: Code</span></span>

<span data-ttu-id="42169-136">オーナーズ マニュアルを読んでも、理解することはできません。この [ Excel チュートリアル](../tutorials/excel-tutorial.md)を使用してコーディングを開始してください。</span><span class="sxs-lookup"><span data-stu-id="42169-136">You can't learn to drive by reading the owner's manual, so start coding with this [Excel tutorial](../tutorials/excel-tutorial.md).</span></span> <span data-ttu-id="42169-137">Office JavaScript ライブラリとアドインのマニフェストにある一部の XML を使用します。</span><span class="sxs-lookup"><span data-stu-id="42169-137">You'll be using the Office JavaScript library and some XML in the add-in's manifest.</span></span> <span data-ttu-id="42169-138">後の手順で両方の背景を学習することになるので、何も覚える必要はありません。</span><span class="sxs-lookup"><span data-stu-id="42169-138">There's no need to memorize anything, because you'll be getting more background about both in a later step.</span></span>

## <a name="step-4-understand-the-javascript-library"></a><span data-ttu-id="42169-139">手順 4: JavaScript ライブラリを理解する</span><span class="sxs-lookup"><span data-stu-id="42169-139">Step 4: Understand the JavaScript library</span></span>

<span data-ttu-id="42169-140">Microsoft Learn「[Office JavaScript API について理解する](/learn/modules/intro-office-add-ins/3-apis)」のこのチュートリアルで、Office JavaScript ライブラリの全体像を把握します。</span><span class="sxs-lookup"><span data-stu-id="42169-140">Get the big picture of the Office JavaScript library with this tutorial from Microsoft Learn: [Understand the Office JavaScript APIs](/learn/modules/intro-office-add-ins/3-apis).</span></span>

<span data-ttu-id="42169-141">次に、API を実行して調査するためのサンドボックスである [Script Lab ツール](explore-with-script-lab.md)を使用して、Office JavaScript API を学習します。</span><span class="sxs-lookup"><span data-stu-id="42169-141">Then explore the Office JavaScript APIs with the [Script Lab tool](explore-with-script-lab.md) -- a sandbox for running and exploring the APIs.</span></span>

### <a name="special-resource-for-vsto-add-in-developers"></a><span data-ttu-id="42169-142">VSTO アドイン開発者向けの特別なリソース</span><span class="sxs-lookup"><span data-stu-id="42169-142">Special resource for VSTO add-in developers</span></span>

<span data-ttu-id="42169-143">サンプルのアドインを見るには、[Excel アドイン JavaScript SalesTracker](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker) が良いでしょう。</span><span class="sxs-lookup"><span data-stu-id="42169-143">This would be a good place to take a look at the sample add-in, [Excel Add-in JavaScript SalesTracker](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker).</span></span> <span data-ttu-id="42169-144">これは VSTO アドインと Office Web アドインの共通点や違いを強調するために作成されたもので、このサンプルの readme では比較する上での重要なポイントについてを紹介しています。</span><span class="sxs-lookup"><span data-stu-id="42169-144">It was created to highlight the similarities and differences between VSTO add-ins and Office Web Add-ins, and the readme of the sample calls out the important points of comparison.</span></span>

## <a name="step-5-understand-the-manifest"></a><span data-ttu-id="42169-145">手順 5: マニフェストを理解する</span><span class="sxs-lookup"><span data-stu-id="42169-145">Step 5: Understand the manifest</span></span>

<span data-ttu-id="42169-146">Web アドイン マニフェストの目的を理解し、[Office アドインの XML マニフェスト](../develop/add-in-manifests.md)で XML マークアップの概要について説明します。</span><span class="sxs-lookup"><span data-stu-id="42169-146">Get an understanding of the purposes of the web add-in manifest and an introduction to its XML markup in [Office Add-ins XML manifest](../develop/add-in-manifests.md).</span></span>

## <a name="step-6-for-vsto-developers-only-reuse-your-vsto-code"></a><span data-ttu-id="42169-147">手順 6 (VSTO 開発者のみ): VSTO コードを再利用する</span><span class="sxs-lookup"><span data-stu-id="42169-147">Step 6 (for VSTO developers only): Reuse your VSTO code</span></span>

<span data-ttu-id="42169-148">VSTO アドインのコードをサーバー上の Web アプリケーションのバックエンドへと移動し、JavaScript や TypeScript で Web API として利用できるようにすることにより、Office Web アドインで VSTO アドインのコードを再利用することができます。</span><span class="sxs-lookup"><span data-stu-id="42169-148">You can reuse some of your VSTO add-in code in an Office web add-in by moving it to your web application's back end on the server and making it available to your JavaScript or TypeScript as a web API.</span></span> <span data-ttu-id="42169-149">ガイダンスについては、「[チュートリアル: 共有コード ライブラリを使用して VSTO アドインと Office アドインの間でコードを共有する](../tutorials/migrate-vsto-to-office-add-in-shared-code-library-tutorial.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="42169-149">For guidance, see [Tutorial: Share code between both a VSTO Add-in and an Office add-in by using a shared code library](../tutorials/migrate-vsto-to-office-add-in-shared-code-library-tutorial.md).</span></span>

## <a name="next-steps"></a><span data-ttu-id="42169-150">次の手順</span><span class="sxs-lookup"><span data-stu-id="42169-150">Next Steps</span></span>

<span data-ttu-id="42169-151">おめでとうございます。Office Web アドインの VSTO アドイン開発者向け学習パスを完了しました!</span><span class="sxs-lookup"><span data-stu-id="42169-151">Congratulations on finishing the VSTO add-in developer's learning path for Office Web Add-ins!</span></span> <span data-ttu-id="42169-152">ドキュメントの詳細については、以下をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="42169-152">Here are some suggestions for further exploration of our documentation:</span></span>

- <span data-ttu-id="42169-153">その他の Office アプリケーション向けのチュートリアルおよびクイック スタート:</span><span class="sxs-lookup"><span data-stu-id="42169-153">Tutorials or quick starts for other Office applications:</span></span>

  - [<span data-ttu-id="42169-154">OneNote クイック スタート</span><span class="sxs-lookup"><span data-stu-id="42169-154">OneNote quick start</span></span>](../quickstarts/onenote-quickstart.md)
  - [<span data-ttu-id="42169-155">Outlook チュートリアル</span><span class="sxs-lookup"><span data-stu-id="42169-155">Outlook tutorial</span></span>](/outlook/add-ins/addin-tutorial)
  - [<span data-ttu-id="42169-156">PowerPoint チュートリアル</span><span class="sxs-lookup"><span data-stu-id="42169-156">PowerPoint tutorial</span></span>](../tutorials/powerpoint-tutorial.md)
  - [<span data-ttu-id="42169-157">Project クイック スタート</span><span class="sxs-lookup"><span data-stu-id="42169-157">Project quick start</span></span>](../quickstarts/project-quickstart.md)
  - [<span data-ttu-id="42169-158">Word チュートリアル</span><span class="sxs-lookup"><span data-stu-id="42169-158">Word tutorial</span></span>](../tutorials/word-tutorial.md)

- <span data-ttu-id="42169-159">その他の重要な主題:</span><span class="sxs-lookup"><span data-stu-id="42169-159">Other important subjects:</span></span>

  - [<span data-ttu-id="42169-160">Office アドインを開発する</span><span class="sxs-lookup"><span data-stu-id="42169-160">Develop Office Add-ins</span></span>](../develop/develop-overview.md)
  - [<span data-ttu-id="42169-161">Office アドイン開発のベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="42169-161">Best practices for developing Office Add-ins</span></span>](../concepts/add-in-development-best-practices.md)
  - [<span data-ttu-id="42169-162">Office アドインを設計する</span><span class="sxs-lookup"><span data-stu-id="42169-162">Design Office Add-ins</span></span>](../design/add-in-design.md)
  - [<span data-ttu-id="42169-163">Office アドインのテストとデバッグ</span><span class="sxs-lookup"><span data-stu-id="42169-163">Test and debug Office Add-ins</span></span>](../testing/test-debug-office-add-ins.md)
  - [<span data-ttu-id="42169-164">Office アドインを展開し、発行する</span><span class="sxs-lookup"><span data-stu-id="42169-164">Deploy and publish Office Add-ins</span></span>](../publish/publish.md)
  - [<span data-ttu-id="42169-165">Resources</span><span class="sxs-lookup"><span data-stu-id="42169-165">Resources</span></span>](../resources/resources-links-help.md)
  - [<span data-ttu-id="42169-166">Microsoft 365 開発者プログラムについて学ぶ</span><span class="sxs-lookup"><span data-stu-id="42169-166">Learn about the Microsoft 365 Developer Program</span></span>](https://developer.microsoft.com/microsoft-365/dev-program)
