---
title: Visual Studio を使用して Office アドインを開発する
description: Visual Studio を使用して Office アドインを開発する方法
ms.date: 12/31/2019
localization_priority: Priority
ms.openlocfilehash: 2acafc08dcbb8e1fde439f2be289afff3185272c
ms.sourcegitcommit: d5ac9284d1e96dc91a9168d7641e44d88535e1a7
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/31/2019
ms.locfileid: "40915071"
---
# <a name="develop-office-add-ins-with-visual-studio"></a><span data-ttu-id="139a8-103">Visual Studio を使用して Office アドインを開発する</span><span class="sxs-lookup"><span data-stu-id="139a8-103">Develop Office Add-ins with Visual Studio Code</span></span>

<span data-ttu-id="139a8-104">この記事では、Visual Studio を使用して Office アドインを開発する方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="139a8-104">This article describes how to use Visual Studio Code (VS Code) to develop an Office Add-in.</span></span> <span data-ttu-id="139a8-105">アドインの作成が既に完了している場合は、「[Visual Studio を使用して アドインを開発する](#develop-the-add-in-using-visual-studio)」セクションに進んでください。</span><span class="sxs-lookup"><span data-stu-id="139a8-105">If you've already created your add-in, you can skip ahead to the [Develop the add-in using Visual Studio](#develop-the-add-in-using-visual-studio) section.</span></span>

> [!NOTE]
> <span data-ttu-id="139a8-106">Visual Studio を使用する代わりに、Office アドイン用の Yeoman ジェネレーターと VS コードを使用して Office アドインを作成することもできます。</span><span class="sxs-lookup"><span data-stu-id="139a8-106">As an alternative to using Visual Studio, you may choose to use the Yeoman generator for Office Add-ins and VS Code to create an Office Add-in.</span></span> <span data-ttu-id="139a8-107">この選択肢の詳細については、「[Office アドインの作成 ](../overview/office-add-ins-fundamentals.md#creating-an-office-add-in)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="139a8-107">For more information about this choice, see [Creating an Office Add-in](../overview/office-add-ins-fundamentals.md#creating-an-office-add-in).</span></span>

## <a name="create-the-add-in-project-using-visual-studio"></a><span data-ttu-id="139a8-108">Visual Studio を使用してアドイン プロジェクトを作成する</span><span class="sxs-lookup"><span data-stu-id="139a8-108">Create the add-in project in Visual Studio</span></span>

<span data-ttu-id="139a8-109">Visual Studio は、Excel、Outlook、Word、および PowerPoint 用の Office アドインの作成に使用できます。</span><span class="sxs-lookup"><span data-stu-id="139a8-109">Visual Studio can be used to create Office Add-ins for Excel, Word, PowerPoint, or Outlook.</span></span> <span data-ttu-id="139a8-110">Office アドイン プロジェクトは Visual Studio ソリューションの一部として作成され、HTML、CSS、および JavaScript が使用されます。</span><span class="sxs-lookup"><span data-stu-id="139a8-110">An Office Add-in project gets created as part of a Visual Studio solution and uses HTML, CSS, and JavaScript.</span></span> <span data-ttu-id="139a8-111">Visual Studio を使用して Office アドインを作成するには、作成するアドインに対応するクイック スタートの指示に従います。</span><span class="sxs-lookup"><span data-stu-id="139a8-111">To create an Office Add-in with the Yeoman generator, follow instructions in the 5-minute quick start that corresponds to the type of add-in you'd like to create.</span></span>

- [<span data-ttu-id="139a8-112">Excel クイック スタート</span><span class="sxs-lookup"><span data-stu-id="139a8-112">Excel quick start</span></span>](../quickstarts/excel-quickstart-jquery.md?tabs=visualstudio)
- [<span data-ttu-id="139a8-113">Outlook クイック スタート</span><span class="sxs-lookup"><span data-stu-id="139a8-113">Outlook quick start</span></span>](/outlook/add-ins/quick-start?context=office/dev/add-ins/context&tabs=visualstudio)
- [<span data-ttu-id="139a8-114">Word クイック スタート</span><span class="sxs-lookup"><span data-stu-id="139a8-114">Word quick start</span></span>](../quickstarts/word-quickstart.md?tabs=visualstudio)
- [<span data-ttu-id="139a8-115">PowerPoint クイック スタート</span><span class="sxs-lookup"><span data-stu-id="139a8-115">PowerPoint quick start</span></span>](../quickstarts/powerpoint-quickstart.md?tabs=visualstudio)

<span data-ttu-id="139a8-116">Visual Studio では、OneNote または Project 用の Office アドインの作成はサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="139a8-116">Visual Studio doesn't support creating Office Add-ins for OneNote or Project.</span></span> <span data-ttu-id="139a8-117">これらのいずれかのホスト用の Office アドインを作成するには、[OneNote クイック スタート](../quickstarts/onenote-quickstart.md)または [Project クイック スタート](../quickstarts/project-quickstart.md)で説明するように、Office アドイン用の Yeoman ジェネレーターを使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="139a8-117">To create Office Add-ins for either of these hosts, you'll need to use the Yeoman generator for Office Add-ins, as described in the [OneNote quick start](../quickstarts/onenote-quickstart.md) or the [Project quick start](../quickstarts/project-quickstart.md).</span></span>

## <a name="develop-the-add-in-using-visual-studio"></a><span data-ttu-id="139a8-118">Visual Studio を使用してアドインを開発する</span><span class="sxs-lookup"><span data-stu-id="139a8-118">Develop the add-in using Visual Studio</span></span>

<span data-ttu-id="139a8-119">Visual Studio では、機能が制限された基本的なアドインが作成されます。</span><span class="sxs-lookup"><span data-stu-id="139a8-119">The Yeoman generator creates a basic add-in with limited functionality.</span></span> <span data-ttu-id="139a8-120">[マニフェスト](add-in-manifests.md)、HTML、JavaScript、および CSS の各ファイルを Visual Studio で編集することで、アドインをカスタマイズできます。</span><span class="sxs-lookup"><span data-stu-id="139a8-120">You can customize the add-in by editing the [manifest](add-in-manifests.md), HTML, JavaScript or TypeScript, and CSS files in VS Code.</span></span> <span data-ttu-id="139a8-121">Visual Studio により作成されるアドイン プロジェクトのプロジェクト構造とファイルの概要については、アドインを作成するために実行したクイック スタート内の Visual Studio ガイダンスを参照してください。</span><span class="sxs-lookup"><span data-stu-id="139a8-121">For a high-level description of the project structure and files in the add-in project that the Yeoman generator creates, see the Yeoman generator guidance within the 5-minute quick start that corresponds to the type of add-in you've created.</span></span> 

> [!TIP]
> <span data-ttu-id="139a8-122">Office アドインは Web アプリケーションであるため、アドインをカスタマイズするには、少なくとも Web 開発の基本的なスキルが必要です。</span><span class="sxs-lookup"><span data-stu-id="139a8-122">Because an Office Add-in is a web application, you'll need at least basic web development skills to customize your add-in.</span></span> <span data-ttu-id="139a8-123">JavaScript を使い慣れていない場合は、「[Mozilla JavaScript チュートリアル](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction)」を確認することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="139a8-123">If you're new to JavaScript, we recommend reviewing the [Mozilla JavaScript tutorial](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).</span></span>

<span data-ttu-id="139a8-124">アドインをカスタマイズするには、このドキュメントの[「中心概念」 > 「開発」](develop-overview.md)項目で説明する概念の他、作成するるアドインに対応するドキュメント内の、ホスト固有の項目 (例: [Excel](../excel/index.md)) で説明する概念を理解する必要があります。</span><span class="sxs-lookup"><span data-stu-id="139a8-124">To customize your add-in, you'll need to understand concepts described in the [Core concepts > Develop](develop-overview.md) area of this documentation, as well as concepts described in the host-specific area of documentation that corresponds to the add-in you're building (for example, [Excel](../excel/index.md)).</span></span> 

## <a name="test-and-debug-the-add-in"></a><span data-ttu-id="139a8-125">アドインのテストとデバッグ</span><span class="sxs-lookup"><span data-stu-id="139a8-125">Test and debug the add-in</span></span>

<span data-ttu-id="139a8-126">Office アドインのテスト、デバッグ、およびトラブルシューティングの方法は、プラットフォームによって異なります。</span><span class="sxs-lookup"><span data-stu-id="139a8-126">Methods for testing, debugging, and troubleshooting Office Add-ins vary by platform.</span></span> <span data-ttu-id="139a8-127">詳細については、「[Visual Studio で Office アドインをデバッグする](debug-office-add-ins-in-visual-studio.md)」および「[Office アドインのテストとデバッグ](../testing/test-debug-office-add-ins.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="139a8-127">For more information, see [Debug Office Add-ins in Visual Studio](debug-office-add-ins-in-visual-studio.md) and [Test and debug Office Add-ins](../testing/test-debug-office-add-ins.md).</span></span>

## <a name="publish-the-add-in"></a><span data-ttu-id="139a8-128">アドインを発行する</span><span class="sxs-lookup"><span data-stu-id="139a8-128">Publish the add-in</span></span>

<span data-ttu-id="139a8-129">Office アドインは、Web アプリケーションとマニフェスト ファイルで構成されています。</span><span class="sxs-lookup"><span data-stu-id="139a8-129">An Office Add-in consists of a web application and a manifest file.</span></span> <span data-ttu-id="139a8-130">Web アプリケーションはアドインのユーザー インターフェイスと機能を定義し、マニフェストは Web アプリケーションの場所を指定し、アドインの設定と機能を定義します。</span><span class="sxs-lookup"><span data-stu-id="139a8-130">The web application defines the add-in's user interface and functionality, while the manifest specifies the location of the web application and defines settings and capabilities of the add-in.</span></span>

<span data-ttu-id="139a8-131">Visual Studio で開発中のアドインは、ローカル Web サーバー上 (`localhost`) で実行されます。</span><span class="sxs-lookup"><span data-stu-id="139a8-131">While you're developing your add-in in Visual Studio, your add-in runs on your local web server (`localhost`).</span></span> <span data-ttu-id="139a8-132">アドインが正常に機能し、他のユーザーがアクセスできるように公開する準備ができた場合、次の手順を実行する必要があります。</span><span class="sxs-lookup"><span data-stu-id="139a8-132">When your add-in is working as desired and you're ready to publish it for other users to access, complete the following steps:</span></span>

1. <span data-ttu-id="139a8-133">Web アプリケーションを Web サーバーまたは Web ホスティング サービス (例: Microsoft Azure) に展開します。</span><span class="sxs-lookup"><span data-stu-id="139a8-133">Deploy the web application to a web server or web hosting service (for example, Microsoft Azure).</span></span>
2. <span data-ttu-id="139a8-134">マニフェストを更新して、展開されたアプリケーションの URL を指定します。</span><span class="sxs-lookup"><span data-stu-id="139a8-134">Update the manifest to specify the URL of the deployed application.</span></span> 
3. <span data-ttu-id="139a8-135">[Office アドインを展開](../publish/publish.md)するために使用する方法を選択し、指示に従ってマニフェスト ファイルを公開します。</span><span class="sxs-lookup"><span data-stu-id="139a8-135">Choose the method you'd like to use to [deploy and publish your Office Add-in](../publish/publish.md) your add-in, and follow the instructions to publish the manifest file.</span></span>

## <a name="see-also"></a><span data-ttu-id="139a8-136">関連項目</span><span class="sxs-lookup"><span data-stu-id="139a8-136">See also</span></span>

- [<span data-ttu-id="139a8-137">Office アドインを構築する</span><span class="sxs-lookup"><span data-stu-id="139a8-137">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)
- [<span data-ttu-id="139a8-138">Office アドインの中心概念</span><span class="sxs-lookup"><span data-stu-id="139a8-138">Core concepts for Office Add-ins</span></span>](../overview/core-concepts-office-add-ins.md)
- [<span data-ttu-id="139a8-139">Office アドインを開発する</span><span class="sxs-lookup"><span data-stu-id="139a8-139">Develop Office Add-ins</span></span>](../develop/develop-overview.md)
- [<span data-ttu-id="139a8-140">Office アドインを設計する</span><span class="sxs-lookup"><span data-stu-id="139a8-140">Design Office Add-ins</span></span>](../design/add-in-design.md)
- [<span data-ttu-id="139a8-141">Office アドインのテストとデバッグ</span><span class="sxs-lookup"><span data-stu-id="139a8-141">Test and debug Office Add-ins</span></span>](../testing/test-debug-office-add-ins.md)
- [<span data-ttu-id="139a8-142">Office アドインを発行する</span><span class="sxs-lookup"><span data-stu-id="139a8-142">Publish Office Add-ins</span></span>](../publish/publish.md)