---
title: Office アドインを構築する
description: Office アドイン構築の概要。
ms.date: 02/27/2020
localization_priority: Priority
ms.openlocfilehash: 5520a147ed1dfe234d78b4e83081e355bc3e1872
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/28/2020
ms.locfileid: "47292458"
---
# <a name="building-office-add-ins"></a><span data-ttu-id="1c720-103">Office アドインを構築する</span><span class="sxs-lookup"><span data-stu-id="1c720-103">Building Office Add-ins</span></span>

> [!TIP]
> <span data-ttu-id="1c720-104">この記事を読む前に、「[Office Add-ins platform overview (Office アドイン プラットフォームの概要)](office-add-ins.md)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="1c720-104">Please review [Office Add-ins platform overview](office-add-ins.md) before reading this article.</span></span>

<span data-ttu-id="1c720-105">Office アドインは、Office アプリケーションの UI と機能を拡張し、Office ドキュメント内のコンテンツを操作します。</span><span class="sxs-lookup"><span data-stu-id="1c720-105">Office Add-ins extend the UI and functionality of Office applications and interact with content in Office documents.</span></span> <span data-ttu-id="1c720-106">Word、Excel、PowerPoint、OneNote、Project、Outlook の拡張と操作を行うアドインの構築には、一般的な Web テクノロジを使用します。</span><span class="sxs-lookup"><span data-stu-id="1c720-106">You'll use familiar web technologies to create Office Add-ins that extend and interact with Word, Excel, PowerPoint, OneNote, Project, or Outlook.</span></span> <span data-ttu-id="1c720-107">構築するアドインは、Windows、Mac、iPad やブラウザー上など、複数のプラットフォーム上の Office で実行できます。</span><span class="sxs-lookup"><span data-stu-id="1c720-107">The add-ins you build can run in Office across multiple platforms, including Windows, Mac, iPad, and in a browser.</span></span> <span data-ttu-id="1c720-108">この記事では、Office アドイン開発の概要を説明します。</span><span class="sxs-lookup"><span data-stu-id="1c720-108">This article provides an introduction to developing Office Add-ins.</span></span>

## <a name="creating-an-office-add-in"></a><span data-ttu-id="1c720-109">Office アドインの作成</span><span class="sxs-lookup"><span data-stu-id="1c720-109">Creating an Office Add-in</span></span> 

<span data-ttu-id="1c720-110">Office アドイン用の Yeoman ジェネレーターまたは Visual Studio を使用して Office アドインを作成することができます。</span><span class="sxs-lookup"><span data-stu-id="1c720-110">You can create an Office Add-in by using the Yeoman generator for Office Add-ins or Visual Studio.</span></span>

### <a name="yeoman-generator-for-office-add-ins"></a><span data-ttu-id="1c720-111">Office アドイン用の Yeoman ジェネレーター</span><span class="sxs-lookup"><span data-stu-id="1c720-111">Yeoman generator for Office Add-ins</span></span>

<span data-ttu-id="1c720-112">[Office アドイン用の Yeoman ジェネレーター](https://github.com/officedev/generator-office)を使用することで、Visual Studio Code やその他のエディターで管理することができる、Node.js Office アドイン プロジェクトを作成できます。</span><span class="sxs-lookup"><span data-stu-id="1c720-112">The [Yeoman generator for Office Add-ins](https://github.com/officedev/generator-office) can be used to create a Node.js Office Add-in project that can be managed with Visual Studio Code or any other editor.</span></span> <span data-ttu-id="1c720-113">ジェネレーターでは、次のいずれのホスト用の Office アドインも作成できます。</span><span class="sxs-lookup"><span data-stu-id="1c720-113">The generator can create Office Add-ins for any of the following:</span></span>

- <span data-ttu-id="1c720-114">Excel</span><span class="sxs-lookup"><span data-stu-id="1c720-114">Excel</span></span>
- <span data-ttu-id="1c720-115">OneNote</span><span class="sxs-lookup"><span data-stu-id="1c720-115">OneNote</span></span>
- <span data-ttu-id="1c720-116">Outlook</span><span class="sxs-lookup"><span data-stu-id="1c720-116">Outlook</span></span>
- <span data-ttu-id="1c720-117">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="1c720-117">PowerPoint</span></span>
- <span data-ttu-id="1c720-118">Project</span><span class="sxs-lookup"><span data-stu-id="1c720-118">Project</span></span>
- <span data-ttu-id="1c720-119">Word</span><span class="sxs-lookup"><span data-stu-id="1c720-119">Word</span></span>
- <span data-ttu-id="1c720-120">Excel のカスタム関数</span><span class="sxs-lookup"><span data-stu-id="1c720-120">Excel custom functions</span></span>

<span data-ttu-id="1c720-121">プロジェクトを作成するのに、HTML、CSS、および JavaScript を使用するのか、Angular または React を使用するのかを選択できます。</span><span class="sxs-lookup"><span data-stu-id="1c720-121">You can choose to create the project using HTML, CSS and JavaScript, or using Angular or React.</span></span> <span data-ttu-id="1c720-122">いずれのフレームワークを選択した場合も、JavaScript と Typescript の間から選択することができます。</span><span class="sxs-lookup"><span data-stu-id="1c720-122">For whichever framework you choose, you can choose between JavaScript and Typescript as well.</span></span> <span data-ttu-id="1c720-123">Yeoman ジェネレーターを使用してアドインを作成する方法については、「[Visual Studio Code を使用して Office アドインを開発する](../develop/develop-add-ins-vscode.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="1c720-123">For more information about creating add-ins with the Yeoman generator, see [Develop Office Add-ins with Visual Studio Code](../develop/develop-add-ins-vscode.md).</span></span>

### <a name="visual-studio"></a><span data-ttu-id="1c720-124">Visual Studio</span><span class="sxs-lookup"><span data-stu-id="1c720-124">Visual Studio</span></span>

<span data-ttu-id="1c720-125">Visual Studio は、Excel、Outlook、Word、および PowerPoint 用の Office アドインの作成に使用できます。</span><span class="sxs-lookup"><span data-stu-id="1c720-125">Visual Studio can be used to create Office Add-ins for Excel, Outlook, Word, and PowerPoint.</span></span> <span data-ttu-id="1c720-126">Office アドイン プロジェクトは Visual Studio ソリューションの一部として作成され、HTML、CSS、および JavaScript が使用されます。</span><span class="sxs-lookup"><span data-stu-id="1c720-126">An Office Add-in project gets created as part of a Visual Studio solution and uses HTML, CSS, and JavaScript.</span></span> <span data-ttu-id="1c720-127">Visual Studio を使用してアドインを作成する方法については、「[Visual Studio を使用して Office アドインを開発する](../develop/develop-add-ins-visual-studio.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="1c720-127">For more information about creating add-ins with Visual Studio, see [Develop Office Add-ins with Visual Studio](../develop/develop-add-ins-visual-studio.md).</span></span>

[!include[Yeoman vs Visual Studio comparision](../includes/yeoman-generator-recommendation.md)]

## <a name="exploring-apis-with-script-lab"></a><span data-ttu-id="1c720-128">Script Lab を使用して API を調べる</span><span class="sxs-lookup"><span data-stu-id="1c720-128">Exploring APIs with Script Lab</span></span>

<span data-ttu-id="1c720-129">Script Lab は、Excel や Word などの Office プログラムでの作業中に Office JavaScript API を調査し、コード スニペットを実行できるようにするアドインです。</span><span class="sxs-lookup"><span data-stu-id="1c720-129">Script Lab is an add-in that enables you to explore the Office JavaScript API and run code snippets while you're working in an Office program such as Excel or Word.</span></span> <span data-ttu-id="1c720-130">これは、[AppSource](https://appsource.microsoft.com/product/office/WA104380862) から無料で利用でき、アドインで必要な機能のプロトタイプを作成したり検証したりする場合に、開発ツールキットに含めておくと便利なツールです。</span><span class="sxs-lookup"><span data-stu-id="1c720-130">It's available for free via [AppSource](https://appsource.microsoft.com/product/office/WA104380862) and is a useful tool to include in your development toolkit as you prototype and verify the functionality you want in your add-in.</span></span> <span data-ttu-id="1c720-131">Script Lab では、組み込みのサンプルのライブラリにアクセスして、簡単に API を試すことができます。また、独自のコードの開始点としてサンプルを使用することもできます。</span><span class="sxs-lookup"><span data-stu-id="1c720-131">In Script Lab, you can access a library of built-in samples to quickly try out APIs or even use a sample as the starting point for your own code.</span></span> 

<span data-ttu-id="1c720-132">次の 1 分間のビデオで、Script Lab の実際の動作をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="1c720-132">The following one-minute video shows Script Lab in action.</span></span>

<span data-ttu-id="1c720-133">[![Excel、Word、PowerPoint での Script Lab の実行を紹介するプレビュー ビデオ。](../images/screenshot-wide-youtube.png 'Script Lab のプレビュー ビデオ')](https://aka.ms/scriptlabvideo)</span><span class="sxs-lookup"><span data-stu-id="1c720-133">[![Preview video showing Script Lab running in Excel, Word, and PowerPoint.](../images/screenshot-wide-youtube.png 'Script Lab preview video')](https://aka.ms/scriptlabvideo)</span></span>

<span data-ttu-id="1c720-134">Script Lab の詳細については、「[Script Lab を使用して Office JavaScript API を調べる](../overview/explore-with-script-lab.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="1c720-134">For more information about Script Lab, see [Explore Office JavaScript APIs using Script Lab](../overview/explore-with-script-lab.md).</span></span>

## <a name="extending-the-office-ui"></a><span data-ttu-id="1c720-135">Office UI の拡張</span><span class="sxs-lookup"><span data-stu-id="1c720-135">Extending the Office UI</span></span>

<span data-ttu-id="1c720-136">Office アドインは、作業ウィンドウ、コンテンツ アドイン、ダイアログ ボックスなど、アドイン コマンドや HTML コンテナーを使用 Office UI を拡張することができます。</span><span class="sxs-lookup"><span data-stu-id="1c720-136">An Office Add-in can extend the Office UI by using add-in commands and HTML containers such as task panes, content add-ins, or dialog boxes.</span></span>

- <span data-ttu-id="1c720-137">[アドイン コマンド](../design/add-in-commands.md) を使用すると、Office の既定のリボンにカスタム タブ、ボタン、メニューを追加したり、ユーザーが Office ドキュメント内のテキストまたは Excel 内のオブジェクトを右クリックした際に表示される既定のコンテキスト メニューを拡張したりすることができます。</span><span class="sxs-lookup"><span data-stu-id="1c720-137">[Add-in commands](../design/add-in-commands.md) can be used to add custom tabs, buttons, and menus to the default ribbon in Office, or to extend the default context menu that appears when users right-click text in an Office document or an object in Excel.</span></span> <span data-ttu-id="1c720-138">ユーザーがアドイン コマンドを選択すると、アドイン コマンドで指定されているタスク (JavaScript コードの実行、作業ウィンドウを開く、ダイアログ ボックスの起動など) が実行されます。</span><span class="sxs-lookup"><span data-stu-id="1c720-138">When users select an add-in command, they initiate the task that the add-in command specifies, such as running JavaScript code, opening a task pane, or launching a dialog box.</span></span>

- <span data-ttu-id="1c720-139">[作業ウィンドウ](../design/task-pane-add-ins.md)、[コンテンツ アドイン](../design/content-add-ins.md)、[ダイアログ ボックス](../design/dialog-boxes.md)などの HTML コンテナーを使用すると、カスタム UI を表示させたり Office アプリケーション内で追加機能を表示させたりすることができます。</span><span class="sxs-lookup"><span data-stu-id="1c720-139">HTML containers like [task panes](../design/task-pane-add-ins.md), [content add-ins](../design/content-add-ins.md), and [dialog boxes](../design/dialog-boxes.md) can be used to display custom UI and expose additional functionality within an Office application.</span></span> <span data-ttu-id="1c720-140">各作業ウィンドウ、コンテンツ アドイン、またはダイアログ ボックスのコンテンツと機能は、指定した Web ページに由来します。</span><span class="sxs-lookup"><span data-stu-id="1c720-140">The content and functionality of each task pane, content add-in, or dialog box derives from a web page that you specify.</span></span> <span data-ttu-id="1c720-141">これらの Web ページでは、Office JavaScript API を使用することで、アドインが実行されている Office ドキュメントのコンテンツを操作できます。また、外部 Web サービスの呼び出しやユーザー認証の要求など、Web ページが一般的に行うその他の機能も実行できます。</span><span class="sxs-lookup"><span data-stu-id="1c720-141">Those web pages can use the Office JavaScript API to interact with content in the Office document where the add-in is running, and can also do other things that web pages typically do, like call external web services, facilitate user authentication, and more.</span></span>

<span data-ttu-id="1c720-142">次の図では、リボン上に表示されるアドイン コマンド、ドキュメント右側に表示される作業ウィンドウ、およびドキュメント上に表示されるダイアログ ボックスまたはコンテンツ アドインを示しています。</span><span class="sxs-lookup"><span data-stu-id="1c720-142">The following image shows an add-in command in the ribbon, a task pane to the right of the document, and a dialog box or content add-in over the document.</span></span>

![Office ドキュメントのリボン、タスク ウィンドウ、ダイアログ ボックス上のアドイン コマンドを示す図](../images/add-in-ui-elements.png)

<span data-ttu-id="1c720-144">Office UI の拡張に関する詳細については、「[Office アドイン用の Office UI 要素](../design/interface-elements.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="1c720-144">For more information about extending the Office UI, see [Office UI elements for Office Add-ins](../design/interface-elements.md).</span></span>

## <a name="core-development-concepts"></a><span data-ttu-id="1c720-145">開発の中心概念</span><span class="sxs-lookup"><span data-stu-id="1c720-145">Core development concepts</span></span> 

<span data-ttu-id="1c720-146">Office アドインは、2 つの部分から構成されます。</span><span class="sxs-lookup"><span data-stu-id="1c720-146">An Office Add-in consists of two parts:</span></span>

- <span data-ttu-id="1c720-147">アドインの設定と機能を定義るアドイン マニフェスト (XML ファイル)。</span><span class="sxs-lookup"><span data-stu-id="1c720-147">The add-in manifest (an XML file) that defines the settings and capabilities of the add-in.</span></span>

- <span data-ttu-id="1c720-148">作業ウィンドウ、コンテンツ アドイン、ダイアログ ボックスなど、アドインの UI と機能を定義する Web アプリケーション。</span><span class="sxs-lookup"><span data-stu-id="1c720-148">The web application that defines the UI and functionality of add-in components such as task panes, content add-ins, and dialog boxes.</span></span>

<span data-ttu-id="1c720-149">Web アプリケーションでは、Office JavaScript API を使用することで、アドインが実行されている Office ドキュメント内のコンテンツを操作します。</span><span class="sxs-lookup"><span data-stu-id="1c720-149">The web application uses the Office JavaScript API to interact with content in the Office document where the add-in is running.</span></span> <span data-ttu-id="1c720-150">アドインは、外部 Web サービスの呼び出しやユーザー認証の要求など、Web ページが一般的に行うその他の機能も実行することができます。</span><span class="sxs-lookup"><span data-stu-id="1c720-150">Your add-in can also do other things that web applications typically do, like call external web services, facilitate user authentication, and more.</span></span>

### <a name="defining-an-add-ins-settings-and-capabilities"></a><span data-ttu-id="1c720-151">アドインの設定と機能を定義する</span><span class="sxs-lookup"><span data-stu-id="1c720-151">Defining an add-in's settings and capabilities</span></span>

<span data-ttu-id="1c720-152">Office アドインのマニフェスト (XML ファイル) は、アドインの設定と機能を定義します。</span><span class="sxs-lookup"><span data-stu-id="1c720-152">An Office Add-in's manifest (an XML file) defines the settings and capabilities of the add-in.</span></span> <span data-ttu-id="1c720-153">次のような要素を定義するには、マニフェストを構成します。</span><span class="sxs-lookup"><span data-stu-id="1c720-153">You'll configure the manifest to specify things such as:</span></span>

- <span data-ttu-id="1c720-154">アドインを説明するメタデータ (ID、バージョン、説明、表示名、既定のロケールなど)。</span><span class="sxs-lookup"><span data-stu-id="1c720-154">Metadata that describes the add-in (for example, ID, version, description, display name, default locale).</span></span>
- <span data-ttu-id="1c720-155">アドインが実行される Office アプリケーション。</span><span class="sxs-lookup"><span data-stu-id="1c720-155">Office applications where the add-in will run.</span></span>
- <span data-ttu-id="1c720-156">アドインに必要なアクセス許可。</span><span class="sxs-lookup"><span data-stu-id="1c720-156">Permissions that the add-in requires.</span></span>
- <span data-ttu-id="1c720-157">アドインによって作成されるカスタム UI (カスタム タブ、リボンのボタンなど) などの統合も含めた、アドインの Office との統合方法。</span><span class="sxs-lookup"><span data-stu-id="1c720-157">How the add-in integrates with Office, including any custom UI that the add-in creates (for example, custom tabs, ribbon buttons).</span></span>
- <span data-ttu-id="1c720-158">ブランドおよびコマンドの図像としてアドインで使用される画像の場所。</span><span class="sxs-lookup"><span data-stu-id="1c720-158">Location of images that the add-in uses for branding and command iconography.</span></span>
- <span data-ttu-id="1c720-159">アドインの寸法 (例: コンテンツ アドインの寸法、Outlook アドインに対して要求される高さなど)。</span><span class="sxs-lookup"><span data-stu-id="1c720-159">Dimensions of the add-in (for example, dimensions for content add-ins, requested height for Outlook add-ins).</span></span>
- <span data-ttu-id="1c720-160">メッセージや予定のコンテキストでアドインをアクティブにさせるタイミングを指定するルール (Outlook アドインのみ)。</span><span class="sxs-lookup"><span data-stu-id="1c720-160">Rules that specify when the add-in activates in the context of a message or appointment (for Outlook add-ins only).</span></span>

<span data-ttu-id="1c720-161">マニフェストの詳細については、「[Office アドインの XML マニフェスト](add-in-manifests.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="1c720-161">For detailed information about the manifest, see [Office Add-ins XML manifest](add-in-manifests.md).</span></span>

### <a name="interacting-with-content-in-an-office-document"></a><span data-ttu-id="1c720-162">Office ドキュメント内のコンテンツを操作する</span><span class="sxs-lookup"><span data-stu-id="1c720-162">Interacting with content in an Office document</span></span>

<span data-ttu-id="1c720-163">Office アドインでは、Office JavaScript API を使用することで、アドインが実行されている Office ドキュメント内のコンテンツを操作できます。</span><span class="sxs-lookup"><span data-stu-id="1c720-163">An Office Add-in can use the Office JavaScript APIs to interact with content in the Office document where the add-in is running.</span></span> 

#### <a name="accessing-the-office-javascript-api-library"></a><span data-ttu-id="1c720-164">Office JavaScript API ライブラリへのアクセス</span><span class="sxs-lookup"><span data-stu-id="1c720-164">Accessing the Office JavaScript API library</span></span>

[!include[information about accessing the Office JS API library](../includes/office-js-access-library.md)]

#### <a name="api-models"></a><span data-ttu-id="1c720-165">API モデル</span><span class="sxs-lookup"><span data-stu-id="1c720-165">API models</span></span>

[!include[information about the Office JS API models](../includes/office-js-api-models.md)]

#### <a name="api-requirement-sets"></a><span data-ttu-id="1c720-166">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="1c720-166">API requirement sets</span></span>

[!include[information about the Office JS API requirement sets](../includes/office-js-requirement-sets.md)]

## <a name="testing-and-debugging-an-office-add-in"></a><span data-ttu-id="1c720-167">Office アドインのテストとデバッグ</span><span class="sxs-lookup"><span data-stu-id="1c720-167">Testing and debugging an Office Add-in</span></span>

<span data-ttu-id="1c720-168">アドインの開発中は、_サイドロード_という手法を使用してアドインをローカルでテストできます。</span><span class="sxs-lookup"><span data-stu-id="1c720-168">As you develop your add-in, you can test it locally by using a technique known as _sideloading_.</span></span> <span data-ttu-id="1c720-169">アドインをサイドロードする手順はプラットフォームによって異なり、一部のケースでは、製品ごとに異なります。</span><span class="sxs-lookup"><span data-stu-id="1c720-169">The procedure for sideloading an add-in varies by platform, and in some cases, by product as well.</span></span> <span data-ttu-id="1c720-170">同様に、アドインのデバッグ手順も、プラットフォームや製品によって異なります。</span><span class="sxs-lookup"><span data-stu-id="1c720-170">Likewise, the procedure for debugging an add-in can also vary by platform and product.</span></span> <span data-ttu-id="1c720-171">テストとデバッグの詳細については、「[Office アドインのテストとデバッグ](../testing/test-debug-office-add-ins.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="1c720-171">For more information about testing and debugging, see [Test and debug Office Add-ins](../testing/test-debug-office-add-ins.md).</span></span>

## <a name="publishing-an-office-add-in"></a><span data-ttu-id="1c720-172">Office アドインの公開</span><span class="sxs-lookup"><span data-stu-id="1c720-172">Publishing an Office Add-in</span></span>

<span data-ttu-id="1c720-173">アドインを他のユーザーと共有する準備ができたら、目的に一番合った展開方法を使用してアドインを共有します。</span><span class="sxs-lookup"><span data-stu-id="1c720-173">When you're ready to share your add-in with others, you'll do so by using the deployment method that best meets your objectives.</span></span> <span data-ttu-id="1c720-174">たとえば、組織内のユーザーにアドインを展開する場合は、一元展開を使用するか、アドインを SharePoint アプリ カタログで公開することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="1c720-174">For example, to deploy an add-in to users within your organization, you might use centralized deployment or publish the add-in to a SharePoint app catalog.</span></span> <span data-ttu-id="1c720-175">すべてのユーザーが入手できるようにアドインを一般公開する場合は、アドインを AppSource で公開できます。</span><span class="sxs-lookup"><span data-stu-id="1c720-175">If you want to share your add-in publicly for anyone to obtain, you can publish the add-in to AppSource.</span></span> <span data-ttu-id="1c720-176">公開の詳細については、「[Office アドインの展開と公開](../publish/publish.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="1c720-176">For more information about publishing, see [Deploy and publish Office Add-ins](../publish/publish.md).</span></span>

## <a name="next-steps"></a><span data-ttu-id="1c720-177">次のステップ</span><span class="sxs-lookup"><span data-stu-id="1c720-177">Next steps</span></span>

<span data-ttu-id="1c720-178">この記事では、Office アドインの異なる作成方法を説明し、Office JavaScript API の調査とアドイン機能のプロトタイプ作成における効果的なツールとして Script Lab を紹介し、Office アドインの開発、テスト、および公開に関する重要な概念の説明を行いました。</span><span class="sxs-lookup"><span data-stu-id="1c720-178">This article has outlined the different ways to create Office Add-ins, introduced Script Lab as a valuable tool for exploring Office JavaScript APIs and prototyping add-in functionality, and described important Office Add-ins development, testing, and publishing concepts.</span></span> <span data-ttu-id="1c720-179">初歩的な情報の説明は以上になります。Office アドインにの行程を先に進むには、 次の手順を実行してください。</span><span class="sxs-lookup"><span data-stu-id="1c720-179">Now that you've explored this introductory information, consider continuing your Office Add-ins journey along the following paths.</span></span>

### <a name="create-an-office-add-in"></a><span data-ttu-id="1c720-180">Office アドインを作成する</span><span class="sxs-lookup"><span data-stu-id="1c720-180">Create an Office add-in</span></span>

<span data-ttu-id="1c720-181">[5 分間のクイック スタート](/office/dev/add-ins/)を完了することで、Excel、OneNote、Outlook、PowerPoint、Project、または Word 用の基本的なアドインを簡単に作成することができます。</span><span class="sxs-lookup"><span data-stu-id="1c720-181">You can quickly create a basic add-in for Excel, OneNote, Outlook, PowerPoint, Project, or Word by completing a [5-minute quick start](/office/dev/add-ins/).</span></span> <span data-ttu-id="1c720-182">以前にクイック スタートを完了している場合で、より複雑なアドインを作成したい場合は、[チュートリアル](/office/dev/add-ins/)を試してみてください。</span><span class="sxs-lookup"><span data-stu-id="1c720-182">If you've previously completed a quick start and want to create a slightly more complex add-in, you should try the [tutorial](/office/dev/add-ins/).</span></span>

### <a name="explore-the-apis-with-script-lab"></a><span data-ttu-id="1c720-183">Script Lab を使用して API を調べる</span><span class="sxs-lookup"><span data-stu-id="1c720-183">Explore the APIs with Script Lab</span></span>

<span data-ttu-id="1c720-184">Office JavaScript API でどのような機能が提供されているかを把握するには、[Script Lab](explore-with-script-lab.md) に組み込まれているサンプルのライブラリを参照してください。</span><span class="sxs-lookup"><span data-stu-id="1c720-184">Explore the library of built-in samples in [Script Lab](explore-with-script-lab.md) to get a sense for the capabilities of the Office JavaScript APIs.</span></span>

### <a name="learn-more"></a><span data-ttu-id="1c720-185">詳細情報</span><span class="sxs-lookup"><span data-stu-id="1c720-185">Learn more</span></span>

<span data-ttu-id="1c720-186">Office アドインの開発、テスト、公開の詳細については、このドキュメントを参照してください。</span><span class="sxs-lookup"><span data-stu-id="1c720-186">Learn more about developing, testing, and publishing Office Add-ins by exploring this documentation.</span></span>

> [!TIP]
> <span data-ttu-id="1c720-187">どのようなアドインを構築する場合でも、このドキュメントの 「[中心概念](core-concepts-office-add-ins.md)」セクションに記載する情報に加え、構築するアドインの種類に対応するアプリケーション固有のセクション (たとえば、[Excel](../excel/index.yml)) に記載する情報を使用してください。</span><span class="sxs-lookup"><span data-stu-id="1c720-187">For any add-in that you build, you'll use information in the [Core concepts](core-concepts-office-add-ins.md) section of this documentation, along with information in the application-specific section that corresponds to the type of add-in you're building (for example, [Excel](../excel/index.yml)).</span></span>
>
> ![目次を表示する画像](../images/top-level-toc.png)

## <a name="see-also"></a><span data-ttu-id="1c720-189">関連項目</span><span class="sxs-lookup"><span data-stu-id="1c720-189">See also</span></span> 

- [<span data-ttu-id="1c720-190">Office アドイン プラットフォームの概要</span><span class="sxs-lookup"><span data-stu-id="1c720-190">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="1c720-191">Office アドインの中心概念</span><span class="sxs-lookup"><span data-stu-id="1c720-191">Core concepts for Office Add-ins</span></span>](../overview/core-concepts-office-add-ins.md)
- [<span data-ttu-id="1c720-192">Office アドインを開発する</span><span class="sxs-lookup"><span data-stu-id="1c720-192">Develop Office Add-ins</span></span>](../develop/develop-overview.md)
- [<span data-ttu-id="1c720-193">Visual Studio Code を使用して Office アドインを開発する</span><span class="sxs-lookup"><span data-stu-id="1c720-193">Develop Office Add-ins with Visual Studio Code</span></span>](../develop/develop-add-ins-vscode.md)
- [<span data-ttu-id="1c720-194">Visual Studio を使用して Office アドインを開発する</span><span class="sxs-lookup"><span data-stu-id="1c720-194">Develop Office Add-ins with Visual Studio</span></span>](../develop/develop-add-ins-visual-studio.md)
- [<span data-ttu-id="1c720-195">Office アドインを設計する</span><span class="sxs-lookup"><span data-stu-id="1c720-195">Design Office Add-ins</span></span>](../design/add-in-design.md)
- [<span data-ttu-id="1c720-196">Office アドインのテストとデバッグ</span><span class="sxs-lookup"><span data-stu-id="1c720-196">Test and debug Office Add-ins</span></span>](../testing/test-debug-office-add-ins.md)
- [<span data-ttu-id="1c720-197">Office アドインの公開</span><span class="sxs-lookup"><span data-stu-id="1c720-197">Publish Office Add-ins</span></span>](../publish/publish.md)
