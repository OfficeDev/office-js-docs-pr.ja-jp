---
title: Office アドインを開発する
description: Office アドイン開発の概要を説明します。
ms.date: 10/14/2020
localization_priority: Priority
ms.openlocfilehash: c0a4ebbae234c9ac1da3b0765a04e31f956ef617
ms.sourcegitcommit: ceb8dd66f3fb9c963fce8446c2f6c65ead56fbc1
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/18/2020
ms.locfileid: "49131893"
---
# <a name="develop-office-add-ins"></a><span data-ttu-id="960d3-103">Office アドインを開発する</span><span class="sxs-lookup"><span data-stu-id="960d3-103">Develop Office Add-ins</span></span>

> [!TIP]
> <span data-ttu-id="960d3-104">この記事を読む前に、「[Office Add-ins platform overview (Office アドイン プラットフォームの概要)](../overview/office-add-ins.md)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="960d3-104">Please review [Office Add-ins platform overview](../overview/office-add-ins.md) before reading this article.</span></span>

<span data-ttu-id="960d3-105">すべての Office アドインは、Office アドイン プラットフォーム上で構築します。</span><span class="sxs-lookup"><span data-stu-id="960d3-105">All Office Add-ins are built upon the Office Add-ins platform.</span></span> <span data-ttu-id="960d3-106">どのようなアドインを構築する場合でも、アプリケーションやプラットフォームの可用性、Office JavaScript API のプログラミング パターン、アドインの設定と機能をマニフェスト ファイル上で指定する方法、マニフェストファイルのcapabilities、UIとユーザーエクスペリエンスをデザインする方法など、重要な概念を理解する必要があります。</span><span class="sxs-lookup"><span data-stu-id="960d3-106">For any add-in you build, you'll need to understand important concepts like application and platform availability, Office JavaScript API programming patterns, how to specify an add-in's settings and capabilities in the manifest file, how to design the UI and user experience, and more.</span></span> <span data-ttu-id="960d3-107">開発に関するこれらの中心概念については、ドキュメントの **開発ライフサイクル** > **開発** セクションを参照してください。</span><span class="sxs-lookup"><span data-stu-id="960d3-107">Core development concepts like these are covered here in the **Development lifecycle** > **Develop** section of the documentation.</span></span> <span data-ttu-id="960d3-108">構築するアドインに対応するアプリケーション固有のドキュメント (たとえば、 [Excel](../excel/index.yml)) を詳しく見る前に、ここに記載される情報を確認してください。</span><span class="sxs-lookup"><span data-stu-id="960d3-108">Review the information here before exploring the application-specific documentation that corresponds to the add-in you're building (for example, [Excel](../excel/index.yml)).</span></span>

## <a name="creating-an-office-add-in"></a><span data-ttu-id="960d3-109">Office アドインの作成</span><span class="sxs-lookup"><span data-stu-id="960d3-109">Creating an Office Add-in</span></span>

<span data-ttu-id="960d3-110">Office アドイン用の Yeoman ジェネレーターまたは Visual Studio を使用して Office アドインを作成することができます。</span><span class="sxs-lookup"><span data-stu-id="960d3-110">You can create an Office Add-in by using the Yeoman generator for Office Add-ins or Visual Studio.</span></span>

### <a name="yeoman-generator-for-office-add-ins"></a><span data-ttu-id="960d3-111">Office アドイン用の Yeoman ジェネレーター</span><span class="sxs-lookup"><span data-stu-id="960d3-111">Yeoman generator for Office Add-ins</span></span>

<span data-ttu-id="960d3-112">[Office アドイン用の Yeoman ジェネレーター](https://github.com/officedev/generator-office)を使用することで、Visual Studio Code やその他のエディターで管理することができる、Node.js Office アドイン プロジェクトを作成できます。</span><span class="sxs-lookup"><span data-stu-id="960d3-112">The [Yeoman generator for Office Add-ins](https://github.com/officedev/generator-office) can be used to create a Node.js Office Add-in project that can be managed with Visual Studio Code or any other editor.</span></span> <span data-ttu-id="960d3-113">ジェネレーターでは、次のいずれのホスト用の Office アドインも作成できます。</span><span class="sxs-lookup"><span data-stu-id="960d3-113">The generator can create Office Add-ins for any of the following:</span></span>

- <span data-ttu-id="960d3-114">Excel</span><span class="sxs-lookup"><span data-stu-id="960d3-114">Excel</span></span>
- <span data-ttu-id="960d3-115">OneNote</span><span class="sxs-lookup"><span data-stu-id="960d3-115">OneNote</span></span>
- <span data-ttu-id="960d3-116">Outlook</span><span class="sxs-lookup"><span data-stu-id="960d3-116">Outlook</span></span>
- <span data-ttu-id="960d3-117">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="960d3-117">PowerPoint</span></span>
- <span data-ttu-id="960d3-118">Project</span><span class="sxs-lookup"><span data-stu-id="960d3-118">Project</span></span>
- <span data-ttu-id="960d3-119">Word</span><span class="sxs-lookup"><span data-stu-id="960d3-119">Word</span></span>
- <span data-ttu-id="960d3-120">Excel のカスタム関数</span><span class="sxs-lookup"><span data-stu-id="960d3-120">Excel custom functions</span></span>

<span data-ttu-id="960d3-121">プロジェクトを作成するのに、HTML、CSS、および JavaScript を使用するのか、Angular または React を使用するのかを選択できます。</span><span class="sxs-lookup"><span data-stu-id="960d3-121">You can choose to create the project using HTML, CSS and JavaScript, or using Angular or React.</span></span> <span data-ttu-id="960d3-122">いずれのフレームワークを選択した場合も、JavaScript と Typescript の間から選択することができます。</span><span class="sxs-lookup"><span data-stu-id="960d3-122">For whichever framework you choose, you can choose between JavaScript and Typescript as well.</span></span> <span data-ttu-id="960d3-123">Yeoman ジェネレーターを使用してアドインを作成する方法については、「[Visual Studio Code を使用して Office アドインを開発する](../develop/develop-add-ins-vscode.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="960d3-123">For more information about creating add-ins with the Yeoman generator, see [Develop Office Add-ins with Visual Studio Code](../develop/develop-add-ins-vscode.md).</span></span>

### <a name="visual-studio"></a><span data-ttu-id="960d3-124">Visual Studio</span><span class="sxs-lookup"><span data-stu-id="960d3-124">Visual Studio</span></span>

<span data-ttu-id="960d3-125">Visual Studio は、Excel、Outlook、Word、および PowerPoint 用の Office アドインの作成に使用できます。</span><span class="sxs-lookup"><span data-stu-id="960d3-125">Visual Studio can be used to create Office Add-ins for Excel, Outlook, Word, and PowerPoint.</span></span> <span data-ttu-id="960d3-126">Office アドイン プロジェクトは Visual Studio ソリューションの一部として作成され、HTML、CSS、および JavaScript が使用されます。</span><span class="sxs-lookup"><span data-stu-id="960d3-126">An Office Add-in project gets created as part of a Visual Studio solution and uses HTML, CSS, and JavaScript.</span></span> <span data-ttu-id="960d3-127">Visual Studio を使用してアドインを作成する方法については、「[Visual Studio を使用して Office アドインを開発する](../develop/develop-add-ins-visual-studio.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="960d3-127">For more information about creating add-ins with Visual Studio, see [Develop Office Add-ins with Visual Studio](../develop/develop-add-ins-visual-studio.md).</span></span>

[!include[Yeoman vs Visual Studio comparison](../includes/yeoman-generator-recommendation.md)]

## <a name="understanding-the-two-parts-of-an-office-add-in"></a><span data-ttu-id="960d3-128">Office アドインの2つの部分について</span><span class="sxs-lookup"><span data-stu-id="960d3-128">Understanding the two parts of an Office add-in</span></span>

<span data-ttu-id="960d3-129">Office アドインは、2 つの部分から構成されます。</span><span class="sxs-lookup"><span data-stu-id="960d3-129">An Office Add-in consists of two parts:</span></span>

- <span data-ttu-id="960d3-130">アドインの設定と機能を定義るアドイン マニフェスト (XML ファイル)。</span><span class="sxs-lookup"><span data-stu-id="960d3-130">The add-in manifest (an XML file) that defines the settings and capabilities of the add-in.</span></span>

- <span data-ttu-id="960d3-131">作業ウィンドウ、コンテンツ アドイン、ダイアログ ボックスなど、アドインの UI と機能を定義する Web アプリケーション。</span><span class="sxs-lookup"><span data-stu-id="960d3-131">The web application that defines the UI and functionality of add-in components such as task panes, content add-ins, and dialog boxes.</span></span>

<span data-ttu-id="960d3-132">Web アプリケーションでは、Office JavaScript API を使用することで、アドインが実行されている Office ドキュメント内のコンテンツを操作します。</span><span class="sxs-lookup"><span data-stu-id="960d3-132">The web application uses the Office JavaScript API to interact with content in the Office document where the add-in is running.</span></span> <span data-ttu-id="960d3-133">アドインは、外部 Web サービスの呼び出しやユーザー認証の要求など、Web ページが一般的に行うその他の機能も実行することができます。</span><span class="sxs-lookup"><span data-stu-id="960d3-133">Your add-in can also do other things that web applications typically do, like call external web services, facilitate user authentication, and more.</span></span>

### <a name="defining-an-add-ins-settings-and-capabilities"></a><span data-ttu-id="960d3-134">アドインの設定と機能を定義する</span><span class="sxs-lookup"><span data-stu-id="960d3-134">Defining an add-in's settings and capabilities</span></span>

<span data-ttu-id="960d3-135">Office アドインのマニフェスト (XML ファイル) は、アドインの設定と機能を定義します。</span><span class="sxs-lookup"><span data-stu-id="960d3-135">An Office Add-in's manifest (an XML file) defines the settings and capabilities of the add-in.</span></span> <span data-ttu-id="960d3-136">次のような要素を定義するには、マニフェストを構成します。</span><span class="sxs-lookup"><span data-stu-id="960d3-136">You'll configure the manifest to specify things such as:</span></span>

- <span data-ttu-id="960d3-137">アドインを説明するメタデータ (ID、バージョン、説明、表示名、既定のロケールなど)。</span><span class="sxs-lookup"><span data-stu-id="960d3-137">Metadata that describes the add-in (for example, ID, version, description, display name, default locale).</span></span>
- <span data-ttu-id="960d3-138">アドインが実行される Office アプリケーション。</span><span class="sxs-lookup"><span data-stu-id="960d3-138">Office applications where the add-in will run.</span></span>
- <span data-ttu-id="960d3-139">アドインに必要なアクセス許可。</span><span class="sxs-lookup"><span data-stu-id="960d3-139">Permissions that the add-in requires.</span></span>
- <span data-ttu-id="960d3-140">アドインによって作成されるカスタム UI (カスタム タブ、リボンのボタンなど) などの統合も含めた、アドインの Office との統合方法。</span><span class="sxs-lookup"><span data-stu-id="960d3-140">How the add-in integrates with Office, including any custom UI that the add-in creates (for example, custom tabs, ribbon buttons).</span></span>
- <span data-ttu-id="960d3-141">ブランドおよびコマンドの図像としてアドインで使用される画像の場所。</span><span class="sxs-lookup"><span data-stu-id="960d3-141">Location of images that the add-in uses for branding and command iconography.</span></span>
- <span data-ttu-id="960d3-142">アドインの寸法 (例: コンテンツ アドインの寸法、Outlook アドインに対して要求される高さなど)。</span><span class="sxs-lookup"><span data-stu-id="960d3-142">Dimensions of the add-in (for example, dimensions for content add-ins, requested height for Outlook add-ins).</span></span>
- <span data-ttu-id="960d3-143">メッセージや予定のコンテキストでアドインをアクティブにさせるタイミングを指定するルール (Outlook アドインのみ)。</span><span class="sxs-lookup"><span data-stu-id="960d3-143">Rules that specify when the add-in activates in the context of a message or appointment (for Outlook add-ins only).</span></span>

<span data-ttu-id="960d3-144">マニフェストの詳細については、「[Office アドインの XML マニフェスト](add-in-manifests.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="960d3-144">For detailed information about the manifest, see [Office Add-ins XML manifest](add-in-manifests.md).</span></span>

### <a name="interacting-with-content-in-an-office-document"></a><span data-ttu-id="960d3-145">Office ドキュメント内のコンテンツを操作する</span><span class="sxs-lookup"><span data-stu-id="960d3-145">Interacting with content in an Office document</span></span>

<span data-ttu-id="960d3-146">Office アドインでは、Office JavaScript API を使用することで、アドインが実行されている Office ドキュメント内のコンテンツを操作できます。</span><span class="sxs-lookup"><span data-stu-id="960d3-146">An Office Add-in can use the Office JavaScript APIs to interact with content in the Office document where the add-in is running.</span></span>

#### <a name="accessing-the-office-javascript-api-library"></a><span data-ttu-id="960d3-147">Office JavaScript API ライブラリへのアクセス</span><span class="sxs-lookup"><span data-stu-id="960d3-147">Accessing the Office JavaScript API library</span></span>

[!include[information about accessing the Office JS API library](../includes/office-js-access-library.md)]

#### <a name="api-models"></a><span data-ttu-id="960d3-148">API モデル</span><span class="sxs-lookup"><span data-stu-id="960d3-148">API models</span></span>

[!include[information about the Office JS API models](../includes/office-js-api-models.md)]

#### <a name="api-requirement-sets"></a><span data-ttu-id="960d3-149">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="960d3-149">API requirement sets</span></span>

[!include[information about the Office JS API requirement sets](../includes/office-js-requirement-sets.md)]

#### <a name="exploring-apis-with-script-lab"></a><span data-ttu-id="960d3-150">Script Lab を使用して API を調べる</span><span class="sxs-lookup"><span data-stu-id="960d3-150">Exploring APIs with Script Lab</span></span>

<span data-ttu-id="960d3-151">Script Lab は、Excel や Word などの Office プログラムでの作業中に Office JavaScript API を調査し、コード スニペットを実行できるようにするアドインです。</span><span class="sxs-lookup"><span data-stu-id="960d3-151">Script Lab is an add-in that enables you to explore the Office JavaScript API and run code snippets while you're working in an Office program such as Excel or Word.</span></span> <span data-ttu-id="960d3-152">これは、[AppSource](https://appsource.microsoft.com/product/office/WA104380862) から無料で利用でき、アドインで必要な機能のプロトタイプを作成したり検証したりする場合に、開発ツールキットに含めておくと便利なツールです。</span><span class="sxs-lookup"><span data-stu-id="960d3-152">It's available for free via [AppSource](https://appsource.microsoft.com/product/office/WA104380862) and is a useful tool to include in your development toolkit as you prototype and verify the functionality you want in your add-in.</span></span> <span data-ttu-id="960d3-153">Script Lab では、組み込みのサンプルのライブラリにアクセスして、簡単に API を試すことができます。また、独自のコードの開始点としてサンプルを使用することもできます。</span><span class="sxs-lookup"><span data-stu-id="960d3-153">In Script Lab, you can access a library of built-in samples to quickly try out APIs or even use a sample as the starting point for your own code.</span></span>

<span data-ttu-id="960d3-154">次の 1 分間のビデオで、Script Lab の実際の動作をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="960d3-154">The following one-minute video shows Script Lab in action.</span></span>

<span data-ttu-id="960d3-155">[![Excel、Word、PowerPoint での Script Lab の実行を紹介するショート ビデオ。](../images/screenshot-wide-youtube.png 'Script Lab のプレビュー ビデオ')](https://aka.ms/scriptlabvideo)</span><span class="sxs-lookup"><span data-stu-id="960d3-155">[![Short video that shows Script Lab running in Excel, Word, and PowerPoint](../images/screenshot-wide-youtube.png 'Script Lab preview video')](https://aka.ms/scriptlabvideo)</span></span>

<span data-ttu-id="960d3-156">Script Lab の詳細については、「[Script Lab を使用して Office JavaScript API を調べる](../overview/explore-with-script-lab.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="960d3-156">For more information about Script Lab, see [Explore Office JavaScript APIs using Script Lab](../overview/explore-with-script-lab.md).</span></span>

## <a name="extending-the-office-ui"></a><span data-ttu-id="960d3-157">Office UI の拡張</span><span class="sxs-lookup"><span data-stu-id="960d3-157">Extending the Office UI</span></span>

<span data-ttu-id="960d3-158">Office アドインは、作業ウィンドウ、コンテンツ アドイン、ダイアログ ボックスなど、アドイン コマンドや HTML コンテナーを使用 Office UI を拡張することができます。</span><span class="sxs-lookup"><span data-stu-id="960d3-158">An Office Add-in can extend the Office UI by using add-in commands and HTML containers such as task panes, content add-ins, or dialog boxes.</span></span>

- <span data-ttu-id="960d3-159">[アドイン コマンド](../design/add-in-commands.md) を使用すると、Office の既定のリボンにカスタム タブ、ボタン、メニューを追加したり、ユーザーが Office ドキュメント内のテキストまたは Excel 内のオブジェクトを右クリックした際に表示される既定のコンテキスト メニューを拡張したりすることができます。</span><span class="sxs-lookup"><span data-stu-id="960d3-159">[Add-in commands](../design/add-in-commands.md) can be used to add custom tabs, buttons, and menus to the default ribbon in Office, or to extend the default context menu that appears when users right-click text in an Office document or an object in Excel.</span></span> <span data-ttu-id="960d3-160">ユーザーがアドイン コマンドを選択すると、アドイン コマンドで指定されているタスク (JavaScript コードの実行、作業ウィンドウを開く、ダイアログ ボックスの起動など) が実行されます。</span><span class="sxs-lookup"><span data-stu-id="960d3-160">When users select an add-in command, they initiate the task that the add-in command specifies, such as running JavaScript code, opening a task pane, or launching a dialog box.</span></span>

- <span data-ttu-id="960d3-161">[作業ウィンドウ](../design/task-pane-add-ins.md)、[コンテンツ アドイン](../design/content-add-ins.md)、[ダイアログ ボックス](../design/dialog-boxes.md)などの HTML コンテナーを使用すると、カスタム UI を表示させたり Office アプリケーション内で追加機能を表示させたりすることができます。</span><span class="sxs-lookup"><span data-stu-id="960d3-161">HTML containers like [task panes](../design/task-pane-add-ins.md), [content add-ins](../design/content-add-ins.md), and [dialog boxes](../design/dialog-boxes.md) can be used to display custom UI and expose additional functionality within an Office application.</span></span> <span data-ttu-id="960d3-162">各作業ウィンドウ、コンテンツ アドイン、またはダイアログ ボックスのコンテンツと機能は、指定した Web ページに由来します。</span><span class="sxs-lookup"><span data-stu-id="960d3-162">The content and functionality of each task pane, content add-in, or dialog box derives from a web page that you specify.</span></span> <span data-ttu-id="960d3-163">これらの Web ページでは、Office JavaScript API を使用することで、アドインが実行されている Office ドキュメントのコンテンツを操作できます。また、外部 Web サービスの呼び出しやユーザー認証の要求など、Web ページが一般的に行うその他の機能も実行できます。</span><span class="sxs-lookup"><span data-stu-id="960d3-163">Those web pages can use the Office JavaScript API to interact with content in the Office document where the add-in is running, and can also do other things that web pages typically do, like call external web services, facilitate user authentication, and more.</span></span>

<span data-ttu-id="960d3-164">次の図では、リボン上に表示されるアドイン コマンド、ドキュメント右側に表示される作業ウィンドウ、およびドキュメント上に表示されるダイアログ ボックスまたはコンテンツ アドインを示しています。</span><span class="sxs-lookup"><span data-stu-id="960d3-164">The following image shows an add-in command in the ribbon, a task pane to the right of the document, and a dialog box or content add-in over the document.</span></span>

![Office ドキュメントのリボン、タスク ウィンドウ、ダイアログ ボックス / コンテンツ アドイン上のアドイン コマンドを示す図](../images/add-in-ui-elements.png)

<span data-ttu-id="960d3-166">Office UI の拡張とアドインのUXのデザインに関する詳細については、「[Office アドイン用の Office UI 要素](../design/interface-elements.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="960d3-166">For more information about extending the Office UI and designing the add-in's UX, see [Office UI elements for Office Add-ins](../design/interface-elements.md).</span></span>

## <a name="next-steps"></a><span data-ttu-id="960d3-167">次の手順</span><span class="sxs-lookup"><span data-stu-id="960d3-167">Next steps</span></span>

<span data-ttu-id="960d3-168">この記事では、Office アドインの異なる作成方法を説明し、 アドインがOffice UIを拡張する方法やAPIセットの説明、Office JavaScript APIの探索をするための有益なツールとしてScript Labを紹介し、アドイン機能のプロトタイプ作成の説明を行いました。</span><span class="sxs-lookup"><span data-stu-id="960d3-168">This article has outlined the different ways to create Office Add-ins, introduced the ways that an add-in can extend the Office UI, described the API sets, and introduced Script Lab as a valuable tool for exploring Office JavaScript APIs and prototyping add-in functionality.</span></span> <span data-ttu-id="960d3-169">初歩的な情報の説明は以上になります。Office アドインの行程を先に進むには、 次の手順を実行してください。</span><span class="sxs-lookup"><span data-stu-id="960d3-169">Now that you've explored this introductory information, consider continuing your Office Add-ins journey along the following paths.</span></span>

### <a name="create-an-office-add-in"></a><span data-ttu-id="960d3-170">Office アドインを作成する</span><span class="sxs-lookup"><span data-stu-id="960d3-170">Create an Office Add-in</span></span>

<span data-ttu-id="960d3-171">[5 分間のクイック スタート](/office/dev/add-ins/)を完了することで、Excel、OneNote、Outlook、PowerPoint、Project、または Word 用の基本的なアドインを簡単に作成することができます。</span><span class="sxs-lookup"><span data-stu-id="960d3-171">You can quickly create a basic add-in for Excel, OneNote, Outlook, PowerPoint, Project, or Word by completing a [5-minute quick start](/office/dev/add-ins/).</span></span> <span data-ttu-id="960d3-172">以前にクイック スタートを完了している場合で、より複雑なアドインを作成したい場合は、[チュートリアル](/office/dev/add-ins/)を試してみてください。</span><span class="sxs-lookup"><span data-stu-id="960d3-172">If you've previously completed a quick start and want to create a slightly more complex add-in, you should try the [tutorial](/office/dev/add-ins/).</span></span>

### <a name="learn-more"></a><span data-ttu-id="960d3-173">詳細情報</span><span class="sxs-lookup"><span data-stu-id="960d3-173">Learn more</span></span>

<span data-ttu-id="960d3-174">Office アドインの開発、テスト、公開の詳細については、このドキュメントを参照してください。</span><span class="sxs-lookup"><span data-stu-id="960d3-174">Learn more about developing, testing, and publishing Office Add-ins by exploring this documentation.</span></span>

> [!TIP]
> <span data-ttu-id="960d3-175">どのようなアドインを構築する場合でも、このドキュメントの 「[開発ライフサイクル](../overview/core-concepts-office-add-ins.md)」セクションに記載する情報に加え、構築するアドインの種類に対応するアプリケーション固有のセクション (たとえば、[Excel](../excel/index.yml)) に記載する情報を使用してください。</span><span class="sxs-lookup"><span data-stu-id="960d3-175">For any add-in that you build, you'll use information in the [Development lifecycle](../overview/core-concepts-office-add-ins.md) section of this documentation, along with information in the application-specific section that corresponds to the type of add-in you're building (for example, [Excel](../excel/index.yml)).</span></span>

## <a name="see-also"></a><span data-ttu-id="960d3-176">関連項目</span><span class="sxs-lookup"><span data-stu-id="960d3-176">See also</span></span>

- [<span data-ttu-id="960d3-177">Office アドイン プラットフォームの概要</span><span class="sxs-lookup"><span data-stu-id="960d3-177">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
- [<span data-ttu-id="960d3-178">Microsoft 365 開発者プログラムについてご説明します</span><span class="sxs-lookup"><span data-stu-id="960d3-178">Learn about the Microsoft 365 Developer Program</span></span>](https://developer.microsoft.com/microsoft-365/dev-program)
- [<span data-ttu-id="960d3-179">Office アドインを設計する</span><span class="sxs-lookup"><span data-stu-id="960d3-179">Design Office Add-ins</span></span>](../design/add-in-design.md)
- [<span data-ttu-id="960d3-180">Office アドインのテストとデバッグ</span><span class="sxs-lookup"><span data-stu-id="960d3-180">Test and debug Office Add-ins</span></span>](../testing/test-debug-office-add-ins.md)
- [<span data-ttu-id="960d3-181">Office アドインの公開</span><span class="sxs-lookup"><span data-stu-id="960d3-181">Publish Office Add-ins</span></span>](../publish/publish.md)
