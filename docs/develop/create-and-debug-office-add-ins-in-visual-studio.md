---
title: Visual Studio での Office アドインの作成とデバッグ
description: ''
ms.date: 10/01/2018
ms.openlocfilehash: 224a4781b894e9bf165d279c30ca16d18bea956d
ms.sourcegitcommit: c400a220783b03a739449e2d3ff00bbffe5ec7c1
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/20/2018
ms.locfileid: "25681841"
---
# <a name="create-and-debug-office-add-ins-in-visual-studio"></a><span data-ttu-id="ad950-102">Visual Studio での Office アドインの作成とデバッグ</span><span class="sxs-lookup"><span data-stu-id="ad950-102">Create and debug Office Add-ins in Visual Studio</span></span>

<span data-ttu-id="ad950-p101">この記事では、Visual Studio を使用して、最初の Office アドインを作成する方法について説明します。ここに示す手順は Visual Studio 2017 に基づいたものです。別のバージョンの Visual Studio を使用している場合は、わずかに手順が異なることがあります。</span><span class="sxs-lookup"><span data-stu-id="ad950-p101">This article describes how to use Visual Studio to create your first Office Add-in. The steps in this article based on Visual Studio 2015. If you're using another version of Visual Studio, the procedures might vary slightly.</span></span>

> [!NOTE]
> <span data-ttu-id="ad950-106">OneNote 用のアドインを使い始めるには、「[最初の OneNote アドインをビルドする](../onenote/onenote-add-ins-getting-started.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="ad950-106">To get started with an add-in for OneNote, see [Build your first OneNote add-in](../onenote/onenote-add-ins-getting-started.md).</span></span>

## <a name="create-an-office-add-in-project-in-visual-studio"></a><span data-ttu-id="ad950-107">Visual Studio での Office アドイン プロジェクトの作成</span><span class="sxs-lookup"><span data-stu-id="ad950-107">Create an Office Add-in project in Visual Studio</span></span>


<span data-ttu-id="ad950-p102">作業を開始するために、[Office Developer Tools](https://www.visualstudio.com/features/office-tools-vs.aspx) がインストールされていることと、Microsoft Office のバージョンを確認します。[Office 365 Developer プログラム](https://developer.microsoft.com/office/dev-program)に参加するか、以下の手順を実行して[最新バージョン](../develop/install-latest-office-version.md)を取得できます。</span><span class="sxs-lookup"><span data-stu-id="ad950-p102">To get started, make sure you have the [Office Developer Tools](https://www.visualstudio.com/features/office-tools-vs.aspx) installed, and a version of Microsoft Office. You can join the [Office 365 Developer Program](https://developer.microsoft.com/office/dev-program), or follow these instructions to get the [latest version](../develop/install-latest-office-version.md).</span></span>

1. <span data-ttu-id="ad950-110">[Visual Studio] メニュー バーで、**[ファイル]**  >  **[新規作成]**  >  **[プロジェクト]** の順に選択します。</span><span class="sxs-lookup"><span data-stu-id="ad950-110">On the Visual Studio menu bar, choose  **File** > **New** > **Project**.</span></span>
2. <span data-ttu-id="ad950-111">プロジェクトの種類の一覧で、**[Visual C#]** または **[Visual Basic]** の下にある **[Office/SharePoint]** を展開し、**[Web アドイン]** を選択してからアドイン プロジェクトのいずれかを選択します。</span><span class="sxs-lookup"><span data-stu-id="ad950-111">In the list of project types under  **Visual C#** or **Visual Basic**, expand  **Office/SharePoint**, choose  **Web Add-ins**, and then select one of the Add-in projects.</span></span>
3. <span data-ttu-id="ad950-112">プロジェクトに名前を付けて、プロジェクトを作成するために **[OK]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="ad950-112">Name the project, and then choose  **OK** to create the project.</span></span>

<span data-ttu-id="ad950-113">Visual Studio 2017 で **[OK]** を選択した後、次のアドイン プロジェクト テンプレートに追加の選択肢があります 。</span><span class="sxs-lookup"><span data-stu-id="ad950-113">In Visual Studio 2017, the following add-in project templates have additional choices after you choose **OK**:</span></span>

<span data-ttu-id="ad950-114">**PowerPoint**</span><span class="sxs-lookup"><span data-stu-id="ad950-114">**PowerPoint**</span></span>
- <span data-ttu-id="ad950-115">作業ウィンドウ アドインを作成する **PowerPoint の新しい機能を追加** することができます。</span><span class="sxs-lookup"><span data-stu-id="ad950-115">You can choose to **Add new functionalities to PowerPoint** which creates a task pane add-in.</span></span>
- <span data-ttu-id="ad950-116">または **PowerPoint スライドにコンテンツを挿入する** コンテンツを追加で作成することもできます。</span><span class="sxs-lookup"><span data-stu-id="ad950-116">Or you can choose to **Insert content into PowerPoint slides** which creates a content add-in.</span></span>

<span data-ttu-id="ad950-117">**Excel**</span><span class="sxs-lookup"><span data-stu-id="ad950-117">**Excel**</span></span> 
- <span data-ttu-id="ad950-118">作業ウィンドウ アドインを作成する **Excel の新しい機能を追加** することができます。</span><span class="sxs-lookup"><span data-stu-id="ad950-118">You can choose to **Add new functionalities to Excel** which creates a task pane add-in.</span></span>
- <span data-ttu-id="ad950-119">または、 **Excel のスプレッドシートにコンテンツを挿入する** コンテンツを追加で作成することもできます。</span><span class="sxs-lookup"><span data-stu-id="ad950-119">Or you can choose to **Insert content into Excel spreadsheet** which creates a content add-in.</span></span>
    - <span data-ttu-id="ad950-120">コンテンツを追加で作成する場合の **基本的なアドインを** 最低限のスタート コードとコンテンツの追加のプロジェクトを作成する追加の選択肢があります。</span><span class="sxs-lookup"><span data-stu-id="ad950-120">If you create a content add-in, you have an additional choice of **Basic Add-in** which creates a content add-in project with minimal starter code.</span></span>
    - <span data-ttu-id="ad950-121">または **アドインがドキュメントのビジュアル化** を視覚化し、データにバインドする初期のコードを含むことができます。</span><span class="sxs-lookup"><span data-stu-id="ad950-121">Or you can choose a **Document Visualization Add-in** which includes starter code to visualize and bind to data.</span></span>

<span data-ttu-id="ad950-122">ウィザードを完了した後 Visual Studio の 2 つのプロジェクトを含むソリューションを作成します。</span><span class="sxs-lookup"><span data-stu-id="ad950-122">When you've completed the wizard, Visual Studio creates a solution for you that contains two projects.</span></span> <span data-ttu-id="ad950-123">Home.html の既定のページが開くことがわかります。</span><span class="sxs-lookup"><span data-stu-id="ad950-123">You'll see the default Home.html page open.</span></span>

|<span data-ttu-id="ad950-124">**プロジェクト**</span><span class="sxs-lookup"><span data-stu-id="ad950-124">**Project**</span></span>|<span data-ttu-id="ad950-125">**説明**</span><span class="sxs-lookup"><span data-stu-id="ad950-125">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="ad950-126">アドイン プロジェクト</span><span class="sxs-lookup"><span data-stu-id="ad950-126">Add-in project</span></span>|<span data-ttu-id="ad950-p104">アドインを記述するすべての設定を含む XML マニフェスト ファイルのみが含まれます。これらの設定は、Office ホストがアドインをアクティブ化するタイミングと、アドインの表示場所を決定するのに役立ちます。すぐにプロジェクトを実行し、アドインを使用できるように、Visual Studio によってこのファイルのコンテンツが生成されます。これらの設定は、マニフェスト エディターを使用していつでも変更できます。</span><span class="sxs-lookup"><span data-stu-id="ad950-p104">Contains only an XML manifest file, which contains all the settings that describe your add-in. These settings help the Office host determine when your add-in should be activated and where the add-in should appear. Visual Studio generates the contents of this file for you so that you can run the project and use your add-in immediately. You change these settings any time by using the Manifest editor.</span></span>|
|<span data-ttu-id="ad950-131">Web アプリケーション プロジェクト</span><span class="sxs-lookup"><span data-stu-id="ad950-131">Web application project</span></span>|<span data-ttu-id="ad950-132">Office 対応 HTML および JavaScript ページを開発するために必要なすべてのファイルおよびファイル参照を含む、アドインのコンテンツ ページが含まれています。</span><span class="sxs-lookup"><span data-stu-id="ad950-132">Contains the content pages of your add-in, including all the files and file references that you need to develop Office-aware HTML and JavaScript pages.</span></span> <span data-ttu-id="ad950-133">アドインを開発している間、Visual Studio はローカル IIS サーバー上の Web アプリケーションをホストします。</span><span class="sxs-lookup"><span data-stu-id="ad950-133">While you develop your add-in, Visual Studio hosts the web application on your local IIS server.</span></span> <span data-ttu-id="ad950-134">発行する準備ができたら、このプロジェクトをホストするサーバーを検索する必要があります。</span><span class="sxs-lookup"><span data-stu-id="ad950-134">When you're ready to publish, you'll have to find a server to host this project.</span></span> <span data-ttu-id="ad950-135">ASP.NET Web アプリケーション プロジェクト の詳細については、「[ ASP.NET Web プロジェクト](http://msdn.microsoft.com/library/cdcd712f-96b0-4165-8b5d-9d0566650a28%28Office.15%29.aspx) 」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="ad950-135">To learn more about ASP.NET web application projects, see [ASP.NET Web Projects](http://msdn.microsoft.com/library/cdcd712f-96b0-4165-8b5d-9d0566650a28%28Office.15%29.aspx).</span></span>|

## <a name="modify-your-add-in-settings"></a><span data-ttu-id="ad950-136">アドイン設定の変更</span><span class="sxs-lookup"><span data-stu-id="ad950-136">Modify your add-in settings</span></span>


<span data-ttu-id="ad950-137">アドインの設定を変更するには、プロジェクトの XML マニフェスト ファイルを編集します。</span><span class="sxs-lookup"><span data-stu-id="ad950-137">To modify the settings of your add-in, edit the XML manifest file of the project.</span></span> <span data-ttu-id="ad950-138">[**ソリューション エクスプローラー**] で、アドイン プロジェクト ノードを展開し、XML マニフェストを格納するフォルダーを展開して、XML マニフェストを選択します。</span><span class="sxs-lookup"><span data-stu-id="ad950-138">In  **Solution Explorer**, expand the add-in project node, expand the folder that contains the XML manifest, and choose the XML manifest.</span></span> <span data-ttu-id="ad950-139">ファイル内の任意の要素をポイントして、要素の目的を説明するヒントを表示できます。</span><span class="sxs-lookup"><span data-stu-id="ad950-139">You can point to any element in the file to view a tooltip that describes the purpose of the element.</span></span> <span data-ttu-id="ad950-140">マニフェスト ファイルの詳細については、「[Office アドイン XML マニフェスト](../develop/add-in-manifests.md)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="ad950-140">For more information about the manfiest file, see [Office Add-ins XML manifest](../develop/add-in-manifests.md).</span></span>


## <a name="develop-the-contents-of-your-add-in"></a><span data-ttu-id="ad950-141">アドインのコンテンツの開発</span><span class="sxs-lookup"><span data-stu-id="ad950-141">Develop the contents of your add-in</span></span>

<span data-ttu-id="ad950-142">アドイン プロジェクトはアドインを説明する設定を変更でき、Web アプリケーションはアドインに表示されるコンテンツを提供します。</span><span class="sxs-lookup"><span data-stu-id="ad950-142">While the add-in project lets you modify the settings that describe your add-in, the web application provides the content that appears in the add-in.</span></span> 

<span data-ttu-id="ad950-143">Web アプリケーション プロジェクトには、既定の HTML ページとを開始するのに使用できる JavaScript ファイルが含まれています。</span><span class="sxs-lookup"><span data-stu-id="ad950-143">The web application project contains a default HTML page and JavaScript file that you can use to get started.</span></span> <span data-ttu-id="ad950-144">これらのファイルには、Office 用の JavaScript API を含む他の JavaScript ライブラリへの参照が含まれています。</span><span class="sxs-lookup"><span data-stu-id="ad950-144">These files are convenient because they contain references to other JavaScript libraries including the JavaScript API for Office.</span></span> <span data-ttu-id="ad950-145">これらのファイルを更新し、さらに HTML と JavaScript ファイルを追加することによって、アドインを開発できます。</span><span class="sxs-lookup"><span data-stu-id="ad950-145">You can develop your add-in by updating these files, and adding more HTML and JavaScript files.</span></span> <span data-ttu-id="ad950-146">次の表は、既定の HTML や JavaScript ファイルについて説明します。</span><span class="sxs-lookup"><span data-stu-id="ad950-146">The following table describes default HTML and JavaScript files.</span></span>

> [!NOTE]
> <span data-ttu-id="ad950-147">Web プロジェクトのルート フォルダーで **ホーム** フォルダーを使用してプロジェクト テンプレートの種類に応じて次の表のファイルがあります。</span><span class="sxs-lookup"><span data-stu-id="ad950-147">The files in the table below may be in the root folder of the web project, or the **Home** folder depending on the type of project template you used.</span></span>

|<span data-ttu-id="ad950-148">**ファイル**</span><span class="sxs-lookup"><span data-stu-id="ad950-148">**File**</span></span>|<span data-ttu-id="ad950-149">**説明**</span><span class="sxs-lookup"><span data-stu-id="ad950-149">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="ad950-150">**Home.html**</span><span class="sxs-lookup"><span data-stu-id="ad950-150">**Home.html**</span></span>|<span data-ttu-id="ad950-151">アドインの既定の HTML ページです。</span><span class="sxs-lookup"><span data-stu-id="ad950-151">The default HTML page of the add-in.</span></span> <span data-ttu-id="ad950-152">アクティブ化すると、ドキュメント、電子メール メッセージ、または予定アイテムでは、アドイン内の最初のページとしてこのページが表示されます。</span><span class="sxs-lookup"><span data-stu-id="ad950-152">This page appears as the first page inside of the add-in when it is activated in a document, email message or appointment item.</span></span> <span data-ttu-id="ad950-153">このファイルには、すべてのファイル参照を開始する必要があるが含まれています。</span><span class="sxs-lookup"><span data-stu-id="ad950-153">This file is convenient because it contains all of the file references that you need to get started.</span></span> <span data-ttu-id="ad950-154">このファイルに HTML コードを追加することによって、アドインの開発を開始できます。</span><span class="sxs-lookup"><span data-stu-id="ad950-154">You can start developing your add-in by adding your HTML code to this file.</span></span>|
|<span data-ttu-id="ad950-155">**Home.js**</span><span class="sxs-lookup"><span data-stu-id="ad950-155">**Home.js**</span></span>|<span data-ttu-id="ad950-156">Home.html ページに関連付けられた JavaScript ファイルです。</span><span class="sxs-lookup"><span data-stu-id="ad950-156">The JavaScript file associated with the Home.html page.</span></span> <span data-ttu-id="ad950-157">Home.js ファイルで Home.html ページの動作に固有のコードを配置することができます。</span><span class="sxs-lookup"><span data-stu-id="ad950-157">You can place any code that is specific to the behavior of the Home.html page in the Home.js file.</span></span> <span data-ttu-id="ad950-158">Home.js ファイルには、開始するためのいくつかのコード例が含まれています。</span><span class="sxs-lookup"><span data-stu-id="ad950-158">The Home.js file contains some example code to get you started.</span></span>|
|<span data-ttu-id="ad950-159">**home.css**</span><span class="sxs-lookup"><span data-stu-id="ad950-159">**Home.css**</span></span>|<span data-ttu-id="ad950-160">アドインに適用する既定のスタイルを定義します。</span><span class="sxs-lookup"><span data-stu-id="ad950-160">Defines the default styles to apply to your add-in.</span></span> <span data-ttu-id="ad950-161">デザインとスタイルの Office UI のファブリックを使用することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="ad950-161">We recommend using the Office UI Fabric for design and styles.</span></span> <span data-ttu-id="ad950-162">詳細については、 [Office アドインの Office UI Fabric](../design/office-ui-fabric.md)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="ad950-162">For information about Office UI Fabric JS, see [Use Office UI Fabric in Office Add-ins](../design/office-ui-fabric.md).</span></span>|

> [!NOTE]
> <span data-ttu-id="ad950-163">これらのファイルを使用する必要はありません。</span><span class="sxs-lookup"><span data-stu-id="ad950-163">You don't have to use these files.</span></span> <span data-ttu-id="ad950-164">他のファイルをプロジェクトに自由に追加し、代わりに使用することができます。</span><span class="sxs-lookup"><span data-stu-id="ad950-164">Feel free to add other files to the project and use those instead.</span></span> <span data-ttu-id="ad950-165">別の HTML ファイルをアドインの最初のページとして表示する場合は、マニフェスト エディターを開き、およびファイルの名前に **SourceLocation** プロパティを設定します。</span><span class="sxs-lookup"><span data-stu-id="ad950-165">If you want another HTML file to appear as the initial page of the add-in, open the manifest editor, and then point the  **SourceLocation** property to the name of the file.</span></span>

## <a name="debug-your-add-in"></a><span data-ttu-id="ad950-166">アドインのデバッグ</span><span class="sxs-lookup"><span data-stu-id="ad950-166">Debug your add-in</span></span>

<span data-ttu-id="ad950-167">Visual Studio のビルドを提供して、アドインのデバッグを支援するためのプロパティをデバッグします。</span><span class="sxs-lookup"><span data-stu-id="ad950-167">Visual Studio provides build and debug properties to assist with debugging your add-in.</span></span>

### <a name="review-the-build-and-debug-properties"></a><span data-ttu-id="ad950-168">ビルドおよびデバッグ プロパティの確認</span><span class="sxs-lookup"><span data-stu-id="ad950-168">Review the build and debug properties</span></span>

<span data-ttu-id="ad950-p112">ソリューションを起動する前に、Visual Studio で目的のホスト アプリケーションが開けることを確認します。この情報は、アドインのビルドとデバッグに関連する他のプロパティと共に、プロジェクトのプロパティ ページに表示されます。</span><span class="sxs-lookup"><span data-stu-id="ad950-p112">Before you start the solution, verify that Visual Studio will open the host application that you want. That information appears in the property pages of the project along with several other properties that relate to building and debugging the add-in.</span></span>

### <a name="to-open-the-property-pages-of-a-project"></a><span data-ttu-id="ad950-171">プロジェクトのプロパティ ページを開くには</span><span class="sxs-lookup"><span data-stu-id="ad950-171">To open the property pages of a project</span></span>

1. <span data-ttu-id="ad950-172">**ソリューション エクスプ ローラー**では、Web プロジェクトではなく、基本的なアドイン プロジェクトを選択します。</span><span class="sxs-lookup"><span data-stu-id="ad950-172">In  **Solution Explorer**, choose the basic add-in project (not the Web project).</span></span>    
2. <span data-ttu-id="ad950-173">メニュー バーで、[ **表示**] >   [ **プロパティ ウィンドウ**] の順に選択します。</span><span class="sxs-lookup"><span data-stu-id="ad950-173">On the menu bar, choose  View,  Properties Window.</span></span>
    
<span data-ttu-id="ad950-174">次の表に、プロジェクトのプロパティを示します。</span><span class="sxs-lookup"><span data-stu-id="ad950-174">The following table describes the properties of the project.</span></span>



|<span data-ttu-id="ad950-175">**プロパティ**</span><span class="sxs-lookup"><span data-stu-id="ad950-175">**Property**</span></span>|<span data-ttu-id="ad950-176">**説明**</span><span class="sxs-lookup"><span data-stu-id="ad950-176">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="ad950-177">**開始動作**</span><span class="sxs-lookup"><span data-stu-id="ad950-177">**Start Action**</span></span>|<span data-ttu-id="ad950-178">Office デスクトップ クライアントまたは指定のブラウザー内の Office Online クライアントのどちらでアドインをデバッグするか指定します。</span><span class="sxs-lookup"><span data-stu-id="ad950-178">Specifies whether to debug your add-in in an Office desktop client or in an Office Online client in the specified browser.</span></span>|
|<span data-ttu-id="ad950-179">**開始ドキュメント** (コンテンツ アドインと作業ウィンドウ アドインのみ)</span><span class="sxs-lookup"><span data-stu-id="ad950-179">**Start Document** (Content and task pane add-ins only)</span></span>|<span data-ttu-id="ad950-180">プロジェクトの開始時に開くドキュメントを指定します。</span><span class="sxs-lookup"><span data-stu-id="ad950-180">Specifies what document to open when you start the project.</span></span>|
|<span data-ttu-id="ad950-181">**Web プロジェクト**</span><span class="sxs-lookup"><span data-stu-id="ad950-181">**Web Project**</span></span>|<span data-ttu-id="ad950-182">アドインに関連付けられている Web プロジェクトの名前を指定します。</span><span class="sxs-lookup"><span data-stu-id="ad950-182">Specifies the name of the web project associated with the add-in.</span></span>|
|<span data-ttu-id="ad950-183">**電子メール アドレス** (Outlook アドインのみ)</span><span class="sxs-lookup"><span data-stu-id="ad950-183">**Email Address** (Outlook add-ins only)</span></span>|<span data-ttu-id="ad950-184">Outlook アドインのテストに使用する Exchange Server か Exchange Online のユーザー アカウントの電子メール アドレスを指定します。</span><span class="sxs-lookup"><span data-stu-id="ad950-184">Specifies the email address of the user account in Exchange Server or Exchange Online that you want to test your Outlook add-in with.</span></span>|
|<span data-ttu-id="ad950-185">**EWS の URL** (Outlook アドインのみ)</span><span class="sxs-lookup"><span data-stu-id="ad950-185">**EWS Url** (Outlook add-ins only)</span></span>|<span data-ttu-id="ad950-186">Exchange Web サービスの URL (例: https://www.contoso.com/ews/exchange.aspx)。</span><span class="sxs-lookup"><span data-stu-id="ad950-186">Exchange Web service URL (For example: https://www.contoso.com/ews/exchange.aspx).</span></span> |
|<span data-ttu-id="ad950-187">**OWA の URL** (Outlook アドインのみ)</span><span class="sxs-lookup"><span data-stu-id="ad950-187">**OWA Url** (Outlook add-ins only)</span></span>|<span data-ttu-id="ad950-188">Outlook Web App の URL (例: https://www.contoso.com/owa)。</span><span class="sxs-lookup"><span data-stu-id="ad950-188">Outlook Web App URL (For example: https://www.contoso.com/owa).</span></span>|
|<span data-ttu-id="ad950-189">**ユーザー名** (Outlook アドインのみ)</span><span class="sxs-lookup"><span data-stu-id="ad950-189">**User name** (Outlook add-ins only)</span></span>|<span data-ttu-id="ad950-190">Exchange Server または Exchange Online のユーザー アカウントの名前を指定します。</span><span class="sxs-lookup"><span data-stu-id="ad950-190">Specifies the name of your user account in Exchange Server or Exchange Online.</span></span>|
|<span data-ttu-id="ad950-191">**プロジェクト ファイル**</span><span class="sxs-lookup"><span data-stu-id="ad950-191">**Project File**</span></span>|<span data-ttu-id="ad950-192">ビルド、構成、およびその他のプロジェクト情報が含まれているファイルの名前を指定します。</span><span class="sxs-lookup"><span data-stu-id="ad950-192">Specifies the name of the file containing build, configuration, and other information about the project.</span></span>|
|<span data-ttu-id="ad950-193">**プロジェクト フォルダー**</span><span class="sxs-lookup"><span data-stu-id="ad950-193">**Project Folder**</span></span>|<span data-ttu-id="ad950-194">プロジェクト ファイルの場所です。</span><span class="sxs-lookup"><span data-stu-id="ad950-194">The location of the project file.</span></span>|

### <a name="use-an-existing-document-to-debug-the-add-in-content-and-task-pane-add-ins-only"></a><span data-ttu-id="ad950-195">既存のドキュメントを使用してアドインをデバッグする (コンテンツ アドインと作業ウィンドウ アドインのみ)</span><span class="sxs-lookup"><span data-stu-id="ad950-195">Use an existing document to debug the add-in (content and task pane add-ins only)</span></span>

<span data-ttu-id="ad950-p113">アドイン プロジェクトにドキュメントを追加できます。アドインで使用するテスト データを含むドキュメントがある場合、プロジェクトの開始時に Visual Studio によってそのドキュメントが開かれます。</span><span class="sxs-lookup"><span data-stu-id="ad950-p113">You can add documents to the add-in project. If you have a document that contains test data that you want to use with your add-in, Visual Studio opens that document for you when you start the project.</span></span>

### <a name="to-use-an-existing-document-to-debug-the-add-in"></a><span data-ttu-id="ad950-198">既存のドキュメントを使用してアドインをデバッグするには</span><span class="sxs-lookup"><span data-stu-id="ad950-198">To use an existing document to debug the add-in</span></span>

1. <span data-ttu-id="ad950-199">**ソリューション エクスプローラ**で、アドイン プロジェクト フォルダーを選択します。</span><span class="sxs-lookup"><span data-stu-id="ad950-199">In  **Solution Explorer**, choose the add-in project folder.</span></span>
    
    > [!NOTE]
    > <span data-ttu-id="ad950-200">Web アプリケーション プロジェクトではなく、アドイン プロジェクトを選択します。</span><span class="sxs-lookup"><span data-stu-id="ad950-200">Choose the add-in project and not the web application project.</span></span>

2. <span data-ttu-id="ad950-201">**[プロジェクト]** メニューで、**[既存の項目の追加]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="ad950-201">On the  **Project** menu, choose **Add Existing Item**.</span></span>
    
3. <span data-ttu-id="ad950-202">[ **既存の項目の追加**] ダイアログ ボックスで、追加するドキュメントを探して選択します。</span><span class="sxs-lookup"><span data-stu-id="ad950-202">In the  **Add Existing Item** dialog box, locate and select the document that you want to add.</span></span>
    
4. <span data-ttu-id="ad950-203">[ **追加**] を選択して、ドキュメントをプロジェクトに追加します。</span><span class="sxs-lookup"><span data-stu-id="ad950-203">Choose the  **Add** button to add the document to your project.</span></span>
    
5. <span data-ttu-id="ad950-204">**ソリューション エクスプローラ**で、アドイン プロジェクト フォルダーを選択します。</span><span class="sxs-lookup"><span data-stu-id="ad950-204">In  **Solution Explorer**, choose the add-in project folder.</span></span>
6. <span data-ttu-id="ad950-205">メニュー バーで、[ **表示**] >  [ **プロパティ ウィンドウ**] の順に選択します。</span><span class="sxs-lookup"><span data-stu-id="ad950-205">On the menu bar, choose  View,  Properties Window.</span></span>
7. <span data-ttu-id="ad950-206">[プロパティ] ウィンドウでは、 **ドキュメントの開始** ] ボックスの一覧を選択し、プロジェクトに追加したドキュメントを選択します。</span><span class="sxs-lookup"><span data-stu-id="ad950-206">In the  **Start Document** list, choose the document that you added to the project, and then choose the OK button to close the property pages.</span></span> <span data-ttu-id="ad950-207">今すぐプロジェクトを構成して、既存の文書でアドインを起動します。</span><span class="sxs-lookup"><span data-stu-id="ad950-207">Now the project is configured to start your add-in in your existing document.</span></span>

### <a name="start-the-solution"></a><span data-ttu-id="ad950-208">ソリューションの起動</span><span class="sxs-lookup"><span data-stu-id="ad950-208">Start the solution</span></span>

<span data-ttu-id="ad950-209">**デバッグ**を選択して、メニュー バーからソリューションを開始 > **デバッグを開始**します。</span><span class="sxs-lookup"><span data-stu-id="ad950-209">Start the solution from the menu bar by choosing **Debug** > **Start Debugging**.</span></span> <span data-ttu-id="ad950-210">Visual Studio は自動的にソリューションをビルドし、アドインをホストするための Office を起動します。</span><span class="sxs-lookup"><span data-stu-id="ad950-210">Visual Studio will automatically build the solution and start Office to host your add-in.</span></span>

<span data-ttu-id="ad950-211">Visual Studio プロジェクトをビルドするときは、次のタスクを実行します。</span><span class="sxs-lookup"><span data-stu-id="ad950-211">When Visual Studio builds the project it performs the following tasks:</span></span>

1. <span data-ttu-id="ad950-p116">XML マニフェスト ファイルのコピーを作成し、それを  _プロジェクト名_\Output ディレクトリに追加します。このコピーは、Visual Studio を起動してアドインをデバッグするときにホスト アプリケーションで使用されます。</span><span class="sxs-lookup"><span data-stu-id="ad950-p116">Creates a copy of the XML manifest file and adds it to  _ProjectName_\Output directory. The host application consumes this copy when you start Visual Studio and debug the add-in.</span></span>
    
2. <span data-ttu-id="ad950-214">アドインをホスト アプリケーションに表示するための一連のレジストリ エントリをコンピューターに作成します。</span><span class="sxs-lookup"><span data-stu-id="ad950-214">Creates a set of registry entries on your computer that enable the add-in to appear in the host application.</span></span>
    
3. <span data-ttu-id="ad950-215">Web アプリケーション プロジェクトをビルドし、ローカルの IIS Web サーバー (http://localhost)) に展開します。</span><span class="sxs-lookup"><span data-stu-id="ad950-215">Builds the web application project, and then deploys it to the local IIS web server (http://localhost).</span></span> 
    
<span data-ttu-id="ad950-216">次に、Visual Studio は次の操作を実行します。</span><span class="sxs-lookup"><span data-stu-id="ad950-216">Next, Visual Studio does the following:</span></span>

1. <span data-ttu-id="ad950-217">~remoteAppUrlトークンを開始ページの完全修飾アドレス (例: http://localhost/MyAgave.html)) で置き換えることによって、XML マニフェストファイルの  [SourceLocation](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/sourcelocation?view=office-js)  要素を変更します。</span><span class="sxs-lookup"><span data-stu-id="ad950-217">Modifies the [SourceLocation](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/sourcelocation?view=office-js) element of the XML manifest file by replacing the ~remoteAppUrl token with the fully qualified address of the start page (for example, http://localhost/MyAgave.html).</span></span>
    
2. <span data-ttu-id="ad950-218">IIS Express で Web アプリケーション プロジェクトを起動します。</span><span class="sxs-lookup"><span data-stu-id="ad950-218">Starts the web application project in IIS Express.</span></span>
    
3. <span data-ttu-id="ad950-219">ホスト アプリケーションを開きます。</span><span class="sxs-lookup"><span data-stu-id="ad950-219">Opens the host application.</span></span> 
    
<span data-ttu-id="ad950-p117">プロジェクトをビルドする際、Visual Studio は **出力**ウィンドウに検証エラーを表示しません。Visual Studio は、エラーと警告を、発生時に  **ERRORLIST** ウィンドウ内で報告します。Visual Studio は、コードおよびテキスト エディター内で検証エラーを別の色の波形の下線 (波線と呼びます) で示します。このようなマークにより、Visual Studio がコード内で検出した問題が通知されます。詳細については、「 [コードおよびテキスト エディター](https://msdn.microsoft.com/library/se2f663y(v=vs.140).aspx)」を参照してください。検証を有効化または無効化する方法の詳細については、次のトピックを参照してください。</span><span class="sxs-lookup"><span data-stu-id="ad950-p117">Visual Studio doesn't show validation errors in the  **OUTPUT** window when you build the project. Visual Studio reports errors and warnings in the **ERRORLIST** window as they occur. Visual Studio also reports validation errors by showing wavy underlines (known as squiggles) of different colors in the code and text editor. These marks notify you of problems that Visual Studio detected in your code. For more information, see [Code and Text Editor](https://msdn.microsoft.com/library/se2f663y(v=vs.140).aspx). For more information about how to enable or disable validation, see:</span></span> 

- <span data-ttu-id="ad950-226">[[オプション]、[テキスト エディター]、[JavaScript]、[IntelliSense]](https://docs.microsoft.com/visualstudio/ide/reference/options-text-editor-javascript-intellisense?view=vs-2015)</span><span class="sxs-lookup"><span data-stu-id="ad950-226">[Options, Text Editor, JavaScript, IntelliSense](https://docs.microsoft.com/visualstudio/ide/reference/options-text-editor-javascript-intellisense?view=vs-2015)</span></span>
    
- <span data-ttu-id="ad950-227">[方法: Visual Web Developer で HTML 編集用の検証オプションを設定する](https://msdn.microsoft.com/library/0byxkfet(v=vs.100).aspx)</span><span class="sxs-lookup"><span data-stu-id="ad950-227">[How to: Set Validation Options for HTML Editing in Visual Web Developer](https://msdn.microsoft.com/library/0byxkfet(v=vs.100).aspx)</span></span>
    
- <span data-ttu-id="ad950-228">[[検証] ([オプション] ダイアログ ボックス - [テキスト エディター] - [CSS])](https://msdn.microsoft.com/library/se2f663y(v=vs.140).aspx)</span><span class="sxs-lookup"><span data-stu-id="ad950-228">[CSS, see Validation, CSS, Text Editor, Options Dialog Box](https://msdn.microsoft.com/library/se2f663y(v=vs.140).aspx)</span></span>
    
<span data-ttu-id="ad950-229">プロジェクト内の XML マニフェスト ファイルの検証ルールを確認するには、「[Office アドインの XML マニフェスト](../develop/add-in-manifests.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="ad950-229">To review the validation rules of the XML manifest file in your project, see [Office Add-ins XML manifest](../develop/add-in-manifests.md).</span></span>

### <a name="show-an-add-in-in-excel-or-word-and-step-through-your-code"></a><span data-ttu-id="ad950-230">Excel または Word でアドインを表示して、コードをステップ実行</span><span class="sxs-lookup"><span data-stu-id="ad950-230">Show an add-in in Excel, Word, or Project and step through your code</span></span>

<span data-ttu-id="ad950-231">アドイン プロジェクトの **開始ドキュメント** プロパティを Excel または Word に設定した場合、Visual Studio はドキュメントを新規作成し、アドインが表示されます。</span><span class="sxs-lookup"><span data-stu-id="ad950-231">If you set the  **Start Document** property of the add-in project to Excel or Word, Visual Studio creates a new document and the add-in appears.</span></span> <span data-ttu-id="ad950-232">アドイン プロジェクトの\*\*  開始ドキュメント\*\* プロパティを既存のドキュメントを使用するように設定した場合、Visual Studio はドキュメントを開きますが、アドインは手動で挿入する必要があります。</span><span class="sxs-lookup"><span data-stu-id="ad950-232">If you set the **Start Document** property of the add-in project to use an existing document, Visual Studio opens the document, but you have to insert the add-in manually.</span></span>

1. <span data-ttu-id="ad950-233">Excel または Word の [ **挿入** ] タブで、ドロップダウン リストの **[アドイン]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="ad950-233">In Excel or Word, on the  **Insert** tab, choose the **My Add-ins** drop down list.</span></span> <span data-ttu-id="ad950-234">ボタン自体ではなくドロップダウン リストから[**Office の アドイン** ] ダイアログを開きます。</span><span class="sxs-lookup"><span data-stu-id="ad950-234">Choose the list from the drop-down arrow, not the button itself which opens the **Office Add-ins** dialog.</span></span>
2. <span data-ttu-id="ad950-235">**アドインの開発者**下の、アドインを選択します。</span><span class="sxs-lookup"><span data-stu-id="ad950-235">Under **Developer Add-ins**, choose your add-in.</span></span>

<span data-ttu-id="ad950-236">Visual Studio は、ブレーク ポイントを設定し、アドインと対話し、HTML や JavaScript ファイルにコードをステップ実行します。</span><span class="sxs-lookup"><span data-stu-id="ad950-236">In Visual Studio, you can then set break-points. Then, as you interact with your add-in and step through the code in your HTML, JavaScript, and C# or VB code files.</span></span>

### <a name="show-the-outlook-add-in-in-outlook-and-step-through-your-code"></a><span data-ttu-id="ad950-237">Outlook で Outlook のアドインを表示し、コードをステップ実行する</span><span class="sxs-lookup"><span data-stu-id="ad950-237">Show the Outlook add-in in Outlook and step through your code</span></span>

<span data-ttu-id="ad950-238">Outlook でアドインを表示するには、電子メール メッセージまたは予定アイテムを開きます。</span><span class="sxs-lookup"><span data-stu-id="ad950-238">To view the add-in in Outlook, open an email message or appointment item.</span></span>

<span data-ttu-id="ad950-p120">Outlook は、アクティブ化の基準を満たしていれば、アイテムの アドイン をアクティブ化します。アドイン バーが [インスペクタ] ウィンドウまたは閲覧ウィンドウの上部に表示され、Outlook アドインがアドイン バーにボタンとして表示されます。アドインにアドイン コマンドがある場合は、リボンの既定のタブまたは指定されたカスタム タブのいずれかにボタンが表示され、アドイン バーにはアドインは表示されません。</span><span class="sxs-lookup"><span data-stu-id="ad950-p120">Outlook activates the add-in for the item as long as the activation criteria are met. The add-in bar appears at the top of the Inspector window or Reading Pane, and your Outlook add-in appears as a button in the add-in bar. If your add-in has an add-in command, a button will appear in the ribbon, either in the default tab or a specified custom tab, and the add-in will not appear in the add-in bar.</span></span>

<span data-ttu-id="ad950-242">Outlook アドインを表示するには、Outlook アドインのボタンを選択します。</span><span class="sxs-lookup"><span data-stu-id="ad950-242">To view your Outlook add-in, choose the button for your Outlook add-in.</span></span>

<span data-ttu-id="ad950-243">Visual Studio は、ブレーク ポイントを設定し、アドインと対話し、HTML や JavaScript ファイルにコードをステップ実行します。</span><span class="sxs-lookup"><span data-stu-id="ad950-243">In Visual Studio, you can then set break-points. Then, as you interact with your add-in and step through the code in your HTML, JavaScript, and C# or VB code files.</span></span>

<span data-ttu-id="ad950-p121">また、コードを変更してから、Office アドイン を終了してプロジェクトを再度起動しなくても、Outlook アドインへの影響を確認することができます。Outlook で Outlook アドインのショートカット メニューを開き、 **[再読み込み]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="ad950-p121">You can also change your code and review the effects of those changes in your Outlook add-in without having to close the Office Add-in and start the project again. In Outlook, just open the shortcut menu for the Outlook add-in, and then choose  **Reload**.</span></span>


### <a name="modify-code-and-continue-to-debug-the-add-in-without-having-to-start-the-project-again"></a><span data-ttu-id="ad950-246">コードを変更した後、プロジェクトを再び開始することなくアドインのデバッグを続行する</span><span class="sxs-lookup"><span data-stu-id="ad950-246">Modify code and continue to debug the add-in without having to start the project again</span></span>

<span data-ttu-id="ad950-247">ホスト アプリケーションを終了してもう一度プロジェクトを開始することなく、コードを変更してアドインのこれらの変更の効果を確認できます。</span><span class="sxs-lookup"><span data-stu-id="ad950-247">You can change your code and review the effects of those changes in your add-in without having to close the host application and start the project again.</span></span> <span data-ttu-id="ad950-248">コードを変更して保存した後、アドインのショートカット メニューを開いて **再読み込み**を選択します。</span><span class="sxs-lookup"><span data-stu-id="ad950-248">After you change your code, open the shortcut menu for the add-in, and then choose  **Reload**.</span></span>
    

## <a name="next-steps"></a><span data-ttu-id="ad950-249">次の手順</span><span class="sxs-lookup"><span data-stu-id="ad950-249">Next steps</span></span>

- [<span data-ttu-id="ad950-250">Office アドインを展開し、発行する</span><span class="sxs-lookup"><span data-stu-id="ad950-250">Deploy and publish your Office Add-in</span></span>](../publish/publish.md)
    
