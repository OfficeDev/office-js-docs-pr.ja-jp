---
title: Visual Studio での Office アドインの作成とデバッグ
description: ''
ms.date: 03/14/2018
ms.openlocfilehash: 2e5c08a72ec97e26000d6ea7e53dd1d0f2c9e6dc
ms.sourcegitcommit: 30435939ab8b8504c3dbfc62fd29ec6b0f1a7d22
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/12/2018
ms.locfileid: "23945356"
---
# <a name="create-and-debug-office-add-ins-in-visual-studio"></a><span data-ttu-id="29763-102">Visual Studio での Office アドインの作成とデバッグ</span><span class="sxs-lookup"><span data-stu-id="29763-102">Create and debug Office Add-ins in Visual Studio</span></span>

<span data-ttu-id="29763-p101">この記事では、Visual Studio を使用して、最初の Office アドインを作成する方法について説明します。ここに示す手順は Visual Studio 2015 に基づいたものです。別のバージョンの Visual Studio を使用している場合は、わずかに手順が異なることがあります。</span><span class="sxs-lookup"><span data-stu-id="29763-p101">This article describes how to use Visual Studio to create your first Office Add-in. The steps in this article based on Visual Studio 2015. If you're using another version of Visual Studio, the procedures might vary slightly.</span></span>

> [!NOTE]
> <span data-ttu-id="29763-106">OneNote 用のアドインを使い始めるには、「[最初の OneNote アドインをビルドする](../onenote/onenote-add-ins-getting-started.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="29763-106">To get started with an add-in for OneNote, see [Build your first OneNote add-in](../onenote/onenote-add-ins-getting-started.md).</span></span>

## <a name="create-an-office-add-in-project-in-visual-studio"></a><span data-ttu-id="29763-107">Visual StudioでOfficeアドインプロジェクトを作成する</span><span class="sxs-lookup"><span data-stu-id="29763-107">Create an Office Add-in project in Visual Studio</span></span>


<span data-ttu-id="29763-p102">作業を開始するために、[Office Developer Tools](https://www.visualstudio.com/features/office-tools-vs.aspx) がインストールされていることと、Microsoft Office のバージョンを確認します。[Office 365 Developer プログラム](https://developer.microsoft.com/office/dev-program)に参加するか、以下の手順を実行して[最新バージョン](../develop/install-latest-office-version.md)を取得できます。</span><span class="sxs-lookup"><span data-stu-id="29763-p102">To get started, make sure you have the [Office Developer Tools](https://www.visualstudio.com/features/office-tools-vs.aspx) installed, and a version of Microsoft Office. You can join the [Office 365 Developer Program](https://developer.microsoft.com/office/dev-program), or follow these instructions to get the [latest version](../develop/install-latest-office-version.md).</span></span>


1. <span data-ttu-id="29763-110">[Visual Studio] メニュー バーで、**[ファイル]** > **[新規作成]** > **[プロジェクト]** の順に選択します。</span><span class="sxs-lookup"><span data-stu-id="29763-110">On the Visual Studio menu bar, choose  **File** > **New** > **Project**.</span></span>
    
2. <span data-ttu-id="29763-111">プロジェクトの種類の一覧で、**[Visual C#]** または **[Visual Basic]** の下にある **[Office/SharePoint]** を展開し、**[Web アドイン]** を選択してからアドイン プロジェクトのいずれかを選択します。</span><span class="sxs-lookup"><span data-stu-id="29763-111">In the list of project types under  **Visual C#** or **Visual Basic**, expand  **Office/SharePoint**, choose  **Web Add-ins**, and then select one of the Add-in projects.</span></span>  
    
3. <span data-ttu-id="29763-112">プロジェクトに名前を付けて、プロジェクトを作成するために **[OK]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="29763-112">Name the project, and then choose  **OK** to create the project.</span></span>
    
4. <span data-ttu-id="29763-p103">Visual Studio によってソリューションとその 2 つのプロジェクトが作成され、**ソリューション エクスプローラー**に表示されます。既定の Home.html ページが Visual Studio に開かれます。</span><span class="sxs-lookup"><span data-stu-id="29763-p103">Visual Studio creates a solution and its two projects appear in **Solution Explorer**. The default Home.html page opens in Visual Studio.</span></span>
    
<span data-ttu-id="29763-115">Visual Studio 2015 で、追加機能を反映するために、次に示す一部のアドイン プロジェクト テンプレートが更新されました。</span><span class="sxs-lookup"><span data-stu-id="29763-115">In Visual Studio 2015, some of the add-in project templates have been updated to reflect additional functionality:</span></span>


- <span data-ttu-id="29763-p104">コンテンツのアドインは、Excel スプレッドシートに加えて、Access ドキュメントと PowerPoint ドキュメントの本文に表示されます。[Basic Project] オプションを選択すると、最小限のスタート コードで基本のコンテンツ アドイン プロジェクトを作成できます。または、[Document Visualization Project] オプション (Access および Excel のみ) を選択すると、データの視覚化とバインドを行うためのスタート コードが組み込まれているフル機能のコンテンツ アドインを作成できます。</span><span class="sxs-lookup"><span data-stu-id="29763-p104">Content add-ins can appear in the body of Access and PowerPoint documents, in addition to Excel spreadsheets. You can also choose the Basic Project option to create a basic content add-in project with minimal starter code, or the Document Visualization Project option (for Access and Excel only) to create a more full-featured content add-in that includes starter code to visualize and bind to data.</span></span>
    
- <span data-ttu-id="29763-118">Outlook アドインには、電子メール メッセージや予定内にアドインを組み込むオプションだけでなく、電子メール メッセージや予定の閲覧時や新規作成時にアドインが使用可能かどうか指定するオプションも含まれています。</span><span class="sxs-lookup"><span data-stu-id="29763-118">Outlook add-ins include options not just for including your add-in in email messages or appointments, but also for specifying whether the add-in is available when an email message or appointment is being composed as well as read.</span></span>
    

> [!NOTE]
> <span data-ttu-id="29763-p105">Visual Studio では、ほとんどのオプションは、説明を読んで理解できますが、**[電子メール メッセージ]** チェック ボックスは例外です。このチェック ボックスは、メール アイテムだけでなく、会議出席依頼、返信、キャンセルでも表示される Outlook アドインを作成する場合に使用します。</span><span class="sxs-lookup"><span data-stu-id="29763-p105">In Visual Studio most options are understandable from their descriptions except for the  **Email Message** checkbox. Use that checkbox if you want to create an Outlook add-in that appears not just with mail items, but also with meeting requests, responses, and cancellations.</span></span>

<span data-ttu-id="29763-121">ウィザードの完了後、Visual Studio によって 2 つのプロジェクトを含むソリューションが作成されます。</span><span class="sxs-lookup"><span data-stu-id="29763-121">When you've completed the wizard, Visual Studio creates a solution for you that contains two projects.</span></span>



|<span data-ttu-id="29763-122">**プロジェクト**</span><span class="sxs-lookup"><span data-stu-id="29763-122">**Project**</span></span>|<span data-ttu-id="29763-123">**説明**</span><span class="sxs-lookup"><span data-stu-id="29763-123">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="29763-124">アドイン プロジェクト</span><span class="sxs-lookup"><span data-stu-id="29763-124">Add-in project</span></span>|<span data-ttu-id="29763-p106">アドインを記述するすべての設定を含む XML マニフェスト ファイルのみが含まれます。これらの設定は、Office ホストがアドインをアクティブ化するタイミングと、アドインの表示場所を決定するのに役立ちます。すぐにプロジェクトを実行し、アドインを使用できるように、Visual Studio によってこのファイルのコンテンツが生成されます。これらの設定は、マニフェスト エディターを使用していつでも変更できます。</span><span class="sxs-lookup"><span data-stu-id="29763-p106">Contains only an XML manifest file, which contains all the settings that describe your add-in. These settings help the Office host determine when your add-in should be activated and where the add-in should appear. Visual Studio generates the contents of this file for you so that you can run the project and use your add-in immediately. You change these settings any time by using the Manifest editor.</span></span>|
|<span data-ttu-id="29763-129">Web アプリケーション プロジェクト</span><span class="sxs-lookup"><span data-stu-id="29763-129">Web application project</span></span>|<span data-ttu-id="29763-p107">Office 対応の HTML および JavaScript ページを開発するために必要なすべてのファイルとファイル参照を含むアドインのコンテンツ ページが含まれます。アドインを開発している間、Visual Studio は Web アプリケーションをローカル IIS サーバー上でホストします。発行する準備が整ったら、このプロジェクトをホストするサーバーを見つける必要があります。ASP.NET Web アプリケーション プロジェクトの詳細については、「 [ASP.NET Web プロジェクト](http://msdn.microsoft.com/library/cdcd712f-96b0-4165-8b5d-9d0566650a28%28Office.15%29.aspx)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="29763-p107">Contains the content pages of your add-in, including all the files and file references that you need to develop Office-aware HTML and JavaScript pages. While you develop your add-in, Visual Studio hosts the web application on your local IIS server. When you're ready to publish, you'll have to find a server to host this project.To learn more about ASP.NET web application projects, see [ASP.NET Web Projects](http://msdn.microsoft.com/library/cdcd712f-96b0-4165-8b5d-9d0566650a28%28Office.15%29.aspx).</span></span>|

## <a name="modify-your-add-in-settings"></a><span data-ttu-id="29763-133">アドイン設定の変更</span><span class="sxs-lookup"><span data-stu-id="29763-133">Modify your add-in settings</span></span>


<span data-ttu-id="29763-p108">アドインの設定を変更するには、プロジェクトの XML マニフェスト ファイルを編集します。[**ソリューション エクスプローラー**] で、アドイン プロジェクト ノードを展開し、XML マニフェストを格納するフォルダーを展開して、XML マニフェストを選択します。ファイル内の任意の要素をポイントして、要素の目的を説明するヒントを表示できます。マニフェスト ファイルの詳細については、「[Office アドイン XML マニフェスト](../develop/add-in-manifests.md)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="29763-p108">To modify the settings of your add-in, edit the XML manifest file of the project. In  **Solution Explorer**, expand the add-in project node, expand the folder that contains the XML manifest, and choose the XML manifest. You can point to any element in the file to view a tooltip that describes the purpose of the element. For more information about the manfiest file, see [Office Add-ins XML manifest](../develop/add-in-manifests.md).</span></span>


## <a name="develop-the-contents-of-your-add-in"></a><span data-ttu-id="29763-138">アドインのコンテンツの開発</span><span class="sxs-lookup"><span data-stu-id="29763-138">Develop the contents of your add-in</span></span>


<span data-ttu-id="29763-139">アドイン プロジェクトはアドインを説明する設定を変更でき、Web アプリケーションはアドインに表示されるコンテンツを提供します。</span><span class="sxs-lookup"><span data-stu-id="29763-139">While the add-in project lets you modify the settings that describe your add-in, the web application provides the content that appears in the add-in.</span></span> 

<span data-ttu-id="29763-p109">Web アプリケーション プロジェクトには、作業の開始時に使用できる既定の HTML ページと JavaScript ファイルが含まれています。また、プロジェクトに追加するすべてのページに共通の JavaScript ファイルも含まれています。JavaScript API for Office などの他の JavaScript ライブラリへの参照が含まれるので、これらのファイルは便利です。</span><span class="sxs-lookup"><span data-stu-id="29763-p109">The web application project contains a default HTML page and JavaScript file that you can use to get started. The project also contains a JavaScript file that is common to all pages that you add to your project. These files are convenient because they contain references to other JavaScript libraries including the JavaScript API for Office.</span></span> 

<span data-ttu-id="29763-p110">アドインが高度になるにつれて、追加する HTML ファイルや JavaScript ファイルの数が多くなります。アドインと連動させるためにプロジェクト内の他のページに追加できる参照の種類の例として、既定の HTML ファイルと JavaScript ファイルのコンテンツがあります。次の表は、既定の HTML ファイルと JavaScript ファイルを示しています。</span><span class="sxs-lookup"><span data-stu-id="29763-p110">As your add-in becomes more sophisticated, you can add more HTML and JavaScript files. You can use the contents of the default HTML and JavaScript files as examples of the types of references you might want to add to other pages in your project to make them work with your add-in. The following table describes default HTML and JavaScript files.</span></span>



|<span data-ttu-id="29763-146">**ファイル**</span><span class="sxs-lookup"><span data-stu-id="29763-146">**File**</span></span>|<span data-ttu-id="29763-147">**説明**</span><span class="sxs-lookup"><span data-stu-id="29763-147">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="29763-148">**Home.html**</span><span class="sxs-lookup"><span data-stu-id="29763-148">**Home.html**</span></span>|<span data-ttu-id="29763-p111">アドインの既定の HTML ページで、プロジェクトの  **[Home]** フォルダーに存在します。アドインがドキュメント、電子メール メッセージ、または予定のアイテムでアクティブ化されると、このページがアドイン内の最初のページとして表示されます。このファイルには作業を開始する際に必要となるファイル参照がすべて含まれていて便利です。最初のアドインを作成する準備ができたら、このファイルに HTML コードを追加するだけで済みます。</span><span class="sxs-lookup"><span data-stu-id="29763-p111">Located in the  **Home** folder of the project, this is default HTML page of the add-in. This page appears as the first page inside of the add-in when it is activated in a document, email message or appointment item. This file is convenient because it contains all of the file references that you need to get started. When you are ready to create your first add-in, just add your HTML code to this file.</span></span>|
|<span data-ttu-id="29763-153">**Home.js**</span><span class="sxs-lookup"><span data-stu-id="29763-153">**Home.js**</span></span>|<span data-ttu-id="29763-p112">Home.js ページに関連付けられた JavaScript ファイルで、プロジェクトの  **[Home]** フォルダーに存在します。Home.js ファイルには Home.html ページの動作にとって固有のコードを組み込むことができます。Home.js ファイルには、作業の開始に利用できるサンプル コードが含まれています。</span><span class="sxs-lookup"><span data-stu-id="29763-p112">Located in the  **Home** folder of the project, this is the JavaScript file associated with the Home.js page. You can place any code that is specific to the behavior of the Home.html page in the Home.js file. The Home.js file contains some example code to get you started.</span></span>|
|<span data-ttu-id="29763-157">**App.js**</span><span class="sxs-lookup"><span data-stu-id="29763-157">**App.js**</span></span>|<span data-ttu-id="29763-p113">アドイン全体の既定の JavaScript ファイルで、プロジェクトの  **[Add-in]** フォルダーに存在します。App.js ファイルにはアドインの複数ページの動作にとって共通のコードを組み込むことができます。App.js ファイルには作業を開始する際に使用できるサンプル コードがいくつか含まれています。</span><span class="sxs-lookup"><span data-stu-id="29763-p113">Located in the  **Add-in** folder of the project, this is the default JavaScript file of the entire add-in. You can place code that is common to the behavior of multiple pages of your add-in in the App.js file. The App.js file contains some example code to get you started.</span></span>|

> [!NOTE]
> <span data-ttu-id="29763-p114">これらのファイルを必ず使用する必要はありません。他のファイルをプロジェクトに追加して代わりに使用してもかまいません。別の HTML ファイルをアドインの初期ページとして表示する場合は、マニフェスト エディターを開き、**SourceLocation** プロパティにそのファイルの名前を設定します。</span><span class="sxs-lookup"><span data-stu-id="29763-p114">You don't have to use these files. Feel free to add other files to the project and use those instead. If you want another HTML file to appear as the initial page of the add-in, open the manifest editor, and then point the  **SourceLocation** property to the name of the file.</span></span>


## <a name="debug-your-add-in"></a><span data-ttu-id="29763-164">アドインのデバッグ</span><span class="sxs-lookup"><span data-stu-id="29763-164">Debug your add-in</span></span>


<span data-ttu-id="29763-165">アドインを起動する準備ができたら、ビルドとデバッグに関連するプロパティを確認してください。確認が終了したら、ソリューションを起動します。</span><span class="sxs-lookup"><span data-stu-id="29763-165">When you are ready to start your add-in, review build and debug related properties, and then start the solution.</span></span>


### <a name="review-the-build-and-debug-properties"></a><span data-ttu-id="29763-166">ビルドおよびデバッグ プロパティの確認</span><span class="sxs-lookup"><span data-stu-id="29763-166">Review the build and debug properties</span></span>

<span data-ttu-id="29763-p115">ソリューションを起動する前に、Visual Studio で目的のホスト アプリケーションが開けることを確認します。この情報は、アドインのビルドとデバッグに関連する他のプロパティと共に、プロジェクトのプロパティ ページに表示されます。</span><span class="sxs-lookup"><span data-stu-id="29763-p115">Before you start the solution, verify that Visual Studio will open the host application that you want. That information appears in the property pages of the project along with several other properties that relate to building and debugging the add-in.</span></span>


### <a name="to-open-the-property-pages-of-a-project"></a><span data-ttu-id="29763-169">プロジェクトのプロパティ ページを開くには</span><span class="sxs-lookup"><span data-stu-id="29763-169">To open the property pages of a project</span></span>


1. <span data-ttu-id="29763-170">**ソリューション エクスプ ローラー**では、Web プロジェクトではなく、基本的なアドイン プロジェクトを選択します。</span><span class="sxs-lookup"><span data-stu-id="29763-170">In  **Solution Explorer**, choose the basic add-in project (not the Web project).</span></span>
    
2. <span data-ttu-id="29763-171">メニュー バーで、[ **表示**]、[ **プロパティ ウィンドウ**] の順に選択します。</span><span class="sxs-lookup"><span data-stu-id="29763-171">On the menu bar, choose  **View**,  **Properties Window**.</span></span>
    
<span data-ttu-id="29763-172">次の表に、プロジェクトのプロパティを示します。</span><span class="sxs-lookup"><span data-stu-id="29763-172">The following table describes the properties of the project.</span></span>



|<span data-ttu-id="29763-173">**プロパティ**</span><span class="sxs-lookup"><span data-stu-id="29763-173">**Property**</span></span>|<span data-ttu-id="29763-174">**説明**</span><span class="sxs-lookup"><span data-stu-id="29763-174">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="29763-175">**開始動作**</span><span class="sxs-lookup"><span data-stu-id="29763-175">**Start Action**</span></span>|<span data-ttu-id="29763-176">Office デスクトップ クライアントまたは指定のブラウザー内の Office Online クライアントのどちらでアドインをデバッグするか指定します。</span><span class="sxs-lookup"><span data-stu-id="29763-176">Specifies whether to debug your add-in in an Office desktop client or in an Office Online client in the specified browser.</span></span>|
|<span data-ttu-id="29763-177">**開始ドキュメント** (コンテンツ アドインと作業ウィンドウ アドインのみ)</span><span class="sxs-lookup"><span data-stu-id="29763-177">**Start Document** (Content and task pane add-ins only)</span></span>|<span data-ttu-id="29763-178">プロジェクトの開始時に開くドキュメントを指定します。</span><span class="sxs-lookup"><span data-stu-id="29763-178">Specifies what document to open when you start the project.</span></span>|
|<span data-ttu-id="29763-179">**Web プロジェクト**</span><span class="sxs-lookup"><span data-stu-id="29763-179">**Web Project**</span></span>|<span data-ttu-id="29763-180">アドインに関連付けられている Web プロジェクトの名前を指定します。</span><span class="sxs-lookup"><span data-stu-id="29763-180">Specifies the name of the web project associated with the add-in.</span></span>|
|<span data-ttu-id="29763-181">**電子メール アドレス** (Outlook アドインのみ)</span><span class="sxs-lookup"><span data-stu-id="29763-181">**Email Address** (Outlook add-ins only)</span></span>|<span data-ttu-id="29763-182">Outlook アドインのテストに使用する Exchange Server か Exchange Online のユーザー アカウントの電子メール アドレスを指定します。</span><span class="sxs-lookup"><span data-stu-id="29763-182">Specifies the email address of the user account in Exchange Server or Exchange Online that you want to test your Outlook add-in with.</span></span>|
|<span data-ttu-id="29763-183">**EWS の URL** (Outlook アドインのみ)</span><span class="sxs-lookup"><span data-stu-id="29763-183">**EWS Url** (Outlook add-ins only)</span></span>|<span data-ttu-id="29763-184">Exchange Web サービスの URL (例: https://www.contoso.com/ews/exchange.aspx)。</span><span class="sxs-lookup"><span data-stu-id="29763-184">Exchange Web service URL (For example: https://www.contoso.com/ews/exchange.aspx).</span></span> |
|<span data-ttu-id="29763-185">**OWA の URL** (Outlook アドインのみ)</span><span class="sxs-lookup"><span data-stu-id="29763-185">**OWA Url** (Outlook add-ins only)</span></span>|<span data-ttu-id="29763-186">Outlook Web App の URL (例: https://www.contoso.com/owa)。</span><span class="sxs-lookup"><span data-stu-id="29763-186">Outlook Web App URL (For example: https://www.contoso.com/owa).</span></span>|
|<span data-ttu-id="29763-187">**ユーザー名** (Outlook アドインのみ)</span><span class="sxs-lookup"><span data-stu-id="29763-187">**User name** (Outlook add-ins only)</span></span>|<span data-ttu-id="29763-188">Exchange Server または Exchange Online のユーザー アカウントの名前を指定します。</span><span class="sxs-lookup"><span data-stu-id="29763-188">Specifies the name of your user account in Exchange Server or Exchange Online.</span></span>|
|<span data-ttu-id="29763-189">**プロジェクト ファイル**</span><span class="sxs-lookup"><span data-stu-id="29763-189">**Project File**</span></span>|<span data-ttu-id="29763-190">ビルド、構成、およびその他のプロジェクト情報が含まれているファイルの名前を指定します。</span><span class="sxs-lookup"><span data-stu-id="29763-190">Specifies the name of the file containing build, configuration, and other information about the project.</span></span>|
|<span data-ttu-id="29763-191">**プロジェクト フォルダー**</span><span class="sxs-lookup"><span data-stu-id="29763-191">**Project Folder**</span></span>|<span data-ttu-id="29763-192">プロジェクト ファイルの場所です。</span><span class="sxs-lookup"><span data-stu-id="29763-192">The location of the project file.</span></span>|

### <a name="use-an-existing-document-to-debug-the-add-in-content-and-task-pane-add-ins-only"></a><span data-ttu-id="29763-193">既存のドキュメントを使用してアドインをデバッグする (コンテンツ アドインと作業ウィンドウ アドインのみ)</span><span class="sxs-lookup"><span data-stu-id="29763-193">Use an existing document to debug the add-in (content and task pane add-ins only)</span></span>


<span data-ttu-id="29763-p116">アドイン プロジェクトにドキュメントを追加できます。アドインで使用するテスト データを含むドキュメントがある場合、プロジェクトの開始時に Visual Studio によってそのドキュメントが開かれます。</span><span class="sxs-lookup"><span data-stu-id="29763-p116">You can add documents to the add-in project. If you have a document that contains test data that you want to use with your add-in, Visual Studio opens that document for you when you start the project.</span></span>


### <a name="to-use-an-existing-document-to-debug-the-add-in"></a><span data-ttu-id="29763-196">既存のドキュメントを使用してアドインをデバッグするには</span><span class="sxs-lookup"><span data-stu-id="29763-196">To use an existing document to debug the add-in</span></span>


1. <span data-ttu-id="29763-197">**ソリューション エクスプローラー**で、アドイン プロジェクト フォルダーを選択します。</span><span class="sxs-lookup"><span data-stu-id="29763-197">In  **Solution Explorer**, choose the add-in project folder.</span></span>
    
    > [!NOTE]
    > <span data-ttu-id="29763-198">Web アプリケーション プロジェクトではなく、アドイン プロジェクトを選択します。</span><span class="sxs-lookup"><span data-stu-id="29763-198">Choose the add-in project and not the web application project.</span></span>

2. <span data-ttu-id="29763-199">**[プロジェクト]** メニューで、**[既存の項目の追加]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="29763-199">On the  **Project** menu, choose **Add Existing Item**.</span></span>
    
3. <span data-ttu-id="29763-200">[ **既存の項目の追加**] ダイアログ ボックスで、追加するドキュメントを探して選択します。</span><span class="sxs-lookup"><span data-stu-id="29763-200">In the  **Add Existing Item** dialog box, locate and select the document that you want to add.</span></span>
    
4. <span data-ttu-id="29763-201">[ **追加**] を選択して、ドキュメントをプロジェクトに追加します。</span><span class="sxs-lookup"><span data-stu-id="29763-201">Choose the  **Add** button to add the document to your project.</span></span>
    
5. <span data-ttu-id="29763-202">**ソリューション エクスプローラー**で、プロジェクトのショートカット メニューを開き、[ **プロパティ**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="29763-202">In  **Solution Explorer**, open the shortcut menu for the project, and then choose  **Properties**.</span></span>
    
    <span data-ttu-id="29763-203">プロジェクトのプロパティ ページが表示されます。</span><span class="sxs-lookup"><span data-stu-id="29763-203">The property pages for the project appear.</span></span>
    
6. <span data-ttu-id="29763-204">[ **開始ドキュメント**] 一覧で、プロジェクトに追加したドキュメントを選択し、[ **OK**] を選択してプロパティ ページを閉じます。</span><span class="sxs-lookup"><span data-stu-id="29763-204">In the  **Start Document** list, choose the document that you added to the project, and then choose the **OK** button to close the property pages.</span></span>
    

### <a name="start-the-solution"></a><span data-ttu-id="29763-205">ソリューションの起動</span><span class="sxs-lookup"><span data-stu-id="29763-205">Start the solution</span></span>


<span data-ttu-id="29763-p117">Visual Studio は、起動すると、自動的にソリューションをビルドします。ソリューションを起動するには、 **メニュー** バーから **[デバッグ]**、 **[開始]** の順に選択します。</span><span class="sxs-lookup"><span data-stu-id="29763-p117">Visual Studio will automatically build the solution when you start it. You can start the solution from the  **Menu** bar by choosing **Debug**,  **Start**.</span></span> 


> [!NOTE]
> <span data-ttu-id="29763-p118">Internet Explorer でスクリプトのデバッグが有効になっていない場合は、Visual Studio でデバッガーを起動することはできません。スクリプトのデバッグを有効にするには、**[インターネット オプション]** ダイアログ ボックスを開いて、**[詳細設定]** タブをクリックし、**[スクリプトのデバッグを使用しない (Internet Explorer)]** チェック ボックスおよび **[スクリプトのデバッグを使用しない (その他)]** チェック ボックスをオフにします。</span><span class="sxs-lookup"><span data-stu-id="29763-p118">If script debugging isn't enabled in Internet Explorer, you won't be able to start the debugger in Visual Studio. You can enable script debugging by opening the  **Internet Options** dialog box, choosing the **Advanced** tab, and then clearing the **Disable Script Debugging (Internet Explorer)** and **Disable Script Debugging (Other)** check boxes.</span></span>

<span data-ttu-id="29763-210">Visual Studio はプロジェクトをビルドし、次の操作を実行します。</span><span class="sxs-lookup"><span data-stu-id="29763-210">Visual Studio builds the project and does the following:</span></span>


1. <span data-ttu-id="29763-p119">XML マニフェスト ファイルのコピーを作成し、それを  _プロジェクト名_\Output ディレクトリに追加します。このコピーは、Visual Studio を起動してアドインをデバッグするときにホスト アプリケーションで使用されます。</span><span class="sxs-lookup"><span data-stu-id="29763-p119">Creates a copy of the XML manifest file and adds it to  _ProjectName_\Output directory. The host application consumes this copy when you start Visual Studio and debug the add-in.</span></span>
    
2. <span data-ttu-id="29763-213">アドインをホストアプリケーションに表示するための一連のレジストリエントリをコンピューターに作成します。</span><span class="sxs-lookup"><span data-stu-id="29763-213">Creates a set of registry entries on your computer that enable the add-in to appear in the host application.</span></span>
    
3. <span data-ttu-id="29763-214">Web アプリケーション プロジェクトをビルドし、ローカルの IIS Web サーバー (http://localhost)) に展開します。</span><span class="sxs-lookup"><span data-stu-id="29763-214">Builds the web application project, and then deploys it to the local IIS web server (http://localhost).</span></span> 
    
<span data-ttu-id="29763-215">次に、Visual Studio は次の操作を実行します。</span><span class="sxs-lookup"><span data-stu-id="29763-215">Next, Visual Studio does the following:</span></span>


1. <span data-ttu-id="29763-216">~remoteAppUrlトークンを開始ページの完全修飾アドレス (例: http://localhost/MyAgave.html)) で置き換えることによって、XML マニフェストファイルの  [SourceLocation](https://docs.microsoft.com/javascript/office/manifest/sourcelocation?view=office-js)  要素を変更します。</span><span class="sxs-lookup"><span data-stu-id="29763-216">Modifies the SourceLocation element of the XML manifest file by replacing the ~remoteAppUrl token with the fully qualified address of the start page (for example, http://localhost/MyAgave.html).</span></span>
    
2. <span data-ttu-id="29763-217">IIS Express で Web アプリケーション プロジェクトを起動します。</span><span class="sxs-lookup"><span data-stu-id="29763-217">Starts the web application project in IIS Express.</span></span>
    
3. <span data-ttu-id="29763-218">ホスト アプリケーションを開きます。</span><span class="sxs-lookup"><span data-stu-id="29763-218">Opens the host application.</span></span> 
    
<span data-ttu-id="29763-p120">プロジェクトをビルドする際、Visual Studio は **出力**ウィンドウに検証エラーを表示しません。Visual Studio は、エラーと警告を、発生時に  **ERRORLIST** ウィンドウ内で報告します。Visual Studio は、コードおよびテキスト エディター内で検証エラーを別の色の波形の下線 (波線と呼びます) で示します。このようなマークにより、Visual Studio がコード内で検出した問題が通知されます。詳細については、「 [コードおよびテキスト エディター](https://msdn.microsoft.com/library/se2f663y(v=vs.140).aspx)」を参照してください。検証を有効化または無効化する方法の詳細については、次のトピックを参照してください。</span><span class="sxs-lookup"><span data-stu-id="29763-p120">Visual Studio doesn't show validation errors in the  **OUTPUT** window when you build the project. Visual Studio reports errors and warnings in the **ERRORLIST** window as they occur. Visual Studio also reports validation errors by showing wavy underlines (known as squiggles) of different colors in the code and text editor. These marks notify you of problems that Visual Studio detected in your code. For more information, see [Code and Text Editor](https://msdn.microsoft.com/library/se2f663y(v=vs.140).aspx). For more information about how to enable or disable validation, see:</span></span> 

- <span data-ttu-id="29763-225">[[オプション]、[テキスト エディター]、[JavaScript]、[IntelliSense]](https://docs.microsoft.com/visualstudio/ide/reference/options-text-editor-javascript-intellisense?view=vs-2015)</span><span class="sxs-lookup"><span data-stu-id="29763-225">[Options, Text Editor, JavaScript, IntelliSense](https://docs.microsoft.com/visualstudio/ide/reference/options-text-editor-javascript-intellisense?view=vs-2015)</span></span>
    
- <span data-ttu-id="29763-226">[方法:Visual Web Developer で HTML 編集用の検証オプションを設定する](https://msdn.microsoft.com/library/0byxkfet(v=vs.100).aspx)</span><span class="sxs-lookup"><span data-stu-id="29763-226">[How to: Set Validation Options for HTML Editing in Visual Web Developer](https://msdn.microsoft.com/library/0byxkfet(v=vs.100).aspx)</span></span>
    
- <span data-ttu-id="29763-227">[[検証] ([オプション] ダイアログ ボックス - [テキスト エディター] - [CSS])](https://msdn.microsoft.com/library/se2f663y(v=vs.140).aspx)</span><span class="sxs-lookup"><span data-stu-id="29763-227">[CSS, see Validation, CSS, Text Editor, Options Dialog Box](https://msdn.microsoft.com/library/se2f663y(v=vs.140).aspx)</span></span>
    
<span data-ttu-id="29763-228">プロジェクト内の XML マニフェスト ファイルの検証ルールを確認するには、「[Office アドインの XML マニフェスト](../develop/add-in-manifests.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="29763-228">To review the validation rules of the XML manifest file in your project, see [Office Add-ins XML manifest](../develop/add-in-manifests.md).</span></span>


### <a name="show-an-add-in-in-excel-word-or-project-and-step-through-your-code"></a><span data-ttu-id="29763-229">Excel、Word、または Project にアドインを表示し、コードをステップ実行する</span><span class="sxs-lookup"><span data-stu-id="29763-229">Show an add-in in Excel, Word, or Project and step through your code</span></span>


<span data-ttu-id="29763-p121">アドイン プロジェクトの **開始ドキュメント** プロパティを Excel または Word に設定した場合、Visual Studio はドキュメントを新規作成し、アドインが表示されます。アドイン プロジェクトの **開始ドキュメント** プロパティを既存のドキュメントを使用するように設定した場合、Visual Studio はドキュメントを開きますが、アドインは手動で挿入する必要があります。 **開始ドキュメント** を **Microsoft Project** に設定した場合も、アドインを手動で挿入する必要があります。</span><span class="sxs-lookup"><span data-stu-id="29763-p121">If you set the  **Start Document** property of the add-in project to Excel or Word, Visual Studio creates a new document and the add-in appears. If you set the **Start Document** property of the add-in project to use an existing document, Visual Studio opens the document, but you have to insert the add-in manually. If you set the **Start Document** to **Microsoft Project**, you also have to insert the add-in manually.</span></span>


### <a name="to-show-an-office-add-in-in-excel-or-word"></a><span data-ttu-id="29763-233">Office アドイン を Excel または Word で表示するには</span><span class="sxs-lookup"><span data-stu-id="29763-233">To show an Office Add-in in Excel or Word</span></span>


1. <span data-ttu-id="29763-234">Excel または Word の  **[挿入]** タブで、 **[Office アドイン]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="29763-234">In Excel or Word, on the  **Insert** tab, choose **Office Add-ins**.</span></span>
    
2. <span data-ttu-id="29763-235">表示される一覧で、アドインを選択します。</span><span class="sxs-lookup"><span data-stu-id="29763-235">In the list that appears, choose your add-in.</span></span>
    

### <a name="to-show-an-office-add-in-in-project"></a><span data-ttu-id="29763-236">Project で Office アドインを表示するには</span><span class="sxs-lookup"><span data-stu-id="29763-236">To show an Office Add-in in Project</span></span>


1. <span data-ttu-id="29763-237">Project の  **[プロジェクト]** タブで、 **[Office アドイン]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="29763-237">In Project, on the  **Project** tab, choose **Office Add-ins**.</span></span>
    
2. <span data-ttu-id="29763-238">表示される一覧で、アドインを選択します。</span><span class="sxs-lookup"><span data-stu-id="29763-238">In the list that appears, choose your add-in.</span></span>
    
<span data-ttu-id="29763-p122">Visual Studio でブレークポイントを設定できます。ブレークポイントを設定したら、アドインを操作し、HTML、JavaScript、および C# か VB のコード ファイル内のコードをステップ実行します。</span><span class="sxs-lookup"><span data-stu-id="29763-p122">In Visual Studio, you can then set break-points. Then, as you interact with your add-in and step through the code in your HTML, JavaScript, and C# or VB code files.</span></span>


### <a name="show-the-outlook-add-in-in-outlook-and-step-through-your-code"></a><span data-ttu-id="29763-241">Outlook で Outlook のアドインを表示し、コードをステップ実行する</span><span class="sxs-lookup"><span data-stu-id="29763-241">Show the Outlook add-in in Outlook and step through your code</span></span>


<span data-ttu-id="29763-242">Outlook でアドインを表示するには、電子メール メッセージまたは予定アイテムを開きます。</span><span class="sxs-lookup"><span data-stu-id="29763-242">To view the add-in in Outlook, open an email message or appointment item.</span></span>

<span data-ttu-id="29763-p123">Outlook は、アクティブ化の基準を満たしていれば、アイテムの アドイン をアクティブ化します。アドイン バーが [インスペクタ] ウィンドウまたは閲覧ウィンドウの上部に表示され、Outlook アドインがアドイン バーにボタンとして表示されます。アドインにアドイン コマンドがある場合は、リボンの既定のタブまたは指定されたカスタム タブのいずれかにボタンが表示され、アドイン バーにはアドインは表示されません。</span><span class="sxs-lookup"><span data-stu-id="29763-p123">Outlook activates the add-in for the item as long as the activation criteria are met. The add-in bar appears at the top of the Inspector window or Reading Pane, and your Outlook add-in appears as a button in the add-in bar. If your add-in has an add-in command, a button will appear in the ribbon, either in the default tab or a specified custom tab, and the add-in will not appear in the add-in bar.</span></span>

<span data-ttu-id="29763-246">Outlook アドインを表示するには、Outlook アドインのボタンを選択します。</span><span class="sxs-lookup"><span data-stu-id="29763-246">To view your Outlook add-in, choose the button for your Outlook add-in.</span></span>

<span data-ttu-id="29763-p124">Visual Studio では、ブレークポイントを設定できます。ブレークポイントを設定した後、Outlook アドインを操作して、HTML、JavaScript、および C# または VB のコード ファイルのコードをステップ実行します。</span><span class="sxs-lookup"><span data-stu-id="29763-p124">In Visual Studio, you can set break-points. Then, as you interact with your Outlook add-in and step through the code in your HTML, JavaScript, and C# or VB code files.</span></span> 

<span data-ttu-id="29763-p125">また、コードを変更してから、Office アドイン を終了してプロジェクトを再度起動しなくても、Outlook アドインへの影響を確認することができます。Outlook で Outlook アドインのショートカット メニューを開き、 **[再読み込み]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="29763-p125">You can also change your code and review the effects of those changes in your Outlook add-in without having to close the Office Add-in and start the project again. In Outlook, just open the shortcut menu for the Outlook add-in, and then choose  **Reload**.</span></span>


### <a name="modify-code-and-continue-to-debug-the-add-in-without-having-to-start-the-project-again"></a><span data-ttu-id="29763-251">コードを変更した後、プロジェクトを再び開始することなくアドインのデバッグを続行する</span><span class="sxs-lookup"><span data-stu-id="29763-251">Modify code and continue to debug the add-in without having to start the project again</span></span>


<span data-ttu-id="29763-p126">コードを変更したら、ホスト アプリケーションを閉じてプロジェクトを再起動しなくても、アドインに対する変更の影響を確認することができます。コードを変更した後、アドインのショートカット メニューを開き、 **[再読み込み]** を選択します。アドインを再読み込みすると、アドインは Visual Studio デバッガーと切断された状態になります。そのため、変更の影響を確認することはできても、利用できるすべての Iexplore.exe プロセスに Visual Studio デバッガーをアタッチするまでは、コードをステップ実行していくことはできません。</span><span class="sxs-lookup"><span data-stu-id="29763-p126">You can change your code and review the effects of those changes in your add-in without having to close the host application and start the project again. After you change your code, open the shortcut menu for the add-in, and then choose  **Reload**. When you reload the add-in it becomes disconnected with the Visual Studio debugger. Therefore, you can view the effects of your change, but you cannot step through your code again until you attach the Visual Studio debugger to all of the available Iexplore.exe processes.</span></span>


### <a name="to-attach-the-visual-studio-debugger-to-all-of-the-available-iexploreexe-processes"></a><span data-ttu-id="29763-256">使用可能な Iexplore.exe プロセスのすべてに Visual Studio デバッガーをアタッチするには</span><span class="sxs-lookup"><span data-stu-id="29763-256">To attach the Visual Studio debugger to all of the available Iexplore.exe processes</span></span>


1. <span data-ttu-id="29763-257">Visual Studio で、 **[デバッグ]**、 **[プロセスにアタッチ]** の順に選択します。</span><span class="sxs-lookup"><span data-stu-id="29763-257">In Visual Studio, choose  **DEBUG**,  **Attach to Process**.</span></span>
    
2. <span data-ttu-id="29763-258">[ **プロセスにアタッチ**] ダイアログ ボックスで、利用可能なすべての  **Iexplore.exe** プロセスを選択して、 [ **アタッチ**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="29763-258">In the  **Attach to Process** dialog box, choose all of the available **Iexplore.exe** processes, and then choose the **Attach** button.</span></span>
    

## <a name="next-steps"></a><span data-ttu-id="29763-259">次の手順</span><span class="sxs-lookup"><span data-stu-id="29763-259">Next steps</span></span>

- [<span data-ttu-id="29763-260">Office アドインを展開し、発行する</span><span class="sxs-lookup"><span data-stu-id="29763-260">Deploy and publish your Office Add-in</span></span>](../publish/publish.md)
    
