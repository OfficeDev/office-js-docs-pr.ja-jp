---
title: 社内の Project Server OData サービスで REST を使用する Project アドインを作成する
description: Project Professional 2013 の作業ウィンドウ アドインを作成し、アクティブなプロジェクトのコストと作業データと、現在の Project Web App インスタンスのすべてのプロジェクトの平均を比較する方法について学習します。
ms.date: 09/26/2019
localization_priority: Normal
ms.openlocfilehash: c03cd580f9f5d4da654022de811d4a060a99e52d
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/09/2021
ms.locfileid: "53348812"
---
# <a name="create-a-project-add-in-that-uses-rest-with-an-on-premises-project-server-odata-service"></a><span data-ttu-id="be649-103">社内の Project Server OData サービスで REST を使用する Project アドインを作成する</span><span class="sxs-lookup"><span data-stu-id="be649-103">Create a Project add-in that uses REST with an on-premises Project Server OData service</span></span>

<span data-ttu-id="be649-104">この記事では、アクティブなプロジェクトのコストと作業のデータを現在の Project Web App インスタンスの全プロジェクトの平均と比較する Project Professional 2013 用の作業ウィンドウ アドインを作成する方法を説明します。</span><span class="sxs-lookup"><span data-stu-id="be649-104">This article describes how to build a task pane add-in for Project Professional 2013 that compares cost and work data in the active project with the averages for all projects in the current Project Web App instance.</span></span> <span data-ttu-id="be649-105">アドインは jQuery ライブラリと一緒に REST を使用して、サーバー 2013 の **ProjectData** OData Projectアクセスします。</span><span class="sxs-lookup"><span data-stu-id="be649-105">The add-in uses REST with the jQuery library to access the **ProjectData** OData reporting service in Project Server 2013.</span></span>

<span data-ttu-id="be649-106">この記事のコードは、Microsoft Corporation の Saurabh Sanghvi と Arvind Iyer が開発したサンプルに基づいています。</span><span class="sxs-lookup"><span data-stu-id="be649-106">The code in this article is based on a sample developed by Saurabh Sanghvi and Arvind Iyer, Microsoft Corporation.</span></span>

## <a name="prerequisites-for-creating-a-task-pane-add-in-that-reads-project-server-reporting-data"></a><span data-ttu-id="be649-107">Project Server のレポート データを読み取る作業ウィンドウ アドインを作成するための前提条件</span><span class="sxs-lookup"><span data-stu-id="be649-107">Prerequisites for creating a task pane add-in that reads Project Server reporting data</span></span>

<span data-ttu-id="be649-108">Project Server 2013 のオンプレミス インストールで Project Web App インスタンスの **ProjectData** サービスを読み取る Project Project 作業ウィンドウ アドインを作成するための前提条件を次に示します。</span><span class="sxs-lookup"><span data-stu-id="be649-108">The following are the prerequisites for creating a Project task pane add-in that reads the **ProjectData** service of a Project Web App instance in an on-premises installation of Project Server 2013.</span></span>

- <span data-ttu-id="be649-p102">使用するローカルの開発用コンピューターに最新のサービス パックと Windows 更新プログラムをインストールしてあることを確認します。オペレーティング システムは、Windows 7、Windows 8、Windows Server 2008、Windows Server 2012 のいずれでもかまいません。</span><span class="sxs-lookup"><span data-stu-id="be649-p102">Ensure that you have installed the most recent service packs and Windows updates on your local development computer. The operating system can be Windows 7, Windows 8, Windows Server 2008, or Windows Server 2012.</span></span>

- <span data-ttu-id="be649-111">Project Web App との接続には Project Professional 2013が必要です。</span><span class="sxs-lookup"><span data-stu-id="be649-111">Project Professional 2013 is required to connect with Project Web App.</span></span> <span data-ttu-id="be649-112">F5 デバッグを有効にするには、Project Professional 2013 がインストールされている必要Visual Studio。</span><span class="sxs-lookup"><span data-stu-id="be649-112">The development computer must have Project Professional 2013 installed to enable **F5** debugging with Visual Studio.</span></span>

    > [!NOTE]
    > <span data-ttu-id="be649-113">Project Standard 2013 では作業ウィンドウ アドインをホストすることもできますが、作業ウィンドウ アドインにサインインProject Web App。</span><span class="sxs-lookup"><span data-stu-id="be649-113">Project Standard 2013 can also host task pane add-ins, but cannot sign in to Project Web App.</span></span>

- <span data-ttu-id="be649-114">Office Developer Tools for Visual Studio を備えた Visual Studio 2015 には、Office アドインと SharePoint アドインの作成用のテンプレートが含まれています。最新バージョンの Office Developer Tools がインストールされていることを確認してください。 _Office アドインと SharePoint のダウンロード_ の「 [ツール](https://developer.microsoft.com/office/docs) 」セクションを参照してください。</span><span class="sxs-lookup"><span data-stu-id="be649-114">Visual Studio 2015 with Office Developer Tools for Visual Studio includes templates for creating Office and SharePoint Add-ins. Ensure that you have installed the most recent version of Office Developer Tools; see the  _Tools_ section of the [Office Add-ins and SharePoint downloads](https://developer.microsoft.com/office/docs).</span></span>

- <span data-ttu-id="be649-115">この記事の手順とコード例は、ローカル ドメインProject Server 2013 の **ProjectData** サービスにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="be649-115">The procedures and code examples in this article access the **ProjectData** service of Project Server 2013 in a local domain.</span></span> <span data-ttu-id="be649-116">この記事の jQuery メソッドは、web 上Project機能しません。</span><span class="sxs-lookup"><span data-stu-id="be649-116">The jQuery methods in this article do not work with Project on the web.</span></span>

    <span data-ttu-id="be649-117">**ProjectData サービスに開発** コンピューターからアクセス可能な環境を確認します。</span><span class="sxs-lookup"><span data-stu-id="be649-117">Verify that the **ProjectData** service is accessible from your development computer.</span></span>

### <a name="procedure-1-to-verify-that-the-projectdata-service-is-accessible"></a><span data-ttu-id="be649-p105">手順 1. ProjectData サービスにアクセスできることを確認するには</span><span class="sxs-lookup"><span data-stu-id="be649-p105">Procedure 1. To verify that the ProjectData service is accessible</span></span>

1. <span data-ttu-id="be649-p106">ブラウザーで REST クエリからの XML データの直接表示を可能にするには、フィードの読み取りビューをオフにします。Internet Explorer でこれを行う方法については、「 [Project Server 2013 レポート データの OData フィードにクエリを実行する](/previous-versions/office/project-odata/jj163048(v=office.15))」の手順 1. のステップ 4. を参照してください。</span><span class="sxs-lookup"><span data-stu-id="be649-p106">To enable your browser to directly show the XML data from a REST query, turn off the feed reading view. For information about how to do this in Internet Explorer, see Procedure 1, step 4 in [Querying OData feeds for Project reporting data](/previous-versions/office/project-odata/jj163048(v=office.15)).</span></span>

2. <span data-ttu-id="be649-122">ブラウザーを **使用して、/ProjectServerName** **http://ServerName /_api/ProjectData を使用して ProjectData サービスにクエリを実行します**。</span><span class="sxs-lookup"><span data-stu-id="be649-122">Query the **ProjectData** service by using your browser with the following URL: **http://ServerName /ProjectServerName /_api/ProjectData**.</span></span> <span data-ttu-id="be649-123">たとえば、インスタンスのインスタンスが `http://MyServer/pwa` Project Web App、ブラウザーに次の結果が表示されます。</span><span class="sxs-lookup"><span data-stu-id="be649-123">For example, if the Project Web App instance is  `http://MyServer/pwa`, the browser shows the following results.</span></span>

    ```xml
    <?xml version="1.0" encoding="utf-8"?>
        <service xml:base="http://myserver/pwa/_api/ProjectData/"
        xmlns="https://www.w3.org/2007/app"
        xmlns:atom="https://www.w3.org/2005/Atom">
        <workspace>
            <atom:title>Default</atom:title>
            <collection href="Projects">
                <atom:title>Projects</atom:title>
            </collection>
            <collection href="ProjectBaselines">
                <atom:title>ProjectBaselines</atom:title>
            </collection>
            <!-- ... and 33 more collection elements -->
        </workspace>
        </service>
    ```

3. <span data-ttu-id="be649-p108">結果を確認するためにネットワーク資格情報の入力が必要になることもあります。ブラウザーでアクセスが拒否されことを示すメッセージ (エラー 403) が表示された場合は、その Project Web App インスタンスに対するログオンのアクセス許可を与えられていないか、管理者のサポートを必要とするネットワーク上の問題が発生しています。</span><span class="sxs-lookup"><span data-stu-id="be649-p108">You may have to provide your network credentials to see the results. If the browser shows "Error 403, Access Denied," either you do not have logon permission for that Project Web App instance, or there is a network problem that requires administrative help.</span></span>

## <a name="using-visual-studio-to-create-a-task-pane-add-in-for-project"></a><span data-ttu-id="be649-126">Visual Studio を使用して Project 用の作業ウィンドウ アドインを作成する</span><span class="sxs-lookup"><span data-stu-id="be649-126">Using Visual Studio to create a task pane add-in for Project</span></span>

<span data-ttu-id="be649-127">Office Developer Tools for Visual Studio には、Project 2013 用の作業ウィンドウ アドインのためのテンプレートが含まれています。</span><span class="sxs-lookup"><span data-stu-id="be649-127">Office Developer Tools for Visual Studio includes a template for task pane add-ins for Project 2013.</span></span> <span data-ttu-id="be649-128">**HelloProjectOData** という名前のソリューションを作成する場合、ソリューションには次の 2 つのプロジェクトVisual Studioされます。</span><span class="sxs-lookup"><span data-stu-id="be649-128">If you create a solution named **HelloProjectOData**, the solution contains the following two Visual Studio projects:</span></span>

- <span data-ttu-id="be649-129">アドイン プロジェクトは、ソリューションの名前を受け取る。</span><span class="sxs-lookup"><span data-stu-id="be649-129">The add-in project takes the name of the solution.</span></span> <span data-ttu-id="be649-130">アドインの XML マニフェスト ファイルが含まれています。4.5 の.NET Frameworkします。</span><span class="sxs-lookup"><span data-stu-id="be649-130">It includes the XML manifest file for the add-in and targets the .NET Framework 4.5.</span></span> <span data-ttu-id="be649-131">手順 3 は **、HelloProjectOData** アドインのマニフェストを変更する手順を示しています。</span><span class="sxs-lookup"><span data-stu-id="be649-131">Procedure 3 shows the steps to modify the manifest for the **HelloProjectOData** add-in.</span></span>

- <span data-ttu-id="be649-132">Web プロジェクトの名前は **HelloProjectODataWeb** です。</span><span class="sxs-lookup"><span data-stu-id="be649-132">The web project is named **HelloProjectODataWeb**.</span></span> <span data-ttu-id="be649-133">作業ウィンドウ内の Web コンテンツの Web ページ、JavaScript ファイル、CSS ファイル、イメージ、参照、および構成ファイルが含まれます。</span><span class="sxs-lookup"><span data-stu-id="be649-133">It includes the webpages, JavaScript files, CSS files, images, references, and configuration files for the web content in the task pane.</span></span> <span data-ttu-id="be649-134">Web プロジェクトのターゲットは、.NET Framework 4 です。</span><span class="sxs-lookup"><span data-stu-id="be649-134">The web project targets the .NET Framework 4.</span></span> <span data-ttu-id="be649-135">手順 4 と手順 5 では、Web プロジェクト内のファイルを変更して **HelloProjectOData** アドインの機能を作成する方法を示します。</span><span class="sxs-lookup"><span data-stu-id="be649-135">Procedure 4 and Procedure 5 show how to modify the files in the web project to create the functionality of the **HelloProjectOData** add-in.</span></span>

### <a name="procedure-2-to-create-the-helloprojectodata-add-in-for-project"></a><span data-ttu-id="be649-p112">手順 2. Project 用の HelloProjectOData アドインを作成するには</span><span class="sxs-lookup"><span data-stu-id="be649-p112">Procedure 2. To create the HelloProjectOData add-in for Project</span></span>

1. <span data-ttu-id="be649-138">管理者Visual Studio 2015 を実行し、[スタート] ページで [新しい **Project]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="be649-138">Run Visual Studio 2015 as an administrator, and then select **New Project** on the Start page.</span></span>

2. <span data-ttu-id="be649-139">[新 **しいProject]** ダイアログ ボックスで、[テンプレート] **、Visual C#** ノード **、Office/SharePoint** ノードを展開し、[Office アドイン] を **選択します**。中央.NET Framework上部のターゲット フレームワーク ドロップダウン リストで **[4.5.2]** を選択し **、[Office** アドイン] を選択します (次のスクリーンショットを参照)。</span><span class="sxs-lookup"><span data-stu-id="be649-139">In the **New Project** dialog box, expand the **Templates**, **Visual C#**, and **Office/SharePoint** nodes, and then select **Office Add-ins**. Select **.NET Framework 4.5.2** in the target framework drop-down list at the top of the center pane, and then select **Office Add-in** (see the next screenshot).</span></span>

3. <span data-ttu-id="be649-140">これらの Visual Studio プロジェクトを両方とも同じディレクトリに配置するには、[**ソリューションのディレクトリを作成**] を選択し、目的の場所を参照します。</span><span class="sxs-lookup"><span data-stu-id="be649-140">To place both of the Visual Studio projects in the same directory, select **Create directory for solution**, and then browse to the location you want.</span></span>

4. <span data-ttu-id="be649-141">[名前 **] フィールドに** 「HelloProjectOData」と入力し **、[OK] を選択します**。</span><span class="sxs-lookup"><span data-stu-id="be649-141">In the **Name** field, typeHelloProjectOData, and then choose **OK**.</span></span>

    <span data-ttu-id="be649-142">*図 1. Office アドインの作成*</span><span class="sxs-lookup"><span data-stu-id="be649-142">*Figure 1. Creating an Office Add-in*</span></span>

    ![アドインOffice作成します。](../images/pj15-hello-project-o-data-creating-app.png)

5. <span data-ttu-id="be649-144">[**アドインの種類の選択**] ダイアログ ボックスで、[**作業ウィンドウ**] を選択して、[**次へ**] を選択します (次のスクリーンショットを参照)。</span><span class="sxs-lookup"><span data-stu-id="be649-144">In the **Choose the add-in type** dialog box, select **Task pane** and choose **Next** (see the next screenshot).</span></span>

    <span data-ttu-id="be649-145">*図 2. 作成するアドインの種類の選択*</span><span class="sxs-lookup"><span data-stu-id="be649-145">*Figure 2. Choosing the type of add-in to create*</span></span>

    ![作成するアドインの種類を選択します。](../images/pj15-hello-project-o-data-choose-project.png)

6. <span data-ttu-id="be649-147">[**ホスト アプリケーションの選択**] ダイアログ ボックスで、[**Project**] 以外のすべてのチェック ボックスをオフにし (次のスクリーンショットを参照)、[**完了**] をクリックします。</span><span class="sxs-lookup"><span data-stu-id="be649-147">In the **Choose the host applications** dialog box, clear all check boxes except the **Project** check box (see the next screenshot) and choose **Finish**.</span></span>

    <span data-ttu-id="be649-148">*図 3. ホスト アプリケーションの選択*</span><span class="sxs-lookup"><span data-stu-id="be649-148">*Figure 3. Choosing the host application*</span></span>

    ![ホスト アプリケーションProjectを選択します。](../images/create-office-add-in.png)

    <span data-ttu-id="be649-150">Visual Studio **HelloProjectOdata** プロジェクトと **HelloProjectODataWeb プロジェクトを作成** します。</span><span class="sxs-lookup"><span data-stu-id="be649-150">Visual Studio creates the **HelloProjectOdata** project and the **HelloProjectODataWeb** project.</span></span>

<span data-ttu-id="be649-151">[**AddIn**] フォルダー (次のスクリーンショットを参照) には、カスタム CSS スタイルの App.css ファイルが含まれています。</span><span class="sxs-lookup"><span data-stu-id="be649-151">The **AddIn** folder (see the next screenshot) contains the App.css file for custom CSS styles.</span></span> <span data-ttu-id="be649-152">**ホーム** サブフォルダー内の、Home.html ファイルには、CSS ファイル、アドインを使用している JavaScript ファイル、およびアドインの HTML5 コンテンツへの参照が含まれています。</span><span class="sxs-lookup"><span data-stu-id="be649-152">In the **Home** subfolder , the Home.html file contains references to the CSS files and the JavaScript files that the add-in uses, and the HTML5 content for the add-in.</span></span> <span data-ttu-id="be649-153">また、Home.js ファイルは、カスタムの JavaScript コード用です。</span><span class="sxs-lookup"><span data-stu-id="be649-153">Also, the Home.js file is for your custom JavaScript code.</span></span> <span data-ttu-id="be649-154">**Scripts** フォルダーには、jQuery ライブラリのファイルが含まれています。</span><span class="sxs-lookup"><span data-stu-id="be649-154">The **Scripts** folder includes the jQuery library files.</span></span> <span data-ttu-id="be649-155">**Office** サブフォルダーには、office.js や project-15.js などの JavaScript ライブラリ、および Office アドインでの標準の文字列用の言語ライブラリが含まれています。**コンテンツ** フォルダーで Office.css ファイルには、Office のアドインのすべてに使用する既定のスタイルが含まれます。</span><span class="sxs-lookup"><span data-stu-id="be649-155">The **Office** subfolder includes the JavaScript libraries such as office.js and project-15.js, plus the language libraries for standard strings in the Office Add-ins. In the **Content** folder, the Office.css file contains the default styles for all of the Office Add-ins.</span></span>

<span data-ttu-id="be649-156">*図 4. ソリューション エクスプローラーでの既定の Web プロジェクト ファイルの表示*</span><span class="sxs-lookup"><span data-stu-id="be649-156">*Figure 4. Viewing the default web project files in Solution Explorer*</span></span>

![ソリューション Expl で Web プロジェクト ファイルを表示する。](../images/pj15-hello-project-o-data-initial-solution-explorer.png)

<span data-ttu-id="be649-158">**HelloProjectOData プロジェクトの** マニフェストは、HelloProjectOData.xmlファイルです。</span><span class="sxs-lookup"><span data-stu-id="be649-158">The manifest for the **HelloProjectOData** project is the HelloProjectOData.xml file.</span></span> <span data-ttu-id="be649-159">必要に応じてマニフェストを編集して、アドインの説明、アイコンへの参照、追加言語の情報、その他の設定を追加できます。</span><span class="sxs-lookup"><span data-stu-id="be649-159">You can optionally modify the manifest to add a description of the add-in, a reference to an icon, information for additional languages, and other settings.</span></span> <span data-ttu-id="be649-160">手順 3 では、アドインの表示名と説明を変更し、アイコンを追加します。</span><span class="sxs-lookup"><span data-stu-id="be649-160">Procedure 3 simply modifies the add-in display name and description, and adds an icon.</span></span>

<span data-ttu-id="be649-161">マニフェストについて詳しくは、「[Office アドインの XML マニフェスト](../develop/add-in-manifests.md)」と「[Office アドインのマニフェスト向けのスキーマ リファレンス (v1.1)](../develop/add-in-manifests.md#see-also)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="be649-161">For more information about the manifest, see [Office Add-ins XML manifest](../develop/add-in-manifests.md) and [Schema reference for Office Add-ins manifests (v1.1)](../develop/add-in-manifests.md#see-also).</span></span>

### <a name="procedure-3-to-modify-the-add-in-manifest"></a><span data-ttu-id="be649-p115">手順 3. アドインのマニフェストを変更するには</span><span class="sxs-lookup"><span data-stu-id="be649-p115">Procedure 3. To modify the add-in manifest</span></span>

1. <span data-ttu-id="be649-164">Visual Studio で、HelloProjectOData.xml ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="be649-164">In Visual Studio, open the HelloProjectOData.xml file.</span></span>

2. <span data-ttu-id="be649-165">既定の表示名は、Visual Studio プロジェクトの名前です ("HelloProjectOData")。</span><span class="sxs-lookup"><span data-stu-id="be649-165">The default display name is the name of the Visual Studio project ("HelloProjectOData").</span></span> <span data-ttu-id="be649-166">たとえば **、DisplayName** 要素の既定値を "Hello ProjectData" に変更します。</span><span class="sxs-lookup"><span data-stu-id="be649-166">For example, change the default value of the **DisplayName** element to"Hello ProjectData".</span></span>

3. <span data-ttu-id="be649-p117">既定の説明も "HelloProjectOData" です。たとえば、Description 要素の既定値を "Test REST queries of the ProjectData service" に変更します。</span><span class="sxs-lookup"><span data-stu-id="be649-p117">The default description is also "HelloProjectOData". For example, change the default value of the Description element to "Test REST queries of the ProjectData service".</span></span>

4. <span data-ttu-id="be649-169">リボンの [**プロジェクト**] タブの [**Office アドイン**] ドロップダウン リストに表示するアイコンを追加します。</span><span class="sxs-lookup"><span data-stu-id="be649-169">Add an icon to show in the **Office Add-ins** drop-down list on the **PROJECT** tab of the ribbon.</span></span> <span data-ttu-id="be649-170">Visual Studio ソリューションにアイコン ファイルを追加することも、アイコンの URL を使用することもできます。</span><span class="sxs-lookup"><span data-stu-id="be649-170">You can add an icon file in the Visual Studio solution or use a URL for an icon.</span></span> 

<span data-ttu-id="be649-171">次の手順は、アイコン ファイルをソリューションに追加するVisual Studioします。</span><span class="sxs-lookup"><span data-stu-id="be649-171">The following steps show how to add an icon file to the Visual Studio solution.</span></span>

1. <span data-ttu-id="be649-172">ソリューション **エクスプローラーで、Images** という名前のフォルダーに移動します。</span><span class="sxs-lookup"><span data-stu-id="be649-172">In **Solution Explorer**, go to the folder named Images.</span></span>

2. <span data-ttu-id="be649-173">[アドイン] ドロップダウン リスト **Office** 表示するには、アイコンが 32 x 32 ピクセルである必要があります。</span><span class="sxs-lookup"><span data-stu-id="be649-173">To be displayed in the **Office Add-ins** drop-down list, the icon must be 32 x 32 pixels.</span></span> <span data-ttu-id="be649-174">たとえば、Project 2013 SDK をインストールしてから、**[Images]** フォルダーを選択し、SDK から次のファイルを追加します。`\Samples\Apps\HelloProjectOData\HelloProjectODataWeb\Images\NewIcon.png`。</span><span class="sxs-lookup"><span data-stu-id="be649-174">For example, install the Project 2013 SDK, and then choose the **Images** folder and add the following file from the SDK: `\Samples\Apps\HelloProjectOData\HelloProjectODataWeb\Images\NewIcon.png`</span></span>

    <span data-ttu-id="be649-175">または、独自の 32 x 32 アイコンを使用します。または、次のイメージを [NewIcon.png] という名前のファイルにコピーし、そのファイルをフォルダーに追加  `HelloProjectODataWeb\Images` します。</span><span class="sxs-lookup"><span data-stu-id="be649-175">Alternately, use your own 32 x 32 icon; or, copy the following image to a file named NewIcon.png, and then add that file to the  `HelloProjectODataWeb\Images` folder.</span></span>

    ![HelloProjectOData アプリのアイコン。](../images/pj15-hello-project-data-new-icon.jpg)

3. <span data-ttu-id="be649-177">このマニフェスト **HelloProjectOData.xml、Description** 要素の下に **IconUrl** 要素を追加します。ここで、アイコン URL の値は、32x32 アイコン ファイルへの相対パスです。</span><span class="sxs-lookup"><span data-stu-id="be649-177">In the HelloProjectOData.xml manifest, add an **IconUrl** element below the **Description** element, where the value of the icon URL is the relative path to the 32x32 icon file.</span></span> <span data-ttu-id="be649-178">たとえば、次の行を追加します **<IconUrl DefaultValue="~remoteAppUrl/Images/NewIcon.png" />** 。</span><span class="sxs-lookup"><span data-stu-id="be649-178">For example, add the following line: **<IconUrl DefaultValue="~remoteAppUrl/Images/NewIcon.png" />**.</span></span> <span data-ttu-id="be649-179">これでHelloProjectOData.xmlマニフェスト ファイルに次の情報が含まれます **(Id** の値は異なります)。</span><span class="sxs-lookup"><span data-stu-id="be649-179">The HelloProjectOData.xml manifest file now contains the following (your **Id** value will be different):</span></span>

    ```XML
    <?xml version="1.0" encoding="UTF-8"?>
    <OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
            xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="TaskPaneApp">
        <!--IMPORTANT! Id must be unique for each add-in. If you copy this manifest ensure that you change this id to your own GUID. -->
        <Id>c512df8d-a1c5-4d74-8a34-d30f6bbcbd82</Id>
        <Version>1.0</Version>
        <ProviderName> [Provider name]</ProviderName>
        <DefaultLocale>en-US</DefaultLocale>
        <DisplayName DefaultValue="Hello ProjectData" />
        <Description DefaultValue="Test REST queries of the ProjectData service"/>
        <IconUrl DefaultValue="~remoteAppUrl/Images/NewIcon.png" />
        <SupportUrl DefaultValue="[Insert the URL of a page that provides support information for the app]" />
        <Hosts>
            <Host Name="Project" />
        </Hosts>
        <DefaultSettings>
            <SourceLocation DefaultValue="~remoteAppUrl/AddIn/Home/Home.html" />
        </DefaultSettings>
        <Permissions>ReadWriteDocument</Permissions>
    </OfficeApp>
    ```

## <a name="creating-the-html-content-for-the-helloprojectodata-add-in"></a><span data-ttu-id="be649-180">HelloProjectOData アドインの HTML コンテンツを作成する</span><span class="sxs-lookup"><span data-stu-id="be649-180">Creating the HTML content for the HelloProjectOData add-in</span></span>

<span data-ttu-id="be649-181">**HelloProjectOData** アドインは、デバッグとエラー出力を含むサンプルです。これは、実稼働環境での使用を目的としたものではありません。</span><span class="sxs-lookup"><span data-stu-id="be649-181">The **HelloProjectOData** add-in is a sample that includes debugging and error output; it is not intended for production use.</span></span> <span data-ttu-id="be649-182">Before you start coding the HTML content, design the UI and user experience for the add-in, and outline the JavaScript functions that interact with the HTML code.</span><span class="sxs-lookup"><span data-stu-id="be649-182">Before you start coding the HTML content, design the UI and user experience for the add-in, and outline the JavaScript functions that interact with the HTML code.</span></span> <span data-ttu-id="be649-183">For more information, see[Design guidelines for Office Add-ins](../design/add-in-design.md).</span><span class="sxs-lookup"><span data-stu-id="be649-183">For more information, see[Design guidelines for Office Add-ins](../design/add-in-design.md).</span></span> 

<span data-ttu-id="be649-184">作業ウィンドウには、マニフェストの **DisplayName** 要素の値である上部にアドインの表示名が表示されます。</span><span class="sxs-lookup"><span data-stu-id="be649-184">The task pane shows the add-in display name at the top, which is the value of the **DisplayName** element in the manifest.</span></span> <span data-ttu-id="be649-185">次 **のように** 、HelloProjectOData.html ファイルの body 要素には、他の UI 要素が含まれます。</span><span class="sxs-lookup"><span data-stu-id="be649-185">The **body** element in the HelloProjectOData.html file contains the other UI elements, as follows:</span></span>

- <span data-ttu-id="be649-186">サブタイトルは、**ODATA REST QUERY** など、操作の一般的な機能または種類を表します。</span><span class="sxs-lookup"><span data-stu-id="be649-186">A subtitle indicates the general functionality or type of operation, for example, **ODATA REST QUERY**.</span></span>

- <span data-ttu-id="be649-187">**[ProjectData Endpoint の取得**] ボタンは、関数を呼び出して ProjectData サービスのエンドポイントを取得し、テキスト `setOdataUrl` ボックスに表示します。 </span><span class="sxs-lookup"><span data-stu-id="be649-187">The **Get ProjectData Endpoint** button calls the `setOdataUrl` function to get the endpoint of the **ProjectData** service, and display it in a text box.</span></span> <span data-ttu-id="be649-188">Project が Project Web App と接続されていない場合、アドインはエラー ハンドラーを呼び出して、ポップアップ エラー メッセージを表示します。</span><span class="sxs-lookup"><span data-stu-id="be649-188">If Project is not connected with Project Web App, the add-in calls an error handler to display a pop-up error message.</span></span>

- <span data-ttu-id="be649-189">アドイン **が有効な** OData エンドポイントを取得するまで、[すべてのプロジェクトの比較] ボタンは無効になります。</span><span class="sxs-lookup"><span data-stu-id="be649-189">The **Compare All Projects** button is disabled until the add-in gets a valid OData endpoint.</span></span> <span data-ttu-id="be649-190">ボタンを選択すると、関数が呼び出され、REST クエリを使用して ProjectData サービスからプロジェクト コストと作業データ `retrieveOData` **を取得** します。</span><span class="sxs-lookup"><span data-stu-id="be649-190">When you select the button, it calls the `retrieveOData` function, which uses a REST query to get project cost and work data from the **ProjectData** service.</span></span>

- <span data-ttu-id="be649-p125">テーブルには、プロジェクト コスト、実績コスト、作業、および達成率の平均値が表示されます。このテーブルでは、現在アクティブなプロジェクトの値と平均値との比較も行われます。現在の値が全プロジェクトの平均値より大きい場合は、その値が赤で表示されます。現在の値が平均値より小さい場合は、その値が緑で表示されます。現在の値がない場合は、青の **NA** が表示されます。</span><span class="sxs-lookup"><span data-stu-id="be649-p125">A table displays the average values for project cost, actual cost, work, and percent complete. The table also compares the current active project values with the average. If the current value is greater than the average for all projects, the value is displayed as red. If the current value is less than the average, the value is displayed as green. If the current value is not available, the table displays a blue **NA**.</span></span>

    <span data-ttu-id="be649-196">関数 `retrieveOData` は、テーブル `parseODataResult` の値を計算して表示する関数を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="be649-196">The `retrieveOData` function calls the `parseODataResult` function, which calculates and displays values for the table.</span></span>

    > [!NOTE]
    > <span data-ttu-id="be649-197">この例では、アクティブなプロジェクトのコストと作業のデータは発行された値から導出されます。</span><span class="sxs-lookup"><span data-stu-id="be649-197">In this example, cost and work data for the active project are derived from the published values.</span></span> <span data-ttu-id="be649-198">Project で値を変更した場合、**ProjectData** サービスは、そのプロジェクトが発行されるまで変更を反映しません。</span><span class="sxs-lookup"><span data-stu-id="be649-198">If you change values in Project, the **ProjectData** service does not have the changes until the project is published.</span></span>

### <a name="procedure-4-to-create-the-html-content"></a><span data-ttu-id="be649-p127">手順 4. HTML コンテンツを作成するには</span><span class="sxs-lookup"><span data-stu-id="be649-p127">Procedure 4. To create the HTML content</span></span>

1. <span data-ttu-id="be649-201">Home.htm l ファイルの head 要素で、アドインで使用する CSS ファイル用の追加のリンク要素を追加します。</span><span class="sxs-lookup"><span data-stu-id="be649-201">In the **head** element of the Home.html file, add any additional **link** elements for CSS files that your add-in uses.</span></span> <span data-ttu-id="be649-202">Visual Studio プロジェクト テンプレートには、カスタム CSS スタイルに使用できる App.css ファイルのリンクが含まれています。</span><span class="sxs-lookup"><span data-stu-id="be649-202">The Visual Studio project template includes a link for the App.css file that you can use for custom CSS styles.</span></span>

2. <span data-ttu-id="be649-203">アドインで **使用する** JavaScript ライブラリのスクリプト要素を追加します。</span><span class="sxs-lookup"><span data-stu-id="be649-203">Add any additional **script** elements for JavaScript libraries that your add-in uses.</span></span> <span data-ttu-id="be649-204">プロジェクト テンプレートには、スクリプト フォルダー内の jQuery- _[version]_.js、office.js、MicrosoftAjax.jsファイルのリンク **が含** まれています。</span><span class="sxs-lookup"><span data-stu-id="be649-204">The project template includes links for the jQuery- _[version]_.js, office.js, and MicrosoftAjax.js files in the **Scripts** folder.</span></span>

    > [!NOTE]
    > <span data-ttu-id="be649-p130">アドインを展開する前に、office.js の参照と jQuery の参照をコンテンツ配信ネットワーク (CDN) の参照に変更してください。CDN の参照は最新のバージョンと高いパフォーマンスを提供します。</span><span class="sxs-lookup"><span data-stu-id="be649-p130">Before you deploy the add-in, change the office.js reference and the jQuery reference to the content delivery network (CDN) reference. The CDN reference provides the most recent version and better performance.</span></span>

    <span data-ttu-id="be649-207">**HelloProjectOData** アドインは、ポップアップ メッセージにSurfaceErrors.jsファイルも使用します。</span><span class="sxs-lookup"><span data-stu-id="be649-207">The **HelloProjectOData** add-in also uses the SurfaceErrors.js file, which displays errors in a pop-up message.</span></span> <span data-ttu-id="be649-208">テキスト エディターを使用して [Project 2013](../project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md)の最初の作業ウィンドウ アドインを作成するの 「堅牢なプログラミング」セクションからコードをコピーし **、HelloProjectODataWeb** プロジェクトの **Scripts\Office** フォルダーに SurfaceErrors.js ファイルを追加できます。</span><span class="sxs-lookup"><span data-stu-id="be649-208">You can copy the code from the _Robust Programming_ section of [Create your first task pane add-in for Project 2013 by using a text editor](../project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md), and then add a SurfaceErrors.js file in the **Scripts\Office** folder of the **HelloProjectODataWeb** project.</span></span>

    <span data-ttu-id="be649-209">head 要素の更新された HTML コードを次に **示します。** その他の行は、SurfaceErrors.jsです。</span><span class="sxs-lookup"><span data-stu-id="be649-209">Following is the updated HTML code for the **head** element, with the additional line for the SurfaceErrors.js file.</span></span>

    ```HTML
    <!DOCTYPE html>
    <html>
    <head>
    <meta charset="UTF-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
    <title>Test ProjectData Service</title>

    <link rel="stylesheet" type="text/css" href="../Content/Office.css" />

    <!-- Add your CSS styles to the following file. -->
    <link rel="stylesheet" type="text/css" href="../Content/App.css" />

    <!-- Use the CDN reference to the mini-version of jQuery when deploying your add-in. -->
    <!--<script src="http://ajax.aspnetcdn.com/ajax/jquery/jquery-1.9.0.min.js"></script> -->
    <script src="../Scripts/jquery-1.7.1.js"></script>

    <!-- Use the CDN reference to office.js when deploying your add-in. -->
    <!--<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>-->

    <!-- Use the local script references for Office.js to enable offline debugging -->
    <script src="../Scripts/Office/1.0/MicrosoftAjax.js"></script>
    <script src="../Scripts/Office/1.0/Office.js"></script>

    <!-- Add your JavaScript to the following files. -->
    <script src="../Scripts/HelloProjectOData.js"></script>
    <script src="../Scripts/SurfaceErrors.js"></script>
    </head>
    <body>
    <!-- See the code in Step 3. -->
    </body>
    </html>
    ```

3. <span data-ttu-id="be649-210">body 要素 **で** 、テンプレートから既存のコードを削除し、ユーザー インターフェイスのコードを追加します。</span><span class="sxs-lookup"><span data-stu-id="be649-210">In the **body** element, delete the existing code from the template, and then add the code for the user interface.</span></span> <span data-ttu-id="be649-211">要素にデータを入れるか、要素を jQuery ステートメントで操作する場合、その要素に一意の **id** 属性が含まれている必要があります。</span><span class="sxs-lookup"><span data-stu-id="be649-211">If an element is to be filled with data or manipulated by a jQuery statement, the element must include a unique **id** attribute.</span></span> <span data-ttu-id="be649-212">次のコードでは **、jQuery** 関数が使用するボタン、span、**および td** (テーブル セル定義) 要素の id 属性が太字で表示されます。</span><span class="sxs-lookup"><span data-stu-id="be649-212">In the following code, the **id** attributes for the **button**, **span**, and **td** (table cell definition) elements that jQuery functions use are shown in bold font.</span></span>

   <span data-ttu-id="be649-213">次の HTML は、グラフィックス イメージ (会社のロゴなど) を追加します。</span><span class="sxs-lookup"><span data-stu-id="be649-213">The following HTML adds a graphic image, which could be a company logo.</span></span> <span data-ttu-id="be649-214">選択したロゴを使用するか、Project SDK ダウンロードから NewLogo.png ファイルをコピーしてから、ソリューション エクスプローラーを使用してファイルをフォルダーに追加 `HelloProjectODataWeb\Images` できます。</span><span class="sxs-lookup"><span data-stu-id="be649-214">You can use a logo of your choice, or copy the NewLogo.png file from the Project 2013 SDK download, and then use **Solution Explorer** to add the file to the `HelloProjectODataWeb\Images` folder.</span></span>

    ```HTML
    <body>
        <div id="SectionContent">
        <div id="odataQueries">
            ODATA REST QUERY
        </div>
        <div id="odataInfo">
            <button class="button-wide" onclick="setOdataUrl()">Get ProjectData Endpoint</button>
            <br /><br />
            <span class="rest" id="projectDataEndPoint">Endpoint of the 
                <strong>ProjectData</strong> service</span>
            <br />
        </div>
        <div id="compareProjectData">
            <button class="button-wide" disabled="disabled" id="compareProjects"
            onclick="retrieveOData()">Compare All Projects</button>
            <br />
        </div>
        </div>
        <div id="corpInfo">
            <table class="infoTable" aria-readonly="True" style="width: 100%;">
                <tr>
                    <td class="heading_leftCol"></td>
                    <td class="heading_midCol"><strong>Average</strong></td>
                    <td class="heading_rightCol"><strong>Current</strong></td>
                </tr>
                <tr>
                    <td class="row_leftCol"><strong>Project Cost</strong></td>
                    <td class="row_midCol" id="AverageProjectCost">&amp;nbsp;</td>
                    <td class="row_rightCol" id="CurrentProjectCost">&amp;nbsp;</td>
                </tr>
                <tr>
                    <td class="row_leftCol"><strong>Project Actual Cost</strong></td>
                    <td class="row_midCol" id="AverageProjectActualCost">&amp;nbsp;</td>
                    <td class="row_rightCol" id="CurrentProjectActualCost">&amp;nbsp;</td>
                </tr>
                <tr>
                    <td class="row_leftCol"><strong>Project Work</strong></td>
                    <td class="row_midCol" id="AverageProjectWork">&amp;nbsp;</td>
                    <td class="row_rightCol" id="CurrentProjectWork">&amp;nbsp;</td>
                </tr>
                <tr>
                    <td class="row_leftCol"><strong>Project % Complete</strong></td>
                    <td class="row_midCol" id="AverageProjectPercentComplete">&amp;nbsp;</td>
                    <td class="row_rightCol" id="CurrentProjectPercentComplete">&amp;nbsp;</td>
                </tr>
            </table>
        </div>
        <img alt="Corporation" class="logo" src="../../images/NewLogo.png" />
        <br />
        <textarea id="odataText" rows="12" cols="40"></textarea>
    </body>
    ```

## <a name="creating-the-javascript-code-for-the-add-in"></a><span data-ttu-id="be649-215">このアドインの JavaScript コードを作成する</span><span class="sxs-lookup"><span data-stu-id="be649-215">Creating the JavaScript code for the add-in</span></span>

<span data-ttu-id="be649-216">Project 作業ウィンドウ アドインのテンプレートには、一般的な Office 2013 アドインのドキュメント内のデータの基本的な取得および設定アクションを示す既定の初期化コードが含まれています。</span><span class="sxs-lookup"><span data-stu-id="be649-216">The template for a Project task pane add-in includes default initialization code that is designed to demonstrate basic get and set actions for data in a document for a typical Office 2013 add-in.</span></span> <span data-ttu-id="be649-217">Project 2013 では、アクティブ なプロジェクトに書き込むアクションはサポートされていません **。HelloProjectOData** アドインはメソッドを使用しないので、関数内のスクリプトを削除し、既定の HelloProjectOData.js ファイルで関数と関数 `getSelectedDataAsync` `Office.initialize` `setData` `getData` を削除できます。</span><span class="sxs-lookup"><span data-stu-id="be649-217">Because Project 2013 does not support actions that write to the active project, and the **HelloProjectOData** add-in does not use the `getSelectedDataAsync` method, you can delete the script within the `Office.initialize` function, and delete the `setData` function and `getData` function in the default HelloProjectOData.js file.</span></span>

<span data-ttu-id="be649-218">JavaScript には、REST クエリ用のグローバル定数と、いくつかの関数で使用されるグローバル変数が含まれています。</span><span class="sxs-lookup"><span data-stu-id="be649-218">The JavaScript includes global constants for the REST query and global variables that are used in several functions.</span></span> <span data-ttu-id="be649-219">**[ProjectData Endpoint** の取得] ボタンは関数を呼び出し、グローバル変数を初期化し、Projectに接続されているかどうかを `setOdataUrl` Project Web App。</span><span class="sxs-lookup"><span data-stu-id="be649-219">The **Get ProjectData Endpoint** button calls the `setOdataUrl` function, which initializes the global variables and determines whether Project is connected with Project Web App.</span></span>

<span data-ttu-id="be649-220">HelloProjectOData.js ファイルの残りの部分には、ユーザーが [すべてのプロジェクトの比較] を選択すると関数が呼び出され、平均が計算され、比較テーブルに色と単位の書式が設定された値が設定されます。 `retrieveOData`  `parseODataResult`</span><span class="sxs-lookup"><span data-stu-id="be649-220">The remainder of the HelloProjectOData.js file includes two functions: the `retrieveOData` function is called when the user selects **Compare All Projects**; and the `parseODataResult` function calculates averages and then populates the comparison table with values that are formatted for color and units.</span></span>

### <a name="procedure-5-to-create-the-javascript-code"></a><span data-ttu-id="be649-p136">手順 5. JavaScript コードを作成するには</span><span class="sxs-lookup"><span data-stu-id="be649-p136">Procedure 5. To create the JavaScript code</span></span>

1. <span data-ttu-id="be649-223">既定のファイル内のすべてのコードをHelloProjectOData.jsし、グローバル変数と `**` tialize' 関数Office.ini追加します。</span><span class="sxs-lookup"><span data-stu-id="be649-223">Delete all code in the default HelloProjectOData.js file, and then add the global variables and `**`Office.initialize\` function.</span></span> <span data-ttu-id="be649-224">すべて大文字の変数名は、それらが定数であることを示しており、それらは後で **_pwa** 変数と共に使用されて、この例の REST クエリが作成されます。</span><span class="sxs-lookup"><span data-stu-id="be649-224">Variable names that are all capitals imply that they are constants; they are later used with the **_pwa** variable to create the REST query in this example.</span></span>

    ```js
    var PROJDATA = "/_api/ProjectData";
    var PROJQUERY = "/Projects?";
    var QUERY_FILTER = "$filter=ProjectName ne 'Timesheet Administrative Work Items'";
    var QUERY_SELECT1 = "&amp;$select=ProjectId, ProjectName";
    var QUERY_SELECT2 = ", ProjectCost, ProjectWork, ProjectPercentCompleted, ProjectActualCost";
    var _pwa;           // URL of Project Web App.
    var _projectUid;    // GUID of the active project.
    var _docUrl;        // Path of the project document.
    var _odataUrl = ""; // URL of the OData service: http[s]://ServerName /ProjectServerName /_api/ProjectData

    // The initialize function is required for all add-ins.
    Office.initialize = function (reason) {
        // Checks for the DOM to load using the jQuery ready function.
        $(document).ready(function () {
            // After the DOM is loaded, app-specific code can run.
        });
    }
    ```

2. <span data-ttu-id="be649-225">追加 `setOdataUrl` および関連する関数。</span><span class="sxs-lookup"><span data-stu-id="be649-225">Add `setOdataUrl` and related functions.</span></span> <span data-ttu-id="be649-226">関数 `setOdataUrl` はグローバル `getProjectGuid` 変数 `getDocumentUrl` を呼び出して初期化します。</span><span class="sxs-lookup"><span data-stu-id="be649-226">The `setOdataUrl` function calls `getProjectGuid` and `getDocumentUrl` to initialize the global variables.</span></span> <span data-ttu-id="be649-227">[getProjectFieldAsync](/javascript/api/office/office.document)メソッドでは、コールバック パラメーターの匿名関数を使用すると、jQuery ライブラリのメソッドを使用して [すべてのプロジェクトの比較] ボタンを有効にし `removeAttr` **、ProjectData** サービスの URL を表示します。</span><span class="sxs-lookup"><span data-stu-id="be649-227">In the [getProjectFieldAsync method](/javascript/api/office/office.document), the anonymous function for the  _callback_ parameter enables the **Compare All Projects** button by using the `removeAttr` method in the jQuery library, and then displays the URL of the **ProjectData** service.</span></span> <span data-ttu-id="be649-228">Project が Project Web App と接続されていない場合、この関数はエラーをスローし、それによってポップアップ エラー メッセージが表示されます。</span><span class="sxs-lookup"><span data-stu-id="be649-228">If Project is not connected with Project Web App, the function throws an error, which displays a pop-up error message.</span></span> <span data-ttu-id="be649-229">このSurfaceErrors.jsには、メソッドが含 `throwError` まれます。</span><span class="sxs-lookup"><span data-stu-id="be649-229">The SurfaceErrors.js file includes the `throwError` method.</span></span>

   > [!NOTE]
   > <span data-ttu-id="be649-230">Visual Studio を Project Server コンピューター上で実行する場合、**F5** デバッグを使用するには、**_pwa** グローバル変数を初期化する行の後にあるコードのコメントを解除します。</span><span class="sxs-lookup"><span data-stu-id="be649-230">If you run Visual Studio on the Project Server computer, to use **F5** debugging, uncomment the code after the line that initializes the **_pwa** global variable.</span></span> <span data-ttu-id="be649-231">サーバー コンピューターでデバッグするときに jQuery メソッドを使用Projectするには、サーバー URL の値をPWA `ajax` `localhost` 必要があります。リモート コンピューターでVisual Studioを実行する場合 `localhost` 、URL は必要ありません。</span><span class="sxs-lookup"><span data-stu-id="be649-231">To enable using the jQuery `ajax` method when debugging on the Project Server computer, you must set the `localhost` value for the PWA URL.If you run Visual Studio on a remote computer, the  `localhost` URL is not required.</span></span> <span data-ttu-id="be649-232">Before you deploy the add-in, comment out that code.</span><span class="sxs-lookup"><span data-stu-id="be649-232">Before you deploy the add-in, comment out that code.</span></span>

    ```js
    function setOdataUrl() {
        Office.context.document.getProjectFieldAsync(
            Office.ProjectProjectFields.ProjectServerUrl,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    _pwa = String(asyncResult.value.fieldValue);

                    // If you debug with Visual Studio on a local Project Server computer, 
                    // uncomment the following lines to use the localhost URL.
                    //var localhost = location.host.split(":", 1);
                    //var pwaStartPosition = _pwa.lastIndexOf("/");
                    //var pwaLength = _pwa.length - pwaStartPosition;
                    //var pwaName = _pwa.substr(pwaStartPosition, pwaLength);
                    //_pwa = location.protocol + "//" + localhost + pwaName;

                    if (_pwa.substring(0, 4) == "http") {
                        _odataUrl = _pwa + PROJDATA;
                        $("#compareProjects").removeAttr("disabled");
                        getProjectGuid();
                    }
                    else {
                        _odataUrl = "No connection!";
                        throwError(_odataUrl, "You are not connected to Project Web App.");
                    }
                    getDocumentUrl();
                    $("#projectDataEndPoint").text(_odataUrl);
                }
                else {
                    throwError(asyncResult.error.name, asyncResult.error.message);
                }
            }
        );
    }

    // Get the GUID of the active project.
    function getProjectGuid() {
        Office.context.document.getProjectFieldAsync(
            Office.ProjectProjectFields.GUID,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    _projectUid = asyncResult.value.fieldValue;
                }
                else {
                    throwError(asyncResult.error.name, asyncResult.error.message);
                }
            }
        );
    }

    // Get the path of the project in Project web app, which is in the form <>\ProjectName .
    function getDocumentUrl() {
        _docUrl = "Document path:\r\n" + Office.context.document.url;
    }
    ```

3. <span data-ttu-id="be649-233">REST クエリ `retrieveOData` の値を連結し、jQuery で関数を呼び出して ProjectData サービスから要求されたデータを取得する関数を `ajax` **追加** します。</span><span class="sxs-lookup"><span data-stu-id="be649-233">Add the `retrieveOData` function, which concatenates values for the REST query and then calls the `ajax` function in jQuery to get the requested data from the **ProjectData** service.</span></span> <span data-ttu-id="be649-234">**support.cors 変数を使用** すると、関数とのクロスオリジン リソース共有 (CORS) が有効 `ajax` になります。</span><span class="sxs-lookup"><span data-stu-id="be649-234">The **support.cors** variable enables cross-origin resource sharing (CORS) with the `ajax` function.</span></span> <span data-ttu-id="be649-235">**support.cors ステートメントが** 見つからないか、false に設定されている場合、この関数はトランスポート エラー `ajax` **を返** します。</span><span class="sxs-lookup"><span data-stu-id="be649-235">If the **support.cors** statement is missing or is set to **false**, the `ajax` function returns a **No transport** error.</span></span>

   > [!NOTE]
   > <span data-ttu-id="be649-p141">次に示すコードは、Project Server 2013 のオンプレミスのインストールで動作します。Project on the web の場合は、トークン ベースの認証に OAuth を使用できます。詳細については、「[Office アドインにおける同一生成元ポリシーの制限への対処](../develop/addressing-same-origin-policy-limitations.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="be649-p141">The following code works with an on-premises installation of Project Server 2013. For Project on the web, you can use OAuth for token-based authentication. For more information, see [Addressing same-origin policy limitations in Office Add-ins](../develop/addressing-same-origin-policy-limitations.md).</span></span>

   <span data-ttu-id="be649-239">呼び `ajax` 出しでは _、headers_ パラメーターまたは _beforeSend パラメーターのいずれかを使用_ できます。</span><span class="sxs-lookup"><span data-stu-id="be649-239">In the `ajax` call, you can use either the _headers_ parameter or the _beforeSend_ parameter.</span></span> <span data-ttu-id="be649-240">complete _パラメーター_ は匿名関数なので、 の変数と同じスコープに入ります `retrieveOData` 。</span><span class="sxs-lookup"><span data-stu-id="be649-240">The _complete_ parameter is an anonymous function so that it is in the same scope as the variables in `retrieveOData`.</span></span> <span data-ttu-id="be649-241">complete パラメーターの  _関数は_ 、コントロールに結果を表示し、JSON 応答を解析して表示するメソッド `odataText` `parseODataResult` も呼び出します。</span><span class="sxs-lookup"><span data-stu-id="be649-241">The function for the  _complete_ parameter displays results in the `odataText` control and also calls the `parseODataResult` method to parse and display the JSON response.</span></span> <span data-ttu-id="be649-242">error _パラメーター_ は、コントロールにエラー メッセージを書き込み、メソッドを使用してポップアップ メッセージを表示する名前付き `getProjectDataErrorHandler` `odataText` `throwError` 関数を指定します。</span><span class="sxs-lookup"><span data-stu-id="be649-242">The _error_ parameter specifies the named `getProjectDataErrorHandler` function, which writes an error message to the `odataText` control and also uses the `throwError` method to display a pop-up message.</span></span>

    ```js
    // Functions to get and parse the Project Server reporting data./

    // Get data about all projects on Project Server,
    // by using a REST query with the ajax method in jQuery.
    function retrieveOData() {
        var restUrl = _odataUrl + PROJQUERY + QUERY_FILTER + QUERY_SELECT1 + QUERY_SELECT2;
        var accept = "application/json; odata=verbose";
        accept.toLocaleLowerCase();

        // Enable cross-origin scripting (required by jQuery 1.5 and later).
        // This does not work with Project on the web.
        $.support.cors = true;

        $.ajax({
            url: restUrl,
            type: "GET",
            contentType: "application/json",
            data: "",      // Empty string for the optional data.
            //headers: { "Accept": accept },
            beforeSend: function (xhr) {
                xhr.setRequestHeader("ACCEPT", accept);
            },
            complete: function (xhr, textStatus) {
                // Create a message to display in the text box.
                var message = "\r\ntextStatus: " + textStatus +
                    "\r\nContentType: " + xhr.getResponseHeader("Content-Type") +
                    "\r\nStatus: " + xhr.status +
                    "\r\nResponseText:\r\n" + xhr.responseText;

                // xhr.responseText is the result from an XmlHttpRequest, which
                // contains the JSON response from the OData service.
                parseODataResult(xhr.responseText, _projectUid);

                // Write the document name, response header, status, and JSON to the odataText control.
                $("#odataText").text(_docUrl);
                $("#odataText").append("\r\nREST query:\r\n" + restUrl);
                $("#odataText").append(message);

                if (xhr.status != 200 &amp;&amp; xhr.status != 1223 &amp;&amp; xhr.status != 201) {
                    $("#odataInfo").append("<div>" + htmlEncode(restUrl) + "</div>");
                }
            },
            error: getProjectDataErrorHandler
        });
    }

    function getProjectDataErrorHandler(data, errorCode, errorMessage) {
        $("#odataText").text("Error code: " + errorCode + "\r\nError message: \r\n"
        + errorMessage);
        throwError(errorCode, errorMessage);
    }
    ```

4. <span data-ttu-id="be649-243">OData サービスからの JSON 応答を逆シリアル化して処理するメソッド `parseODataResult` を追加します。</span><span class="sxs-lookup"><span data-stu-id="be649-243">Add the `parseODataResult` method, which deserializes and processes the JSON response from the OData service.</span></span> <span data-ttu-id="be649-244">このメソッドは、コストデータと作業データの平均値を小数点以下 1 桁または 2 桁の精度で計算し、正しい色で値を書式設定し、単位 `parseODataResult` ( **$** 、 **hrs**、 **%** または ) を追加し、指定した表のセルに値を表示します。</span><span class="sxs-lookup"><span data-stu-id="be649-244">The `parseODataResult` method calculates average values of the cost and work data to an accuracy of one or two decimal places, formats values with the correct color and adds a unit ( **$**, **hrs**, or **%**), and then displays the values in specified table cells.</span></span>

   <span data-ttu-id="be649-245">アクティブ なプロジェクトの GUID が値と一致する場合 `ProjectId` 、変数 `myProjectIndex` はプロジェクト インデックスに設定されます。</span><span class="sxs-lookup"><span data-stu-id="be649-245">If the GUID of the active project matches the `ProjectId` value, the `myProjectIndex` variable is set to the project index.</span></span> <span data-ttu-id="be649-246">アクティブ なプロジェクトが Project Server で発行された場合、メソッドは、そのプロジェクトのコストと作業データ `myProjectIndex` `parseODataResult` を書式設定して表示します。</span><span class="sxs-lookup"><span data-stu-id="be649-246">If `myProjectIndex` indicates the active project is published on Project Server, the `parseODataResult` method formats and displays cost and work data for that project.</span></span> <span data-ttu-id="be649-247">アクティブなプロジェクトが発行されていない場合は、アクティブなプロジェクトの値は青い **NA** と表示されます。</span><span class="sxs-lookup"><span data-stu-id="be649-247">If the active project is not published, values for the active project are displayed as a blue **NA**.</span></span>

    ```js
    // Calculate the average values of actual cost, cost, work, and percent complete
    // for all projects, and compare with the values for the current project.
    function parseODataResult(oDataResult, currentProjectGuid) {
        // Deserialize the JSON string into a JavaScript object.
        var res = Sys.Serialization.JavaScriptSerializer.deserialize(oDataResult);
        var len = res.d.results.length;
        var projActualCost = 0;
        var projCost = 0;
        var projWork = 0;
        var projPercentCompleted = 0;
        var myProjectIndex = -1;
        for (i = 0; i < len; i++) {
            // If the current project GUID matches the GUID from the OData query,  
            // store the project index.
            if (currentProjectGuid.toLocaleLowerCase() == res.d.results[i].ProjectId) {
                myProjectIndex = i;
            }
            projCost += Number(res.d.results[i].ProjectCost);
            projWork += Number(res.d.results[i].ProjectWork);
            projActualCost += Number(res.d.results[i].ProjectActualCost);
            projPercentCompleted += Number(res.d.results[i].ProjectPercentCompleted);
        }
        var avgProjCost = projCost / len;
        var avgProjWork = projWork / len;
        var avgProjActualCost = projActualCost / len;
        var avgProjPercentCompleted = projPercentCompleted / len;

        // Round off cost to two decimal places, and round off other values to one decimal place.
        avgProjCost = avgProjCost.toFixed(2);
        avgProjWork = avgProjWork.toFixed(1);
        avgProjActualCost = avgProjActualCost.toFixed(2);
        avgProjPercentCompleted = avgProjPercentCompleted.toFixed(1);

        // Display averages in the table, with the correct units.
        document.getElementById("AverageProjectCost").innerHTML = "$"
            + avgProjCost;
        document.getElementById("AverageProjectActualCost").innerHTML
            = "$" + avgProjActualCost;
        document.getElementById("AverageProjectWork").innerHTML
            = avgProjWork + " hrs";
        document.getElementById("AverageProjectPercentComplete").innerHTML
            = avgProjPercentCompleted + "%";

        // Calculate and display values for the current project.
        if (myProjectIndex != -1) {
            var myProjCost = Number(res.d.results[myProjectIndex].ProjectCost);
            var myProjWork = Number(res.d.results[myProjectIndex].ProjectWork);
            var myProjActualCost = Number(res.d.results[myProjectIndex].ProjectActualCost);
            var myProjPercentCompleted =
            Number(res.d.results[myProjectIndex].ProjectPercentCompleted);

            myProjCost = myProjCost.toFixed(2);
            myProjWork = myProjWork.toFixed(1);
            myProjActualCost = myProjActualCost.toFixed(2);
            myProjPercentCompleted = myProjPercentCompleted.toFixed(1);

            document.getElementById("CurrentProjectCost").innerHTML = "$" + myProjCost;

            if (Number(myProjCost) <= Number(avgProjCost)) {
                document.getElementById("CurrentProjectCost").style.color = "green"
            }
            else {
                document.getElementById("CurrentProjectCost").style.color = "red"
            }

            document.getElementById("CurrentProjectActualCost").innerHTML = "$" + myProjActualCost;

            if (Number(myProjActualCost) <= Number(avgProjActualCost)) {
                document.getElementById("CurrentProjectActualCost").style.color = "green"
            }
            else {
                document.getElementById("CurrentProjectActualCost").style.color = "red"
            }

            document.getElementById("CurrentProjectWork").innerHTML = myProjWork + " hrs";

            if (Number(myProjWork) <= Number(avgProjWork)) {
                document.getElementById("CurrentProjectWork").style.color = "red"
            }
            else {
                document.getElementById("CurrentProjectWork").style.color = "green"
            }

            document.getElementById("CurrentProjectPercentComplete").innerHTML = myProjPercentCompleted + "%";

            if (Number(myProjPercentCompleted) <= Number(avgProjPercentCompleted)) {
                document.getElementById("CurrentProjectPercentComplete").style.color = "red"
            }
            else {
                document.getElementById("CurrentProjectPercentComplete").style.color = "green"
            }
        }
        else {
            document.getElementById("CurrentProjectCost").innerHTML = "NA";
            document.getElementById("CurrentProjectCost").style.color = "blue"

            document.getElementById("CurrentProjectActualCost").innerHTML = "NA";
            document.getElementById("CurrentProjectActualCost").style.color = "blue"

            document.getElementById("CurrentProjectWork").innerHTML = "NA";
            document.getElementById("CurrentProjectWork").style.color = "blue"

            document.getElementById("CurrentProjectPercentComplete").innerHTML = "NA";
            document.getElementById("CurrentProjectPercentComplete").style.color = "blue"
        }
    }
    ```

## <a name="testing-the-helloprojectodata-add-in"></a><span data-ttu-id="be649-248">HelloProjectOData アドインのテスト</span><span class="sxs-lookup"><span data-stu-id="be649-248">Testing the HelloProjectOData add-in</span></span>

<span data-ttu-id="be649-249">Visual Studio 2015 年 2015 年に **HelloProjectOData** アドインをテストおよびデバッグするには、Project Professional 2013 を開発用コンピューターにインストールする必要があります。</span><span class="sxs-lookup"><span data-stu-id="be649-249">To test and debug the **HelloProjectOData** add-in with Visual Studio 2015, Project Professional 2013 must be installed on the development computer.</span></span> <span data-ttu-id="be649-250">他のテスト シナリオを有効にするには、Project がローカル コンピューター上のファイルに対して開くか、Project Web App と接続するかを選択できることを確認してください。</span><span class="sxs-lookup"><span data-stu-id="be649-250">To enable different test scenarios, ensure that you can choose whether Project opens for files on the local computer or connects with Project Web App.</span></span> <span data-ttu-id="be649-251">たとえば、次の手順を実行します。</span><span class="sxs-lookup"><span data-stu-id="be649-251">For example, do the following steps.</span></span>

1. <span data-ttu-id="be649-252">リボンの [**ファイル**] タブで、Backstage ビューの [**情報**] タブを選択し、[**アカウントの管理**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="be649-252">On the **FILE** tab on the ribbon, choose the **Info** tab in the Backstage view, and then choose **Manage Accounts**.</span></span>

2. <span data-ttu-id="be649-253">[Web アプリ **Projectアカウント**] ダイアログ ボックスの[利用可能なアカウント] リストには、ローカル コンピューター アカウントに加Project Web App複数のアカウントを **含** めできます。</span><span class="sxs-lookup"><span data-stu-id="be649-253">In the **Project web app Accounts** dialog box, the **Available accounts** list can have multiple Project Web App accounts in addition to the local **Computer** account.</span></span> <span data-ttu-id="be649-254">[ **開始時**] セクションで、[ **アカウントを選択する**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="be649-254">In the **When starting** section, select **Choose an account**.</span></span>

3. <span data-ttu-id="be649-255">Project を閉じます。それにより、Visual Studio でアドインのデバッグ用に Project 起動できるようになります。</span><span class="sxs-lookup"><span data-stu-id="be649-255">Close Project so that Visual Studio can start it for debugging the add-in.</span></span>

<span data-ttu-id="be649-256">基本的なテストでは、次のことを行う必要があります。</span><span class="sxs-lookup"><span data-stu-id="be649-256">Basic tests should include the following:</span></span>

- <span data-ttu-id="be649-257">アドインを Visual Studio から実行し、コストと作業のデータが含まれている Project Web App から発行済みのプロジェクトを開きます。</span><span class="sxs-lookup"><span data-stu-id="be649-257">Run the add-in from Visual Studio, and then open a published project from Project Web App that contains cost and work data.</span></span> <span data-ttu-id="be649-258">アドインに **ProjectData** エンドポイントが表示され、コストと作業時間のデータがテーブルに正しく表示されていることを確認します。</span><span class="sxs-lookup"><span data-stu-id="be649-258">Verify that the add-in displays the **ProjectData** endpoint and correctly displays the cost and work data in the table.</span></span> <span data-ttu-id="be649-259">**odataText** コントロール内の出力により、REST クエリとその他の情報を確認できます。</span><span class="sxs-lookup"><span data-stu-id="be649-259">You can use the output in the **odataText** control to check the REST query and other information.</span></span>

- <span data-ttu-id="be649-260">アドインを再び実行し、Project の起動時に [**ログイン**] ダイアログ ボックスでローカル コンピューターのプロファイルを選択します。</span><span class="sxs-lookup"><span data-stu-id="be649-260">Run the add-in again, where you choose the local computer profile in the **Login** dialog box when Project starts.</span></span> <span data-ttu-id="be649-261">ローカル .mpp ファイルを開き、アドインをテストします。</span><span class="sxs-lookup"><span data-stu-id="be649-261">Open a local .mpp file, and then test the add-in.</span></span> <span data-ttu-id="be649-262">**ProjectData** エンドポイントを取得しようとしたときに、アドインがエラー メッセージを表示することを確認します。</span><span class="sxs-lookup"><span data-stu-id="be649-262">Verify that the add-in displays an error message when you try to get the **ProjectData** endpoint.</span></span>

- <span data-ttu-id="be649-263">アドインを再び実行し、コストと作業のデータのタスクを持つプロジェクトを作成します。</span><span class="sxs-lookup"><span data-stu-id="be649-263">Run the add-in again, where you create a project that has tasks with cost and work data.</span></span> <span data-ttu-id="be649-264">そのプロジェクトを Project Web App に保存することはできますが、発行はしないでください。</span><span class="sxs-lookup"><span data-stu-id="be649-264">You can save the project to Project Web App, but don't publish it.</span></span> <span data-ttu-id="be649-265">アドインが Project Server からのデータを表示するものの、現在のプロジェクトについては **NA** であることを確認します。</span><span class="sxs-lookup"><span data-stu-id="be649-265">Verify that the add-in displays data from Project Server, but **NA** for the current project.</span></span>

### <a name="procedure-6-to-test-the-add-in"></a><span data-ttu-id="be649-p150">手順 6. アドインをテストするには</span><span class="sxs-lookup"><span data-stu-id="be649-p150">Procedure 6. To test the add-in</span></span>

1. <span data-ttu-id="be649-p151">Project Professional 2013 を実行し、Project Web App と接続してから、テスト プロジェクトを作成します。ローカル リソースまたはエンタープライズ リソースにタスクを割り当て、いくつかのタスクに対して達成率のさまざまな値を設定してから、そのプロジェクトを発行します。Project を終了します。それにより、Visual Studio がアドインのデバッグ用に Project を起動できるようになります。</span><span class="sxs-lookup"><span data-stu-id="be649-p151">Run Project Professional 2013, connect with Project Web App, and then create a test project. Assign tasks to local resources or to enterprise resources, set various values of percent complete on some tasks, and then publish the project. Quit Project, which enables Visual Studio to start Project for debugging the add-in.</span></span>

2. <span data-ttu-id="be649-271">Visual Studio で **F5** キーを押します。</span><span class="sxs-lookup"><span data-stu-id="be649-271">In Visual Studio, press **F5**.</span></span> <span data-ttu-id="be649-272">Project Web App にログオンし、前のステップで作成したプロジェクトを開きます。</span><span class="sxs-lookup"><span data-stu-id="be649-272">Log on to Project Web App, and then open the project that you created in the previous step.</span></span> <span data-ttu-id="be649-273">このとき、読み取り専用モードで開くことも、編集モードで開くこともできます。</span><span class="sxs-lookup"><span data-stu-id="be649-273">You can open the project in read-only mode or in edit mode.</span></span>

3. <span data-ttu-id="be649-274">リボンの **[PROJECT]** タブの[Office アドイン] ドロップダウン リストで **、[Hello ProjectData]** を選択します (図 5 を参照)。</span><span class="sxs-lookup"><span data-stu-id="be649-274">On the **PROJECT** tab of the ribbon, in the **Office Add-ins** drop-down list, select **Hello ProjectData** (see Figure 5).</span></span> <span data-ttu-id="be649-275">[ **すべてのプロジェクトを比較**] ボタンが無効化されるはずです。</span><span class="sxs-lookup"><span data-stu-id="be649-275">The **Compare All Projects** button should be disabled.</span></span>

    <span data-ttu-id="be649-276">*図 5. HelloProjectOData アドインの開始*</span><span class="sxs-lookup"><span data-stu-id="be649-276">*Figure 5. Starting the HelloProjectOData add-in*</span></span>

    ![HelloProjectOData アプリをテストします。](../images/pj15-hello-project-data-test-the-app.png)

4. <span data-ttu-id="be649-278">[**ProjectData 概要**] 作業ウィンドウで、[**ProjectData エンドポイントを取得**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="be649-278">In the **Hello ProjectData** task pane, select **Get ProjectData Endpoint**.</span></span> <span data-ttu-id="be649-279">**projectDataEndPoint 行** には **ProjectData** サービスの URL が表示され、[すべてのプロジェクトの比較] ボタンを有効にする必要があります (図 6 を参照)。</span><span class="sxs-lookup"><span data-stu-id="be649-279">The **projectDataEndPoint** line should show the URL of the **ProjectData** service, and the **Compare All Projects** button should be enabled (see Figure 6).</span></span>

5. <span data-ttu-id="be649-280">[**すべてのプロジェクトを比較**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="be649-280">Select **Compare All Projects**.</span></span> <span data-ttu-id="be649-281">アドインは、**ProjectData** サービスからデータを取得している間、一時停止する可能性がありますが、その後、書式化された平均値と現在値をテーブルに表示します。</span><span class="sxs-lookup"><span data-stu-id="be649-281">The add-in may pause while it retrieves data from the **ProjectData** service, and then it should display the formatted average and current values in the table.</span></span>

    <span data-ttu-id="be649-282">*図 6. REST クエリの結果の表示*</span><span class="sxs-lookup"><span data-stu-id="be649-282">*Figure 6. Viewing results of the REST query*</span></span>

    ![REST クエリの結果を表示する。](../images/pj15-hello-project-data-rest-results.png)

6. <span data-ttu-id="be649-284">テキスト ボックス内の出力を調べます。</span><span class="sxs-lookup"><span data-stu-id="be649-284">Examine output in the text box.</span></span> <span data-ttu-id="be649-285">ajax と parseODataResult の呼び出しのドキュメント パス、REST クエリ、状態情報、JSON の結果 **を表示する必要があります**。 </span><span class="sxs-lookup"><span data-stu-id="be649-285">It should show the document path, REST query, status information, and JSON results from the calls to **ajax** and **parseODataResult**.</span></span> <span data-ttu-id="be649-286">出力は、メソッド内のコードを理解、作成、デバッグするのに `parseODataResult` 役立ちます `projCost += Number(res.d.results[i].ProjectCost);` 。</span><span class="sxs-lookup"><span data-stu-id="be649-286">The output helps to understand, create, and debug code in the `parseODataResult` method such as `projCost += Number(res.d.results[i].ProjectCost);`.</span></span>

    <span data-ttu-id="be649-287">次に、1 つのインスタンス内の 3 つのプロジェクトについて、テキストに改行とスペースを追加した出力の例をProject Web Appします。</span><span class="sxs-lookup"><span data-stu-id="be649-287">Following is an example of the output with line breaks and spaces added to the text for clarity, for three projects in a Project Web App instance.</span></span>

    ```json
    Document path: <>\WinProj test1

    REST query:
    http://sphvm-37189/pwa/_api/ProjectData/Projects?$filter=ProjectName ne 'Timesheet Administrative Work Items'
        &amp;$select=ProjectId, ProjectName, ProjectCost, ProjectWork, ProjectPercentCompleted, ProjectActualCost

    textStatus: success
    ContentType: application/json;odata=verbose;charset=utf-8
    Status: 200

    ResponseText:
    {"d":{"results":[
    {"__metadata":
        {"id":"http://sphvm-37189/pwa/_api/ProjectData/Projects(guid'ce3d0d65-3904-e211-96cd-00155d157123')",
        "uri":"http://sphvm-37189/pwa/_api/ProjectData/Projects(guid'ce3d0d65-3904-e211-96cd-00155d157123')",
        "type":"ReportingData.Project"},
        "ProjectId":"ce3d0d65-3904-e211-96cd-00155d157123",
        "ProjectActualCost":"0.000000",
        "ProjectCost":"0.000000",
        "ProjectName":"Task list created in PWA",
        "ProjectPercentCompleted":0,
        "ProjectWork":"16.000000"},
    {"__metadata":
        {"id":"http://sphvm-37189/pwa/_api/ProjectData/Projects(guid'c31023fc-1404-e211-86b2-3c075433b7bd')",
        "uri":"http://sphvm-37189/pwa/_api/ProjectData/Projects(guid'c31023fc-1404-e211-86b2-3c075433b7bd')",
        "type":"ReportingData.Project"},
        "ProjectId":"c31023fc-1404-e211-86b2-3c075433b7bd",
        "ProjectActualCost":"700.000000",
        "ProjectCost":"2400.000000",
        "ProjectName":"WinProj test 2",
        "ProjectPercentCompleted":29,
        "ProjectWork":"48.000000"},
    {"__metadata":
        {"id":"http://sphvm-37189/pwa/_api/ProjectData/Projects(guid'dc81fbb2-b801-e211-9d2a-3c075433b7bd')",
        "uri":"http://sphvm-37189/pwa/_api/ProjectData/Projects(guid'dc81fbb2-b801-e211-9d2a-3c075433b7bd')",
        "type":"ReportingData.Project"},
        "ProjectId":"dc81fbb2-b801-e211-9d2a-3c075433b7bd",
        "ProjectActualCost":"1900.000000",
        "ProjectCost":"5200.000000",
        "ProjectName":"WinProj test1",
        "ProjectPercentCompleted":37,
        "ProjectWork":"104.000000"}
    ]}}
    ```

7. <span data-ttu-id="be649-288">デバッグを停止し (**Shift + F5** キーを押す)、**F5** キーを再び押して、Project の新しいインスタンスを実行します。</span><span class="sxs-lookup"><span data-stu-id="be649-288">Stop debugging (press **Shift + F5**), and then press **F5** again to run a new instance of Project.</span></span> <span data-ttu-id="be649-289">[ **ログイン**] ダイアログ ボックスで、Project Web App ではなく、ローカル  **コンピューター** のプロファイルを選択します。</span><span class="sxs-lookup"><span data-stu-id="be649-289">In the **Login** dialog box, choose the local **Computer** profile, not Project Web App.</span></span> <span data-ttu-id="be649-290">ローカル プロジェクト .mpp ファイルを作成するか開き、[ **ProjectData 概要**] 作業ウィンドウを開き、[ **ProjectData エンドポイントを取得**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="be649-290">Create or open a local project .mpp file, open the **Hello ProjectData** task pane, and then select **Get ProjectData Endpoint**.</span></span> <span data-ttu-id="be649-291">アドインに [接続なし] が **表示されます。**</span><span class="sxs-lookup"><span data-stu-id="be649-291">The add-in should show a **No connection!**</span></span> <span data-ttu-id="be649-292">error (図 7 を参照)、および [ **すべてのプロジェクト** の比較] ボタンを無効にしたままにします。</span><span class="sxs-lookup"><span data-stu-id="be649-292">error (see Figure 7), and the **Compare All Projects** button should remain disabled.</span></span>

   <span data-ttu-id="be649-293">*図 7. Project Web App 接続がない状態でのアドインの使用*</span><span class="sxs-lookup"><span data-stu-id="be649-293">*Figure 7. Using the add-in without a Project web app connection*</span></span>

   ![アプリを使用して、接続Project Web Appします。](../images/pj15-hello-project-data-no-connection.png)

8. <span data-ttu-id="be649-295">デバッグを停止してから、**F5** キーを再び押します。</span><span class="sxs-lookup"><span data-stu-id="be649-295">Stop debugging, and then press **F5** again.</span></span> <span data-ttu-id="be649-296">Project Web App にログオンし、コストと作業のデータが含まれるプロジェクトを作成します。</span><span class="sxs-lookup"><span data-stu-id="be649-296">Log on to Project Web App, and then create a project that contains cost and work data.</span></span> <span data-ttu-id="be649-297">このプロジェクトを保存することはできますが、発行しないでください。</span><span class="sxs-lookup"><span data-stu-id="be649-297">You can save the project, but don't publish it.</span></span>

   <span data-ttu-id="be649-298">**[Hello ProjectData]** 作業ウィンドウで、[すべてのプロジェクトの比較] を選択すると、[現在] 列にフィールドの青い **NA** が表示されます (図 8 を参照)。</span><span class="sxs-lookup"><span data-stu-id="be649-298">In the **Hello ProjectData** task pane, when you select **Compare All Projects**, you should see a blue **NA** for fields in the **Current** column (see Figure 8).</span></span>

   <span data-ttu-id="be649-299">*図 8. 未発行のプロジェクトと他のプロジェクトの比較*</span><span class="sxs-lookup"><span data-stu-id="be649-299">*Figure 8. Comparing an unpublished project with other projects*</span></span>

   ![非公開のプロジェクトを他のユーザーと比較する。](../images/pj15-hello-project-data-not-published.png)

<span data-ttu-id="be649-p159">アドインがこれまでのテストで正常に動作しているとしても、実行する必要のあるテストはまだあります。たとえば、次のことを行います。</span><span class="sxs-lookup"><span data-stu-id="be649-p159">Even if your add-in is working correctly in the previous tests, there are other tests that should be run. For example:</span></span>

- <span data-ttu-id="be649-303">Project Web App からタスクのコストと作業のデータがないプロジェクトを開きます。</span><span class="sxs-lookup"><span data-stu-id="be649-303">Open a project from Project Web App that has no cost or work data for the tasks.</span></span> <span data-ttu-id="be649-304">[**現在**] 列のフィールドには 0 の値が表示されるはずです。</span><span class="sxs-lookup"><span data-stu-id="be649-304">You should see values of zero in the fields in the **Current** column.</span></span>

- <span data-ttu-id="be649-305">タスクがないプロジェクトをテストします。</span><span class="sxs-lookup"><span data-stu-id="be649-305">Test a project that has no tasks.</span></span>

- <span data-ttu-id="be649-p161">アドインに変更を加えて発行した場合は、発行したアドインで再び同様のテストを実行する必要があります。その他の考慮事項については、「[次のステップ](#next-steps)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="be649-p161">If you modify the add-in and publish it, you should run similar tests again with the published add-in. For other considerations, see [Next steps](#next-steps).</span></span>

> [!NOTE]
> <span data-ttu-id="be649-308">**ProjectData** サービスの 1 つのクエリで返されるデータ量には制限があります。データの量はエンティティによって異なります。</span><span class="sxs-lookup"><span data-stu-id="be649-308">There are limits to the amount of data that can be returned in one query of the **ProjectData** service; the amount of data varies by entity.</span></span> <span data-ttu-id="be649-309">たとえば、エンティティ セットの既定の制限はクエリごとに 100 プロジェクトですが、エンティティ セットの既定の制限は `Projects` `Risks` 200 です。</span><span class="sxs-lookup"><span data-stu-id="be649-309">For example, the `Projects` entity set has a default limit of 100 projects per query, but the `Risks` entity set has a default limit of 200.</span></span> <span data-ttu-id="be649-310">For a production installation, the code in the **HelloProjectOData** example should be modified to enable queries of more than 100 projects.</span><span class="sxs-lookup"><span data-stu-id="be649-310">For a production installation, the code in the **HelloProjectOData** example should be modified to enable queries of more than 100 projects.</span></span> <span data-ttu-id="be649-311">For more information, see [Next steps](#next-steps) and [Querying OData feeds for Project reporting data](/previous-versions/office/project-odata/jj163048(v=office.15)).</span><span class="sxs-lookup"><span data-stu-id="be649-311">For more information, see [Next steps](#next-steps) and [Querying OData feeds for Project reporting data](/previous-versions/office/project-odata/jj163048(v=office.15)).</span></span>

## <a name="example-code-for-the-helloprojectodata-add-in"></a><span data-ttu-id="be649-312">HelloProjectOData アドインのコード例</span><span class="sxs-lookup"><span data-stu-id="be649-312">Example code for the HelloProjectOData add-in</span></span>

### <a name="helloprojectodatahtml-file"></a><span data-ttu-id="be649-313">HelloProjectOData.html ファイル</span><span class="sxs-lookup"><span data-stu-id="be649-313">HelloProjectOData.html file</span></span>

<span data-ttu-id="be649-314">次のコードは、**HelloProjectODataWeb** プロジェクトの `Pages\HelloProjectOData.html` ファイルに収められています。</span><span class="sxs-lookup"><span data-stu-id="be649-314">The following code is in the `Pages\HelloProjectOData.html` file of the **HelloProjectODataWeb** project.</span></span>

```HTML
<!DOCTYPE html>
<html>
    <head>
        <meta charset="UTF-8" />
        <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
        <title>Test ProjectData Service</title>

        <link rel="stylesheet" type="text/css" href="../Content/Office.css" />

        <!-- Add your CSS styles to the following file. -->
        <link rel="stylesheet" type="text/css" href="../Content/App.css" />

        <!-- Use the CDN reference to the mini-version of jQuery when deploying your add-in. -->
        <!--<script src="http://ajax.aspnetcdn.com/ajax/jquery/jquery-1.9.0.min.js"></script> -->
        <script src="../Scripts/jquery-1.7.1.js"></script>

        <!-- Use the CDN reference to Office.js when deploying your add-in -->
        <!--<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>-->

        <!-- Use the local script references for Office.js to enable offline debugging -->
        <script src="../Scripts/Office/1.0/MicrosoftAjax.js"></script>
        <script src="../Scripts/Office/1.0/Office.js"></script>

        <!-- Add your JavaScript to the following files. -->
        <script src="../Scripts/HelloProjectOData.js"></script>
        <script src="../Scripts/SurfaceErrors.js"></script>
    </head>
    <body>
        <div id="SectionContent">
        <div id="odataQueries">
            ODATA REST QUERY
        </div>
        <div id="odataInfo">
            <button class="button-wide" onclick="setOdataUrl()">Get ProjectData Endpoint</button>
            <br />
            <br />
            <span class="rest" id="projectDataEndPoint">Endpoint of the 
            <strong>ProjectData</strong> service</span>
            <br />
        </div>
        <div id="compareProjectData">
            <button class="button-wide" disabled="disabled" id="compareProjects"
            onclick="retrieveOData()">
            Compare All Projects</button>
            <br />
        </div>
        </div>
        <div id="corpInfo">
        <table class="infoTable" aria-readonly="True" style="width: 100%;">
            <tr>
            <td class="heading_leftCol"></td>
            <td class="heading_midCol"><strong>Average</strong></td>
            <td class="heading_rightCol"><strong>Current</strong></td>
            </tr>
            <tr>
            <td class="row_leftCol"><strong>Project Cost</strong></td>
            <td class="row_midCol" id="AverageProjectCost">&amp;nbsp;</td>
            <td class="row_rightCol" id="CurrentProjectCost">&amp;nbsp;</td>
            </tr>
            <tr>
            <td class="row_leftCol"><strong>Project Actual Cost</strong></td>
            <td class="row_midCol" id="AverageProjectActualCost">&amp;nbsp;</td>
            <td class="row_rightCol" id="CurrentProjectActualCost">&amp;nbsp;</td>
            </tr>
            <tr>
            <td class="row_leftCol"><strong>Project Work</strong></td>
            <td class="row_midCol" id="AverageProjectWork">&amp;nbsp;</td>
            <td class="row_rightCol" id="CurrentProjectWork">&amp;nbsp;</td>
            </tr>
            <tr>
            <td class="row_leftCol"><strong>Project % Complete</strong></td>
            <td class="row_midCol" id="AverageProjectPercentComplete">&amp;nbsp;</td>
            <td class="row_rightCol" id="CurrentProjectPercentComplete">&amp;nbsp;</td>
            </tr>
        </table>
        </div>
        <img alt="Corporation" class="logo" src="../../images/NewLogo.png" />
        <br />
        <textarea id="odataText" rows="12" cols="40"></textarea>
    </body>
</html>
```

### <a name="helloprojectodatajs-file"></a><span data-ttu-id="be649-315">HelloProjectOData.js ファイル</span><span class="sxs-lookup"><span data-stu-id="be649-315">HelloProjectOData.js file</span></span>

<span data-ttu-id="be649-316">次のコードは、**HelloProjectODataWeb** プロジェクトの `Scripts\Office\HelloProjectOData.js` ファイルに収められています。</span><span class="sxs-lookup"><span data-stu-id="be649-316">The following code is in the `Scripts\Office\HelloProjectOData.js` file of the **HelloProjectODataWeb** project.</span></span>

```js
/* File: HelloProjectOData.js
* JavaScript functions for the HelloProjectOData example task pane app.
* October 2, 2012
*/

var PROJDATA = "/_api/ProjectData";
var PROJQUERY = "/Projects?";
var QUERY_FILTER = "$filter=ProjectName ne 'Timesheet Administrative Work Items'";
var QUERY_SELECT1 = "&amp;$select=ProjectId, ProjectName";
var QUERY_SELECT2 = ", ProjectCost, ProjectWork, ProjectPercentCompleted, ProjectActualCost";
var _pwa;           // URL of Project Web App.
var _projectUid;    // GUID of the active project.
var _docUrl;        // Path of the project document.
var _odataUrl = ""; // URL of the OData service: http[s]://ServerName /ProjectServerName /_api/ProjectData

// The initialize function is required for all add-ins.
Office.initialize = function (reason) {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
    });
}

// Set the global variables, enable the Compare All Projects button,
// and display the URL of the ProjectData service.
// Display an error if Project is not connected with Project Web App.
function setOdataUrl() {
    Office.context.document.getProjectFieldAsync(
        Office.ProjectProjectFields.ProjectServerUrl,
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                _pwa = String(asyncResult.value.fieldValue);

                // If you debug with Visual Studio on a local Project Server computer,
                // uncomment the following lines to use the localhost URL.
                //var localhost = location.host.split(":", 1);
                //var pwaStartPosition = _pwa.lastIndexOf("/");
                //var pwaLength = _pwa.length - pwaStartPosition;
                //var pwaName = _pwa.substr(pwaStartPosition, pwaLength);
                //_pwa = location.protocol + "//" + localhost + pwaName;

                if (_pwa.substring(0, 4) == "http") {
                    _odataUrl = _pwa + PROJDATA;
                    $("#compareProjects").removeAttr("disabled");
                    getProjectGuid();
                }
                else {
                    _odataUrl = "No connection!";
                    throwError(_odataUrl, "You are not connected to Project Web App.");
                }
                getDocumentUrl();
                $("#projectDataEndPoint").text(_odataUrl);
            }
            else {
                throwError(asyncResult.error.name, asyncResult.error.message);
            }
        }
    );
}

// Get the GUID of the active project.
function getProjectGuid() {
    Office.context.document.getProjectFieldAsync(
        Office.ProjectProjectFields.GUID,
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                _projectUid = asyncResult.value.fieldValue;
            }
            else {
                throwError(asyncResult.error.name, asyncResult.error.message);
            }
        }
    );
}

// Get the path of the project in Project web app, which is in the form <>\ProjectName .
function getDocumentUrl() {
    _docUrl = "Document path:\r\n" + Office.context.document.url;
}

//  Functions to get and parse the Project Server reporting data./

// Get data about all projects on Project Server,
// by using a REST query with the ajax method in jQuery.
function retrieveOData() {
    var restUrl = _odataUrl + PROJQUERY + QUERY_FILTER + QUERY_SELECT1 + QUERY_SELECT2;
    var accept = "application/json; odata=verbose";
    accept.toLocaleLowerCase();

    // Enable cross-origin scripting (required by jQuery 1.5 and later).
    // This does not work with Project on the web.
    $.support.cors = true;

    $.ajax({
        url: restUrl,
        type: "GET",
        contentType: "application/json",
        data: "",      // Empty string for the optional data.
        //headers: { "Accept": accept },
        beforeSend: function (xhr) {
            xhr.setRequestHeader("ACCEPT", accept);
        },
        complete: function (xhr, textStatus) {
            // Create a message to display in the text box.
            var message = "\r\ntextStatus: " + textStatus +
                "\r\nContentType: " + xhr.getResponseHeader("Content-Type") +
                "\r\nStatus: " + xhr.status +
                "\r\nResponseText:\r\n" + xhr.responseText;

            // xhr.responseText is the result from an XmlHttpRequest, which 
            // contains the JSON response from the OData service.
            parseODataResult(xhr.responseText, _projectUid);

            // Write the document name, response header, status, and JSON to the odataText control.
            $("#odataText").text(_docUrl);
            $("#odataText").append("\r\nREST query:\r\n" + restUrl);
            $("#odataText").append(message);

            if (xhr.status != 200 &amp;&amp; xhr.status != 1223 &amp;&amp; xhr.status != 201) {
                $("#odataInfo").append("<div>" + htmlEncode(restUrl) + "</div>");
            }
        },
        error: getProjectDataErrorHandler
    });
}

function getProjectDataErrorHandler(data, errorCode, errorMessage) {
    $("#odataText").text("Error code: " + errorCode + "\r\nError message: \r\n"
        + errorMessage);
    throwError(errorCode, errorMessage);
}

// Calculate the average values of actual cost, cost, work, and percent complete
// for all projects, and compare with the values for the current project.
function parseODataResult(oDataResult, currentProjectGuid) {
    // Deserialize the JSON string into a JavaScript object.
    var res = Sys.Serialization.JavaScriptSerializer.deserialize(oDataResult);
    var len = res.d.results.length;
    var projActualCost = 0;
    var projCost = 0;
    var projWork = 0;
    var projPercentCompleted = 0;
    var myProjectIndex = -1;

    for (i = 0; i < len; i++) {
        // If the current project GUID matches the GUID from the OData query,  
        // then store the project index.
        if (currentProjectGuid.toLocaleLowerCase() == res.d.results[i].ProjectId) {
            myProjectIndex = i;
        }
        projCost += Number(res.d.results[i].ProjectCost);
        projWork += Number(res.d.results[i].ProjectWork);
        projActualCost += Number(res.d.results[i].ProjectActualCost);
        projPercentCompleted += Number(res.d.results[i].ProjectPercentCompleted);

    }
    var avgProjCost = projCost / len;
    var avgProjWork = projWork / len;
    var avgProjActualCost = projActualCost / len;
    var avgProjPercentCompleted = projPercentCompleted / len;

    // Round off cost to two decimal places, and round off other values to one decimal place.
    avgProjCost = avgProjCost.toFixed(2);
    avgProjWork = avgProjWork.toFixed(1);
    avgProjActualCost = avgProjActualCost.toFixed(2);
    avgProjPercentCompleted = avgProjPercentCompleted.toFixed(1);

    // Display averages in the table, with the correct units. 
    document.getElementById("AverageProjectCost").innerHTML = "$"
        + avgProjCost;
    document.getElementById("AverageProjectActualCost").innerHTML
        = "$" + avgProjActualCost;
    document.getElementById("AverageProjectWork").innerHTML
        = avgProjWork + " hrs";
    document.getElementById("AverageProjectPercentComplete").innerHTML
        = avgProjPercentCompleted + "%";

    // Calculate and display values for the current project.
    if (myProjectIndex != -1) {

        var myProjCost = Number(res.d.results[myProjectIndex].ProjectCost);
        var myProjWork = Number(res.d.results[myProjectIndex].ProjectWork);
        var myProjActualCost = Number(res.d.results[myProjectIndex].ProjectActualCost);
        var myProjPercentCompleted = Number(res.d.results[myProjectIndex].ProjectPercentCompleted);

        myProjCost = myProjCost.toFixed(2);
        myProjWork = myProjWork.toFixed(1);
        myProjActualCost = myProjActualCost.toFixed(2);
        myProjPercentCompleted = myProjPercentCompleted.toFixed(1);

        document.getElementById("CurrentProjectCost").innerHTML = "$" + myProjCost;

        if (Number(myProjCost) <= Number(avgProjCost)) {
            document.getElementById("CurrentProjectCost").style.color = "green"
        }
        else {
            document.getElementById("CurrentProjectCost").style.color = "red"
        }

        document.getElementById("CurrentProjectActualCost").innerHTML = "$" + myProjActualCost;

        if (Number(myProjActualCost) <= Number(avgProjActualCost)) {
            document.getElementById("CurrentProjectActualCost").style.color = "green"
        }
        else {
            document.getElementById("CurrentProjectActualCost").style.color = "red"
        }

        document.getElementById("CurrentProjectWork").innerHTML = myProjWork + " hrs";

        if (Number(myProjWork) <= Number(avgProjWork)) {
            document.getElementById("CurrentProjectWork").style.color = "red"
        }
        else {
            document.getElementById("CurrentProjectWork").style.color = "green"
        }

        document.getElementById("CurrentProjectPercentComplete").innerHTML = myProjPercentCompleted + "%";

        if (Number(myProjPercentCompleted) <= Number(avgProjPercentCompleted)) {
            document.getElementById("CurrentProjectPercentComplete").style.color = "red"
        }
        else {
            document.getElementById("CurrentProjectPercentComplete").style.color = "green"
        }
    }
    else {    // The current project is not published.
        document.getElementById("CurrentProjectCost").innerHTML = "NA";
        document.getElementById("CurrentProjectCost").style.color = "blue"

        document.getElementById("CurrentProjectActualCost").innerHTML = "NA";
        document.getElementById("CurrentProjectActualCost").style.color = "blue"

        document.getElementById("CurrentProjectWork").innerHTML = "NA";
        document.getElementById("CurrentProjectWork").style.color = "blue"

        document.getElementById("CurrentProjectPercentComplete").innerHTML = "NA";
        document.getElementById("CurrentProjectPercentComplete").style.color = "blue"
    }
}
```

### <a name="appcss-file"></a><span data-ttu-id="be649-317">App.css ファイル</span><span class="sxs-lookup"><span data-stu-id="be649-317">App.css file</span></span>

<span data-ttu-id="be649-318">次のコードは、**HelloProjectODataWeb** プロジェクトの `Content\App.css` ファイルに収められています。</span><span class="sxs-lookup"><span data-stu-id="be649-318">The following code is in the `Content\App.css` file of the **HelloProjectODataWeb** project.</span></span>

```css
/*
*  File: App.css for the HelloProjectOData app.
*  Updated: 10/2/2012
*/

body
{
    font-size: 11pt;
}
h1
{
    font-size: 22pt;
}
h2
{
    font-size: 16pt;
}

/******************************************************************
Code label class
******************************************************************/

.rest 
{
    font-family: 'Courier New';
    font-size: 0.9em;
}

/******************************************************************
Button classes
******************************************************************/

.button-wide {
    width: 210px;
    margin-top: 2px;
}
.button-narrow 
{
    width: 80px;
    margin-top: 2px;
}

/******************************************************************
Table styles
******************************************************************/

.infoTable
{
    text-align: center; 
    vertical-align: middle
}
.heading_leftCol
{
    width: 20px;
    height: 20px;
}
.heading_midCol
{
    width: 100px;
    height: 20px;
    font-size: medium; 
    font-weight: bold; 
}
.heading_rightCol
{
    width: 101px;
    height: 20px;
    font-size: medium;
    font-weight: bold;
}
.row_leftCol
{
    width: 20px;
    font-size: small;
    font-weight: bold;
}
.row_midCol
{
    width: 100px;
}
.row_rightCol
{
    width: 101px;
}
.logo
{
    width: 135px;
    height: 53px;
}
```

### <a name="surfaceerrorsjs-file"></a><span data-ttu-id="be649-319">SurfaceErrors.js ファイル</span><span class="sxs-lookup"><span data-stu-id="be649-319">SurfaceErrors.js file</span></span>

<span data-ttu-id="be649-320">SurfaceErrors.js ファイルのコードは、「[テキスト エディターを使用して Project 2013 用の作業ウィンドウ アドインを初めて作成する](../project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md)」の「_堅牢なプログラミング_」セクションからコピーできます。</span><span class="sxs-lookup"><span data-stu-id="be649-320">You can copy code for the SurfaceErrors.js file from the _Robust Programming_ section of [Create your first task pane add-in for Project 2013 by using a text editor](../project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md).</span></span>

## <a name="next-steps"></a><span data-ttu-id="be649-321">次の手順</span><span class="sxs-lookup"><span data-stu-id="be649-321">Next steps</span></span>

<span data-ttu-id="be649-322">**HelloProjectOData** が AppSource で販売される実稼働アドインだった場合や、SharePoint アプリ カタログで配布する場合は、異なる方法で設計されます。</span><span class="sxs-lookup"><span data-stu-id="be649-322">If **HelloProjectOData** were a production add-in to be sold in AppSource or distributed in a SharePoint app catalog, it would be designed differently.</span></span> <span data-ttu-id="be649-323">たとえば、テキスト ボックスのデバッグ出力や、**ProjectData** エンドポイントを取得するためのボタンはおそらくありません。</span><span class="sxs-lookup"><span data-stu-id="be649-323">For example, there would be no debug output in a text box, and probably no button to get the **ProjectData** endpoint.</span></span> <span data-ttu-id="be649-324">また、100 を超えるプロジェクトを持つProject Web Appを処理するために関数を書き `retireveOData` 換える必要があります。</span><span class="sxs-lookup"><span data-stu-id="be649-324">You would also have to rewrite the `retireveOData` function to handle Project Web App instances that have more than 100 projects.</span></span>

<span data-ttu-id="be649-p164">このアドインには、追加のエラー チェックと、エッジ ケースをキャッチして説明または表示するためのロジックを組み込む必要があります。たとえば、Project Web App インスタンスに、平均期間が 5 日で平均コストが $2400 になる 1000 個のプロジェクトがあって、期間が 20 日より長いのはアクティブ プロジェクトだけだとすると、コストと作業の比較は歪んだものになるでしょう。それは頻度グラフで示すことができます。期間を表示したり、同じような長さのプロジェクトを比較したり、同じ部門または異なる部門のプロジェクトを比較したりするオプションを追加するとよいでしょう。あるいは、表示するフィールドのリストからユーザーが選択できるような方法を追加することもできます。</span><span class="sxs-lookup"><span data-stu-id="be649-p164">The add-in should contain additional error checks, plus logic to catch and explain or show edge cases. For example, if a Project Web App instance has 1000 projects with an average duration of five days and average cost of $2400, and the active project is the only one that has a duration longer than 20 days, the cost and work comparison would be skewed. That could be shown with a frequency graph. You could add options to display duration, compare similar length projects, or compare projects from the same or different departments. Or, add a way for the user to select from a list of fields to display.</span></span>

<span data-ttu-id="be649-330">**ProjectData** サービスの他のクエリでは、クエリ文字列の長さに制限があります。これは、親コレクションから子コレクション内のオブジェクトに対してクエリが実行できる手順の数に影響します。</span><span class="sxs-lookup"><span data-stu-id="be649-330">For other queries of the **ProjectData** service, there are limits to the length of the query string, which affects the number of steps that a query can take from a parent collection to an object in a child collection.</span></span> <span data-ttu-id="be649-331">たとえば、**Projects** から **Tasks** へ、そしてタスク アイテムへという 2 ステップのクエリはうまく動作しますが、**Projects** から、**Tasks**、**Assignments** を経て、割り当てアイテムへという 3 ステップのクエリになると、URL の既定の最大長を超える可能性があります。</span><span class="sxs-lookup"><span data-stu-id="be649-331">For example, a two-step query of **Projects** to **Tasks** to task item works, but a three-step query such as **Projects** to **Tasks** to **Assignments** to assignment item may exceed the default maximum URL length.</span></span> <span data-ttu-id="be649-332">詳しくは、「[Project レポート データの OData フィードにクエリを実行する](/previous-versions/office/project-odata/jj163048(v=office.15))」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="be649-332">For more information, see [Querying OData feeds for Project reporting data](/previous-versions/office/project-odata/jj163048(v=office.15)).</span></span>

<span data-ttu-id="be649-333">実稼働環境で **使用するために HelloProjectOData** アドインを変更する場合は、次の手順を実行します。</span><span class="sxs-lookup"><span data-stu-id="be649-333">If you modify the **HelloProjectOData** add-in for production use, do the following steps.</span></span>

- <span data-ttu-id="be649-334">HelloProjectOData.html ファイルで、パフォーマンスを向上させるために、office.js の参照をローカル プロジェクトから CDN の参照に変更します。</span><span class="sxs-lookup"><span data-stu-id="be649-334">In the HelloProjectOData.html file, for better performance, change the office.js reference from the local project to the CDN reference:</span></span>

    ```HTML
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
    ```

- <span data-ttu-id="be649-335">関数を `retrieveOData` 書き換え、100 を超えるプロジェクトのクエリを有効にする。</span><span class="sxs-lookup"><span data-stu-id="be649-335">Rewrite the `retrieveOData` function to enable queries of more than 100 projects.</span></span> <span data-ttu-id="be649-336">たとえば、`~/ProjectData/Projects()/$count` クエリでプロジェクトの数を取得し、プロジェクト データの REST クエリで _$skip_ 操作と _$top_ 操作を使います。</span><span class="sxs-lookup"><span data-stu-id="be649-336">For example, you could get the number of projects with a `~/ProjectData/Projects()/$count` query, and use the _$skip_ operator and _$top_ operator in the REST query for project data.</span></span> <span data-ttu-id="be649-337">ループの中で複数のクエリを実行し、各クエリからのデータを平均化します。</span><span class="sxs-lookup"><span data-stu-id="be649-337">Run multiple queries in a loop, and then average the data from each query.</span></span> <span data-ttu-id="be649-338">プロジェクト データの各クエリは、次の形式です。</span><span class="sxs-lookup"><span data-stu-id="be649-338">Each query for project data would be of the form:</span></span> 

  `~/ProjectData/Projects()?skip= [numSkipped]&amp;$top=100&amp;$filter=[filter]&amp;$select=[field1,field2, ???????]`

  <span data-ttu-id="be649-p167">For more information, see [OData System Query Options Using the REST Endpoint](/previous-versions/dynamicscrm-2015/developers-guide/gg309461(v=crm.7)). You can also use the [Set-SPProjectOdataConfiguration](/powershell/module/sharepoint-server/Set-SPProjectOdataConfiguration?view=sharepoint-ps&preserve-view=true) command in Windows PowerShell to override the default page size for a query of the **Projects** entity set (or any of the 33 entity sets). See [ProjectData - Project OData service reference](/previous-versions/office/project-odata/jj163015(v=office.15)).</span><span class="sxs-lookup"><span data-stu-id="be649-p167">For more information, see [OData System Query Options Using the REST Endpoint](/previous-versions/dynamicscrm-2015/developers-guide/gg309461(v=crm.7)). You can also use the [Set-SPProjectOdataConfiguration](/powershell/module/sharepoint-server/Set-SPProjectOdataConfiguration?view=sharepoint-ps&preserve-view=true) command in Windows PowerShell to override the default page size for a query of the **Projects** entity set (or any of the 33 entity sets). See [ProjectData - Project OData service reference](/previous-versions/office/project-odata/jj163015(v=office.15)).</span></span>

- <span data-ttu-id="be649-342">アドインを展開するには、「[Office アドインを発行する](../publish/publish.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="be649-342">To deploy the add-in, see [Publish your Office Add-in](../publish/publish.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="be649-343">関連項目</span><span class="sxs-lookup"><span data-stu-id="be649-343">See also</span></span>

- [<span data-ttu-id="be649-344">Project 用の作業ウィンドウ アドイン</span><span class="sxs-lookup"><span data-stu-id="be649-344">Task pane add-ins for Project</span></span>](project-add-ins.md)
- [<span data-ttu-id="be649-345">テキスト エディターを使用して Project 2013 用の作業ウィンドウ アドインを初めて作成する</span><span class="sxs-lookup"><span data-stu-id="be649-345">Create your first task pane add-in for Project 2013 by using a text editor</span></span>](create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md)
- <span data-ttu-id="be649-346">[ProjectData - Project OData サービス リファレンス](/previous-versions/office/project-odata/jj163015(v=office.15))</span><span class="sxs-lookup"><span data-stu-id="be649-346">[ProjectData - Project OData service reference](/previous-versions/office/project-odata/jj163015(v=office.15))</span></span>
- [<span data-ttu-id="be649-347">Office アドインの XML マニフェスト</span><span class="sxs-lookup"><span data-stu-id="be649-347">Office Add-ins XML manifest</span></span>](../develop/add-in-manifests.md)
- [<span data-ttu-id="be649-348">Office アドインを発行する</span><span class="sxs-lookup"><span data-stu-id="be649-348">Publish your Office Add-in</span></span>](../publish/publish.md)
