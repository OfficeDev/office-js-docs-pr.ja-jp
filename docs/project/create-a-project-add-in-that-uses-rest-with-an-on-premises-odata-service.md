---
title: 社内の Project Server OData サービスで REST を使用する Project アドインを作成する
description: ''
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: 1e50d90b844e78620866e94e44377c903b169783
ms.sourcegitcommit: 3f84b2caa73d7fe1eb0d15e32ea4dec459e2ff53
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/12/2019
ms.locfileid: "34910358"
---
# <a name="create-a-project-add-in-that-uses-rest-with-an-on-premises-project-server-odata-service"></a><span data-ttu-id="1e277-102">社内の Project Server OData サービスで REST を使用する Project アドインを作成する</span><span class="sxs-lookup"><span data-stu-id="1e277-102">Create a Project add-in that uses REST with an on-premises Project Server OData service</span></span>

<span data-ttu-id="1e277-p101">この記事では、アクティブなプロジェクトのコストと作業のデータを現在の Project Web App インスタンスの全プロジェクトの平均と比較する Project Professional 2013 用の作業ウィンドウ アドインを作成する方法を説明します。このアドインは、REST と jQuery ライブラリを使って、Project Server 2013 の **ProjectData** OData レポート サービスにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="1e277-p101">This article describes how to build a task pane add-in for Project Professional 2013 that compares cost and work data in the active project with the averages for all projects in the current Project Web App instance. The add-in uses REST with the jQuery library to access the  **ProjectData** OData reporting service in Project Server 2013.</span></span>

<span data-ttu-id="1e277-105">この記事のコードは、Microsoft Corporation の Saurabh Sanghvi と Arvind Iyer が開発したサンプルに基づいています。</span><span class="sxs-lookup"><span data-stu-id="1e277-105">The code in this article is based on a sample developed by Saurabh Sanghvi and Arvind Iyer, Microsoft Corporation.</span></span>

## <a name="prerequisites-for-creating-a-task-pane-add-in-that-reads-project-server-reporting-data"></a><span data-ttu-id="1e277-106">Project Server のレポート データを読み取る作業ウィンドウ アドインを作成するための前提条件</span><span class="sxs-lookup"><span data-stu-id="1e277-106">Prerequisites for creating a task pane add-in that reads Project Server reporting data</span></span>

<span data-ttu-id="1e277-107">Project Server 2013 の社内インストールにおける Project Web App インスタンスの **ProjectData** サービスを読み取る Project 作業ウィンドウ アドインを作成するための前提条件を以下に示します。</span><span class="sxs-lookup"><span data-stu-id="1e277-107">The following are the prerequisites for creating a Project task pane add-in that reads the  **ProjectData** service of a Project Web App instance in an on-premises installation of Project Server 2013:</span></span>

- <span data-ttu-id="1e277-p102">使用するローカルの開発用コンピューターに最新のサービス パックと Windows 更新プログラムをインストールしてあることを確認します。オペレーティング システムは、Windows 7、Windows 8、Windows Server 2008、Windows Server 2012 のいずれでもかまいません。</span><span class="sxs-lookup"><span data-stu-id="1e277-p102">Ensure that you have installed the most recent service packs and Windows updates on your local development computer. The operating system can be Windows 7, Windows 8, Windows Server 2008, or Windows Server 2012.</span></span>

- <span data-ttu-id="1e277-p103">Project Web App との接続には Project Professional 2013が必要です。Visual Studio での  **F5** デバッグを有効にするには、開発用コンピューターに Project Professional 2013がインストールされている必要があります。</span><span class="sxs-lookup"><span data-stu-id="1e277-p103">Project Professional 2013 is required to connect with Project Web App. The development computer must have Project Professional 2013 installed to enable  **F5** debugging with Visual Studio.</span></span>

    > [!NOTE]
    > <span data-ttu-id="1e277-112">Project Standard 2013 でも作業ウィンドウ アドインをホストできますが、Project Web App にはログオンできません。</span><span class="sxs-lookup"><span data-stu-id="1e277-112">Project Standard 2013 can also host task pane add-ins, but cannot log on to Project Web App.</span></span>

- <span data-ttu-id="1e277-113">Office Developer Tools for Visual Studio を備えた Visual Studio 2015 には、Office アドインと SharePoint アドインの作成用のテンプレートが含まれています。最新バージョンの Office Developer Tools がインストールされていることを確認してください。 _Office アドインと SharePoint のダウンロード_の「 [ツール](https://developer.microsoft.com/office/docs) 」セクションを参照してください。</span><span class="sxs-lookup"><span data-stu-id="1e277-113">Visual Studio 2015 with Office Developer Tools for Visual Studio includes templates for creating Office and SharePoint Add-ins. Ensure that you have installed the most recent version of Office Developer Tools; see the  _Tools_ section of the [Office Add-ins and SharePoint downloads](https://developer.microsoft.com/office/docs).</span></span>

- <span data-ttu-id="1e277-p104">この記事の手順とコード例では、ローカル ドメインの Project Server 2013の  **ProjectData** サービスにアクセスします。この記事の jQuery メソッドは Project Online には対応していません。</span><span class="sxs-lookup"><span data-stu-id="1e277-p104">The procedures and code examples in this article access the  **ProjectData** service of Project Server 2013 in a local domain. The jQuery methods in this article do not work with Project Online.</span></span>

    <span data-ttu-id="1e277-116">開発用コンピューターから **ProjectData** サービスにアクセスできることを確認します。</span><span class="sxs-lookup"><span data-stu-id="1e277-116">Verify that the  **ProjectData** service is accessible from your development computer.</span></span>

### <a name="procedure-1-to-verify-that-the-projectdata-service-is-accessible"></a><span data-ttu-id="1e277-p105">手順 1. ProjectData サービスにアクセスできることを確認するには</span><span class="sxs-lookup"><span data-stu-id="1e277-p105">Procedure 1. To verify that the ProjectData service is accessible</span></span>

1. <span data-ttu-id="1e277-p106">ブラウザーで REST クエリからの XML データの直接表示を可能にするには、フィードの読み取りビューをオフにします。Internet Explorer でこれを行う方法については、「 [Project Server 2013 レポート データの OData フィードにクエリを実行する](/previous-versions/office/project-odata/jj163048(v=office.15))」の手順 1. のステップ 4. を参照してください。</span><span class="sxs-lookup"><span data-stu-id="1e277-p106">To enable your browser to directly show the XML data from a REST query, turn off the feed reading view. For information about how to do this in Internet Explorer, see Procedure 1, step 4 in [Querying OData feeds for Project reporting data](/previous-versions/office/project-odata/jj163048(v=office.15)).</span></span>

2. <span data-ttu-id="1e277-121">**ProjectData** サービスにクエリを実行するには、お使いのブラウザーで次の URL にアクセスしてください: **http://ServerName /ProjectServerName /_api/ProjectData**。</span><span class="sxs-lookup"><span data-stu-id="1e277-121">Query the  **ProjectData** service by using your browser with the following URL: **http://ServerName /ProjectServerName /_api/ProjectData**.</span></span> <span data-ttu-id="1e277-122">たとえば、Project Web App インスタンスが `http://MyServer/pwa` である場合、ブラウザーは次の結果を示します。</span><span class="sxs-lookup"><span data-stu-id="1e277-122">For example, if the Project Web App instance is  `http://MyServer/pwa`, the browser shows the following results:</span></span>

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

3. <span data-ttu-id="1e277-p108">結果を確認するためにネットワーク資格情報の入力が必要になることもあります。ブラウザーでアクセスが拒否されことを示すメッセージ (エラー 403) が表示された場合は、その Project Web App インスタンスに対するログオンのアクセス許可を与えられていないか、管理者のサポートを必要とするネットワーク上の問題が発生しています。</span><span class="sxs-lookup"><span data-stu-id="1e277-p108">You may have to provide your network credentials to see the results. If the browser shows "Error 403, Access Denied," either you do not have logon permission for that Project Web App instance, or there is a network problem that requires administrative help.</span></span>

## <a name="using-visual-studio-to-create-a-task-pane-add-in-for-project"></a><span data-ttu-id="1e277-125">Visual Studio を使用して Project 用の作業ウィンドウ アドインを作成する</span><span class="sxs-lookup"><span data-stu-id="1e277-125">Using Visual Studio to create a task pane add-in for Project</span></span>

<span data-ttu-id="1e277-p109">Office Developer Tools for Visual Studio には、Project 2013 用の作業ウィンドウ アドインのためのテンプレートが含まれています。 **HelloProjectOData** という名前のソリューションを作成すると、そのソリューションには次の 2 つの Visual Studio プロジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="1e277-p109">Office Developer Tools for Visual Studio includes a template for task pane add-ins for Project 2013. If you create a solution named  **HelloProjectOData**, the solution contains the following two Visual Studio projects:</span></span>

- <span data-ttu-id="1e277-p110">アドイン プロジェクト。ソリューションの名前が付けられます。このプロジェクトにはアドインの XML マニフェスト ファイルが含まれます。アドイン プロジェクトは .NET Framework 4.5 をターゲットにします。手順 3. に、 **HelloProjectOData** アドインのマニフェストを変更するための手順を示します。</span><span class="sxs-lookup"><span data-stu-id="1e277-p110">The add-in project takes the name of the solution. It includes the XML manifest file for the add-in and targets the .NET Framework 4.5. Procedure 3 shows the steps to modify the manifest for the  **HelloProjectOData** add-in.</span></span>

- <span data-ttu-id="1e277-p111">Web プロジェクト。 **HelloProjectODataWeb** という名前が付けられます。このプロジェクトには、作業ウィンドウ内の Web コンテンツのための Web ページ、JavaScript ファイル、CSS ファイル、イメージ、参照、および構成ファイルが含まれます。Web プロジェクトは .NET Framework 4 をターゲットにします。手順 4. および手順 5. に、Web プロジェクト内のファイルを変更して、 **HelloProjectOData** アドインの機能を作成する方法を示します。</span><span class="sxs-lookup"><span data-stu-id="1e277-p111">The web project is named  **HelloProjectODataWeb**. It includes the webpages, JavaScript files, CSS files, images, references, and configuration files for the web content in the task pane. The web project targets the .NET Framework 4. Procedure 4 and Procedure 5 show how to modify the files in the web project to create the functionality of the  **HelloProjectOData** add-in.</span></span>

### <a name="procedure-2-to-create-the-helloprojectodata-add-in-for-project"></a><span data-ttu-id="1e277-p112">手順 2. Project 用の HelloProjectOData アドインを作成するには</span><span class="sxs-lookup"><span data-stu-id="1e277-p112">Procedure 2. To create the HelloProjectOData add-in for Project</span></span>

1. <span data-ttu-id="1e277-137">管理者として Visual Studio 2015 を実行し、スタート ページで **[新しいプロジェクト]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="1e277-137">Run Visual Studio 2015 as an administrator, and then select  **New Project** on the Start page.</span></span>

2. <span data-ttu-id="1e277-p113">**[新しいプロジェクト]** ダイアログ ボックスで **[テンプレート]**、**[Visual C#]**、**[Office/SharePoint]** ノードを展開し、**[Office アドイン]** を選択します。中央のウィンドウの上部にあるターゲット フレームワークのドロップダウン リストで **[.NET Framework 4.5.2]** を選択し、**[Office アドイン]** を選択します (次のスクリーンショットを参照)。</span><span class="sxs-lookup"><span data-stu-id="1e277-p113">In the  **New Project** dialog box, expand the **Templates**,  **Visual C#**, and  **Office/SharePoint** nodes, and then select \*\* Office Add-ins\*\*. Select  **.NET Framework 4.5.2** in the target framework drop-down list at the top of the center pane, and then select **Office Add-in** (see the next screenshot).</span></span>

3. <span data-ttu-id="1e277-140">これらの Visual Studio プロジェクトを両方とも同じディレクトリに配置するには、 [ **ソリューションのディレクトリを作成**] を選択し、目的の場所を参照します。</span><span class="sxs-lookup"><span data-stu-id="1e277-140">To place both of the Visual Studio projects in the same directory, select  **Create directory for solution**, and then browse to the location you want.</span></span>

4. <span data-ttu-id="1e277-141">**[名前]** フィールドに、「HelloProjectOData」と入力して、**[OK]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="1e277-141">In the  **Name** field, typeHelloProjectOData, and then choose  **OK**.</span></span>

    <span data-ttu-id="1e277-142">*図 1. Office アドインの作成*</span><span class="sxs-lookup"><span data-stu-id="1e277-142">*Figure 1. Creating an Office Add-in*</span></span>

    ![Office アドインの作成](../images/pj15-hello-project-o-data-creating-app.png)

5. <span data-ttu-id="1e277-144">**[アドインの種類の選択]** ダイアログ ボックスで、**[作業ウィンドウ]** を選択して、**[次へ]** を選択します (次のスクリーンショットを参照)。</span><span class="sxs-lookup"><span data-stu-id="1e277-144">In the  **Choose the add-in type** dialog box, select **Task pane** and choose **Next** (see the next screenshot).</span></span>

    <span data-ttu-id="1e277-145">*図 2. 作成するアドインの種類の選択*</span><span class="sxs-lookup"><span data-stu-id="1e277-145">*Figure 2. Choosing the type of add-in to create*</span></span>

    ![作成するアドインの種類の選択](../images/pj15-hello-project-o-data-choose-project.png)

6. <span data-ttu-id="1e277-147">[ **ホスト アプリケーションの選択**] ダイアログ ボックスで、[ **Project**] 以外のすべてのチェック ボックスをオフにし (次のスクリーンショットを参照)、 [ **完了**] をクリックします。</span><span class="sxs-lookup"><span data-stu-id="1e277-147">In the  **Choose the host applications** dialog box, clear all check boxes except the **Project** check box (see the next screenshot) and choose **Finish**.</span></span>

    <span data-ttu-id="1e277-148">*図 3. ホスト アプリケーションの選択*</span><span class="sxs-lookup"><span data-stu-id="1e277-148">*Figure 3. Choosing the host application*</span></span>

    ![Project を唯一のホスト アプリケーションとして選択する](../images/create-office-add-in.png)

    <span data-ttu-id="1e277-150">Visual Studio によって、**HelloProjectOdata** プロジェクトと **HelloProjectODataWeb** プロジェクトが作成されます。</span><span class="sxs-lookup"><span data-stu-id="1e277-150">Visual Studio creates the  **HelloProjectOdata** project and the **HelloProjectODataWeb** project.</span></span>

<span data-ttu-id="1e277-151">**アドイン** フォルダーには、カスタムの CSS スタイル用の App.css ファイルが含まれています (次のスクリーンショットを参照してください)。</span><span class="sxs-lookup"><span data-stu-id="1e277-151">The  **AddIn** folder (see the next screenshot) contains the App.css file for custom CSS styles.</span></span> <span data-ttu-id="1e277-152">**ホーム** サブフォルダー内の、Home.html ファイルには、CSS ファイル、アドインを使用している JavaScript ファイル、およびアドインの HTML5 コンテンツへの参照が含まれています。</span><span class="sxs-lookup"><span data-stu-id="1e277-152">In the **Home** subfolder , the Home.html file contains references to the CSS files and the JavaScript files that the add-in uses, and the HTML5 content for the add-in.</span></span> <span data-ttu-id="1e277-153">また、Home.js ファイルは、カスタムの JavaScript コード用です。</span><span class="sxs-lookup"><span data-stu-id="1e277-153">Also, the Home.js file is for your custom JavaScript code.</span></span> <span data-ttu-id="1e277-154">**Scripts** フォルダーには、jQuery ライブラリのファイルが含まれています。</span><span class="sxs-lookup"><span data-stu-id="1e277-154">The **Scripts** folder includes the jQuery library files.</span></span> <span data-ttu-id="1e277-155">**Office** サブフォルダーには、office.js や project-15.js などの JavaScript ライブラリ、および Office アドインでの標準の文字列用の言語ライブラリが含まれています。**コンテンツ** フォルダーで Office.css ファイルには、Office のアドインのすべてに使用する既定のスタイルが含まれます。</span><span class="sxs-lookup"><span data-stu-id="1e277-155">The **Office** subfolder includes the JavaScript libraries such as office.js and project-15.js, plus the language libraries for standard strings in the Office Add-ins. In the **Content** folder, the Office.css file contains the default styles for all of the Office Add-ins.</span></span>

<span data-ttu-id="1e277-156">*図 4. ソリューション エクスプローラーでの既定の Web プロジェクト ファイルの表示*</span><span class="sxs-lookup"><span data-stu-id="1e277-156">*Figure 4. Viewing the default web project files in Solution Explorer*</span></span>

![ソリューション エクスプローラーで Web プロジェクト ファイルを表示する](../images/pj15-hello-project-o-data-initial-solution-explorer.png)

<span data-ttu-id="1e277-p115">**HelloProjectOData** プロジェクトのマニフェストは、HelloProjectOData.xml ファイルです。必要に応じてマニフェストを編集して、アドインの説明、アイコンへの参照、追加言語の情報、その他の設定を追加できます。手順 3 では、アドインの表示名と説明を変更し、アイコンを追加します。</span><span class="sxs-lookup"><span data-stu-id="1e277-p115">The manifest for the  **HelloProjectOData** project is the HelloProjectOData.xml file. You can optionally modify the manifest to add a description of the add-in, a reference to an icon, information for additional languages, and other settings. Procedure 3 simply modifies the add-in display name and description, and adds an icon.</span></span>

<span data-ttu-id="1e277-161">マニフェストについて詳しくは、「[Office アドインの XML マニフェスト](../develop/add-in-manifests.md)」と「[Office アドインのマニフェスト向けのスキーマ リファレンス (v1.1)](../develop/add-in-manifests.md#see-also)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="1e277-161">For more information about the manifest, see [Office Add-ins XML manifest](../develop/add-in-manifests.md) and [Schema reference for Office Add-ins manifests (v1.1)](../develop/add-in-manifests.md#see-also).</span></span>

### <a name="procedure-3-to-modify-the-add-in-manifest"></a><span data-ttu-id="1e277-p116">手順 3. アドインのマニフェストを変更するには</span><span class="sxs-lookup"><span data-stu-id="1e277-p116">Procedure 3. To modify the add-in manifest</span></span>

1. <span data-ttu-id="1e277-164">Visual Studio で、HelloProjectOData.xml ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="1e277-164">In Visual Studio, open the HelloProjectOData.xml file.</span></span>

2. <span data-ttu-id="1e277-p117">既定の表示名は、Visual Studio プロジェクトの名前です ("HelloProjectOData")。たとえば、 **DisplayName** 要素の既定値を"ProjectData 概要" に変更します。</span><span class="sxs-lookup"><span data-stu-id="1e277-p117">The default display name is the name of the Visual Studio project ("HelloProjectOData"). For example, change the default value of the  **DisplayName** element to"Hello ProjectData".</span></span>

3. <span data-ttu-id="1e277-p118">既定の説明も "HelloProjectOData" です。たとえば、Description 要素の既定値を "Test REST queries of the ProjectData service" に変更します。</span><span class="sxs-lookup"><span data-stu-id="1e277-p118">The default description is also "HelloProjectOData". For example, change the default value of the Description element to "Test REST queries of the ProjectData service".</span></span>

4. <span data-ttu-id="1e277-p119">リボンの **[プロジェクト]** タブの **[Office アドイン]** ドロップダウン リストに表示するアイコンを追加します。Visual Studio ソリューションにアイコン ファイルを追加することも、アイコンの URL を使うこともできます。</span><span class="sxs-lookup"><span data-stu-id="1e277-p119">Add an icon to show in the  **Office Add-ins** drop-down list on the **PROJECT** tab of the ribbon. You can add an icon file in the Visual Studio solution or use a URL for an icon.</span></span> 

<span data-ttu-id="1e277-171">以下の手順は、Visual Studio ソリューションにアイコン ファイルを追加するための方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="1e277-171">The following steps show how to add an icon file to the Visual Studio solution:</span></span>

1. <span data-ttu-id="1e277-172">**ソリューション エクスプローラー**で、Images という名前のフォルダーに移動します。</span><span class="sxs-lookup"><span data-stu-id="1e277-172">In  **Solution Explorer**, go to the folder named Images.</span></span>

2. <span data-ttu-id="1e277-p120">**[Office アドイン]** ドロップダウン リストに表示するためには、アイコンのサイズを 32 x 32 ピクセルにする必要があります。たとえば、Project 2013 SDK をインストールしてから、**[Images]** フォルダーを選択し、SDK から次のファイルを追加します。`\Samples\Apps\HelloProjectOData\HelloProjectODataWeb\Images\NewIcon.png`。</span><span class="sxs-lookup"><span data-stu-id="1e277-p120">To be displayed in the  **Office Add-ins** drop-down list, the icon must be 32 x 32 pixels. For example, install the Project 2013 SDK, and then choose the **Images** folder and add the following file from the SDK: `\Samples\Apps\HelloProjectOData\HelloProjectODataWeb\Images\NewIcon.png`</span></span>

    <span data-ttu-id="1e277-175">または、独自の 32 x 32 アイコンを使用するか、NewIcon.png という名前のファイルに次の画像をコピーして、`HelloProjectODataWeb\Images` フォルダーにそのファイルを追加します。</span><span class="sxs-lookup"><span data-stu-id="1e277-175">Alternately, use your own 32 x 32 icon; or, copy the following image to a file named NewIcon.png, and then add that file to the  `HelloProjectODataWeb\Images` folder:</span></span>

    ![HelloProjectOData アプリのアイコン](../images/pj15-hello-project-data-new-icon.jpg)

3. <span data-ttu-id="1e277-p121">HelloProjectOData.xml マニフェストで、**Description** 要素の下に **IconUrl** 要素を追加します。ここで、アイコン URL の値は 32 x 32 アイコン ファイルへの相対パスです。たとえば、次の行を追加します: **<IconUrl DefaultValue="~remoteAppUrl/Images/NewIcon.png" />**。HelloProjectOData.xml マニフェスト ファイルには次が含まれるようになりました (実際の **ID** 値は異なります)。</span><span class="sxs-lookup"><span data-stu-id="1e277-p121">In the HelloProjectOData.xml manifest, add an  **IconUrl** element below the **Description** element, where the value of the icon URL is the relative path to the 32x32 icon file. For example, add the following line: **<IconUrl DefaultValue="~remoteAppUrl/Images/NewIcon.png" />**. The HelloProjectOData.xml manifest file now contains the following (your  **Id** value will be different):</span></span>

    ```XML
    <?xml version="1.0" encoding="UTF-8"?>
    <OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
            xmlns:xsi="https://www.w3.org/2001/XMLSchema-instance" xsi:type="TaskPaneApp">
        <Id>c512df8d-a1c5-4d74-8a34-d30f6bbcbd82 </Id>
        <Version>1.0</Version>
        <ProviderName> [Provider name]</ProviderName>
        <DefaultLocale>en-US</DefaultLocale>
        <DisplayName DefaultValue="Hello ProjectData" />
        <Description DefaultValue="Test REST queries of the ProjectData service"/>
        <IconUrl DefaultValue="~remoteAppUrl/Images/NewIcon.png" />

        <Hosts>
            <Host Name="Project" />
        </Hosts>
        <DefaultSettings>
            <SourceLocation DefaultValue="~remoteAppUrl/AddIn/Home/Home.html" />
        </DefaultSettings>
        <Permissions>ReadWriteDocument</Permissions>
    </OfficeApp>
    ```

## <a name="creating-the-html-content-for-the-helloprojectodata-add-in"></a><span data-ttu-id="1e277-180">HelloProjectOData アドインの HTML コンテンツを作成する</span><span class="sxs-lookup"><span data-stu-id="1e277-180">Creating the HTML content for the HelloProjectOData add-in</span></span>

<span data-ttu-id="1e277-p122">**HelloProjectOData** アドインは、デバッグとエラー出力を含むサンプルです。運用環境で使うことを想定したものではありません。HTML コンテンツのコーディングを始める前に、このアドインの UI とユーザー エクスペリエンスを設計してください。また、HTML コードとやり取りする JavaScript 関数の概略も記述してください。詳しくは、「[Office アドインの設計ガイドライン](../design/add-in-design.md)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="1e277-p122">The  **HelloProjectOData** add-in is a sample that includes debugging and error output; it is not intended for production use. Before you start coding the HTML content, design the UI and user experience for the add-in, and outline the JavaScript functions that interact with the HTML code. For more information, see[Design guidelines for Office Add-ins](../design/add-in-design.md).</span></span> 

<span data-ttu-id="1e277-p123">作業ウィンドウには、上部にアドインの表示名が表示されます。これはマニフェストの  **DisplayName** 要素の値です。HelloProjectOData.html ファイルの **body** 要素には、次のような他の UI 要素が含まれています。</span><span class="sxs-lookup"><span data-stu-id="1e277-p123">The task pane shows the add-in display name at the top, which is the value of the  **DisplayName** element in the manifest. The **body** element in the HelloProjectOData.html file contains the other UI elements, as follows:</span></span>

- <span data-ttu-id="1e277-186">サブタイトルは、 **ODATA REST QUERY** など、操作の一般的な機能または種類を表します。</span><span class="sxs-lookup"><span data-stu-id="1e277-186">A subtitle indicates the general functionality or type of operation, for example,  **ODATA REST QUERY**.</span></span>

- <span data-ttu-id="1e277-p124">[ **ProjectData エンドポイントを取得**] ボタンは、 **setOdataUrl** 関数を呼び出して、 **ProjectData** サービスのエンドポイントを取得し、それをテキスト ボックスに表示します。Project が Project Web App と接続されていない場合、アドインはエラー ハンドラーを呼び出して、ポップアップ エラー メッセージを表示します。</span><span class="sxs-lookup"><span data-stu-id="1e277-p124">The  **Get ProjectData Endpoint** button calls the **setOdataUrl** function to get the endpoint of the **ProjectData** service, and display it in a text box. If Project is not connected with Project Web App, the add-in calls an error handler to display a pop-up error message.</span></span>

- <span data-ttu-id="1e277-p125">[ **すべてのプロジェクトを比較**] ボタンは、アドインが有効な OData エンドポイントを取得するまで無効化されています。このボタンを選択すると、 **retrieveOData** 関数が呼び出されます。この関数は REST クエリを使用して、 **ProjectData** サービスからプロジェクトのコストと作業のデータを取得します。</span><span class="sxs-lookup"><span data-stu-id="1e277-p125">The  **Compare All Projects** button is disabled until the add-in gets a valid OData endpoint. When you select the button, it calls the **retrieveOData** function, which uses a REST query to get project cost and work data from the **ProjectData** service.</span></span>

- <span data-ttu-id="1e277-p126">テーブルには、プロジェクト コスト、実績コスト、作業、および達成率の平均値が表示されます。このテーブルでは、現在アクティブなプロジェクトの値と平均値との比較も行われます。現在の値が全プロジェクトの平均値より大きい場合は、その値が赤で表示されます。現在の値が平均値より小さい場合は、その値が緑で表示されます。現在の値がない場合は、青の  **NA** が表示されます。</span><span class="sxs-lookup"><span data-stu-id="1e277-p126">A table displays the average values for project cost, actual cost, work, and percent complete. The table also compares the current active project values with the average. If the current value is greater than the average for all projects, the value is displayed as red. If the current value is less than the average, the value is displayed as green. If the current value is not available, the table displays a blue  **NA**.</span></span>

    <span data-ttu-id="1e277-196">**retrieveOData** 関数は、テーブルの値を計算して表示する **parseODataResult** 関数を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="1e277-196">The  **retrieveOData** function calls the **parseODataResult** function, which calculates and displays values for the table.</span></span>

    > [!NOTE]
    > <span data-ttu-id="1e277-p127">この例では、アクティブなプロジェクトのコストと作業のデータは発行された値から導出されます。Project で値を変更した場合、**ProjectData** サービスは、そのプロジェクトが発行されるまで変更を反映しません。</span><span class="sxs-lookup"><span data-stu-id="1e277-p127">In this example, cost and work data for the active project are derived from the published values. If you change values in Project, the  **ProjectData** service does not have the changes until the project is published.</span></span>

### <a name="procedure-4-to-create-the-html-content"></a><span data-ttu-id="1e277-p128">手順 4. HTML コンテンツを作成するには</span><span class="sxs-lookup"><span data-stu-id="1e277-p128">Procedure 4. To create the HTML content</span></span>

1. <span data-ttu-id="1e277-p129">Home.html ファイルの  **head** 要素内に、アドインで使用する CSS ファイルの **link** 要素を追加します。Visual Studio プロジェクト テンプレートには、カスタム CSS スタイルに使用できる App.css ファイルのリンクが含まれています。</span><span class="sxs-lookup"><span data-stu-id="1e277-p129">In the  **head** element of the Home.html file, add any additional **link** elements for CSS files that your add-in uses. The Visual Studio project template includes a link for the App.css file that you can use for custom CSS styles.</span></span>

2. <span data-ttu-id="1e277-p130">アドインで使用する JavaScript ライブラリのために、その他の **script** 要素を追加します。プロジェクト テンプレートには、jQuery のリンクが含まれています。**[スクリプト]** フォルダーには、_[バージョン]_.js、office.js、および MicrosoftAjax.js ファイルがあります。</span><span class="sxs-lookup"><span data-stu-id="1e277-p130">Add any additional  **script** elements for JavaScript libraries that your add-in uses. The project template includes links for the jQuery- _[version]_.js, office.js, and MicrosoftAjax.js files in the  **Scripts** folder.</span></span>

    > [!NOTE]
    > <span data-ttu-id="1e277-p131">アドインを展開する前に、office.js の参照と jQuery の参照をコンテンツ配信ネットワーク (CDN) の参照に変更してください。CDN の参照は最新のバージョンと高いパフォーマンスを提供します。</span><span class="sxs-lookup"><span data-stu-id="1e277-p131">Before you deploy the add-in, change the office.js reference and the jQuery reference to the content delivery network (CDN) reference. The CDN reference provides the most recent version and better performance.</span></span>

    <span data-ttu-id="1e277-p132">また、**HelloProjectOData** アドインは、ポップアップ メッセージでエラーを表示する SurfaceErrors.js ファイルを使用します。「[テキスト エディターを使用して Project 2013 用の作業ウィンドウ アドインを初めて作成する](../project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md)」の「_堅牢なプログラミング_」セクションのコードをコピーし、SurfaceErrors.js ファイルを **HelloProjectODataWeb** プロジェクトの **Scripts\Office** フォルダーに追加できます。</span><span class="sxs-lookup"><span data-stu-id="1e277-p132">The  **HelloProjectOData** add-in also uses the SurfaceErrors.js file, which displays errors in a pop-up message. You can copy the code from the _Robust Programming_ section of [Create your first task pane add-in for Project 2013 by using a text editor](../project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md), and then add a SurfaceErrors.js file in the  **Scripts\Office** folder of the **HelloProjectODataWeb** project.</span></span>

    <span data-ttu-id="1e277-209">次に示すのは、**head** 要素の更新された HTML コードと、SurfaceErrors.js ファイルに追加された行です。</span><span class="sxs-lookup"><span data-stu-id="1e277-209">Following is the updated HTML code for the  **head** element, with the additional line for the SurfaceErrors.js file:</span></span>

    ```HTML
    <!DOCTYPE html>
    <html>
    <head>
    <meta charset="UTF-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
    <title>Test ProjectData Service</title>

    <link rel="stylesheet" type="text/css" href="../Content/Office.css" />

    <!-- Add your CSS styles to the following file -->
    <link rel="stylesheet" type="text/css" href="../Content/App.css" />

    <!-- Use the CDN reference to the mini-version of jQuery when deploying your add-in. -->
    <!--<script src="http://ajax.aspnetcdn.com/ajax/jquery/jquery-1.9.0.min.js"></script> -->
    <script src="../Scripts/jquery-1.7.1.js"></script>

    <!-- Use the CDN reference to office.js when deploying your add-in. -->
    <!--<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>-->

    <!-- Use the local script references for Office.js to enable offline debugging -->
    <script src="../Scripts/Office/1.0/MicrosoftAjax.js"></script>
    <script src="../Scripts/Office/1.0/Office.js"></script>

    <!-- Add your JavaScript to the following files -->
    <script src="../Scripts/HelloProjectOData.js"></script>
    <script src="../Scripts/SurfaceErrors.js"></script>
    </head>
    <body>
    <!-- See the code in Step 3. -->
    </body>
    </html>
    ```

3. <span data-ttu-id="1e277-p133">**body** 要素内で、テンプレートから既存のコードを削除し、ユーザー インターフェイスのコードを追加します。要素にデータを入れるか、要素を jQuery ステートメントで操作する場合、その要素に一意の **id** 属性が含まれている必要があります。次のコードでは、jQuery の関数で使用する **button** 要素、**span** 要素、および **td** (テーブル セル定義) 要素の **id** 属性を太字で示しています。</span><span class="sxs-lookup"><span data-stu-id="1e277-p133">In the **body** element, delete the existing code from the template, and then add the code for the user interface. If an element is to be filled with data or manipulated by a jQuery statement, the element must include a unique **id** attribute. In the following code, the **id** attributes for the **button**,  **span**, and  **td** (table cell definition) elements that jQuery functions use are shown in bold font.</span></span>

   <span data-ttu-id="1e277-p134">次の HTML は、グラフィックス イメージ (会社のロゴなど) を追加します。適切なロゴを選択するか、Project 2013 SDK ダウンロードから NewLogo.png ファイルをコピーした後、**ソリューション エクスプローラー**を使用してファイルを `HelloProjectODataWeb\Images` フォルダーに追加できます。</span><span class="sxs-lookup"><span data-stu-id="1e277-p134">The following HTML adds a graphic image, which could be a company logo. You can use a logo of your choice, or copy the NewLogo.png file from the Project 2013 SDK download, and then use  **Solution Explorer** to add the file to the `HelloProjectODataWeb\Images` folder.</span></span>

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

## <a name="creating-the-javascript-code-for-the-add-in"></a><span data-ttu-id="1e277-215">このアドインの JavaScript コードを作成する</span><span class="sxs-lookup"><span data-stu-id="1e277-215">Creating the JavaScript code for the add-in</span></span>

<span data-ttu-id="1e277-p135">Project 作業ウィンドウ アドイン用のテンプレートには、既定の初期化コードが含まれています。このコードは、典型的な Office 2013 アドイン用にドキュメント内のデータに対する基本的な取得操作と設定操作を例示するものとして設計されています。Project 2013 はアクティブなプロジェクトへの書き込み操作をサポートしておらず、 **HelloProjectOData** アドインは **getSelectedDataAsync** メソッドを使用しないので、 **Office.initialize** 関数内のスクリプトを削除できます。また、既定の HelloProjectOData.js ファイルに含まれている **setData** 関数と **getData** 関数も削除できます。</span><span class="sxs-lookup"><span data-stu-id="1e277-p135">The template for a Project task pane add-in includes default initialization code that is designed to demonstrate basic get and set actions for data in a document for a typical Office 2013 add-in. Because Project 2013 does not support actions that write to the active project, and the  **HelloProjectOData** add-in does not use the **getSelectedDataAsync** method, you can delete the script within the **Office.initialize** function, and delete the **setData** function and **getData** function in the default HelloProjectOData.js file.</span></span>

<span data-ttu-id="1e277-p136">JavaScript には、REST クエリ用のグローバル定数と、いくつかの関数で使用されるグローバル変数が含まれています。[ **ProjectData エンドポイントを取得**] ボタンによって呼び出される  **setOdataUrl** 関数は、グローバル変数を初期化し、Project が Project Web App と接続されているかどうかを調べます。</span><span class="sxs-lookup"><span data-stu-id="1e277-p136">The JavaScript includes global constants for the REST query and global variables that are used in several functions. The  **Get ProjectData Endpoint** button calls the **setOdataUrl** function, which initializes the global variables and determines whether Project is connected with Project Web App.</span></span>

<span data-ttu-id="1e277-220">HelloProjectOData.js ファイルには、 **retrieveOData** 関数と **parseODataResult** 関数も含まれています。前者の関数は、ユーザーが [ **すべてのプロジェクトを比較**] を選択したときに呼び出されます。後者の関数は、平均値を計算し、色と単位について書式設定された値を比較テーブルに内に設定します。</span><span class="sxs-lookup"><span data-stu-id="1e277-220">The remainder of the HelloProjectOData.js file includes two functions: the  **retrieveOData** function is called when the user selects **Compare All Projects**; and the  **parseODataResult** function calculates averages and then populates the comparison table with values that are formatted for color and units.</span></span>

### <a name="procedure-5-to-create-the-javascript-code"></a><span data-ttu-id="1e277-p137">手順 5. JavaScript コードを作成するには</span><span class="sxs-lookup"><span data-stu-id="1e277-p137">Procedure 5. To create the JavaScript code</span></span>

1. <span data-ttu-id="1e277-p138">既定の HelloProjectOData.js ファイル内のコードをすべて削除し、グローバル変数と  **Office.initialize** 関数を追加します。すべて大文字の変数名は、それらが定数であることを示しており、それらは後で **_pwa** 変数と共に使用されて、この例の REST クエリが作成されます。</span><span class="sxs-lookup"><span data-stu-id="1e277-p138">Delete all code in the default HelloProjectOData.js file, and then add the global variables and  **Office.initialize** function. Variable names that are all capitals imply that they are constants; they are later used with the **_pwa** variable to create the REST query in this example.</span></span>

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

2. <span data-ttu-id="1e277-p139">**setOdataUrl** 関数と関連する関数を追加します。**setOdataUrl** 関数は **getProjectGuid** と **getDocumentUrl** を呼び出して、グローバル変数を初期化します。[getProjectFieldAsync](/javascript/api/office/office.document) メソッドでは、_callback_ パラメーター用の匿名関数が、jQuery ライブラリ内の **removeAttr** メソッドを使って **[すべてのプロジェクトを比較]** ボタンを有効にし、**ProjectData** サービスの URL を表示します。Project が Project Web App と接続されていない場合、この関数はエラーをスローし、それによってポップアップ エラー メッセージが表示されます。SurfaceErrors.js ファイルには、**throwError** メソッドが含まれています。</span><span class="sxs-lookup"><span data-stu-id="1e277-p139">Add  **setOdataUrl** and related functions. The **setOdataUrl** function calls **getProjectGuid** and **getDocumentUrl** to initialize the global variables. In the [getProjectFieldAsync method](/javascript/api/office/office.document), the anonymous function for the  _callback_ parameter enables the **Compare All Projects** button by using the **removeAttr** method in the jQuery library, and then displays the URL of the **ProjectData** service. If Project is not connected with Project Web App, the function throws an error, which displays a pop-up error message. The SurfaceErrors.js file includes the **throwError** method.</span></span>

   > [!NOTE]
   > <span data-ttu-id="1e277-p140">Visual Studio を Project Server コンピューターで実行している場合、**F5** キーによるデバッグを使用するには、**_pwa** グローバル変数を初期化する行の後にあるコードをコメント解除します。Project Server コンピューターでのデバッグ時に jQuery の **ajax** メソッドを使用できるようにするには、PWA URL に **localhost** 値を設定する必要があります。Visual Studio をリモート コンピューターで実行する場合は、**localhost** は必要ありません。アドインを展開する前に、そのコードをコメント化してください。</span><span class="sxs-lookup"><span data-stu-id="1e277-p140">If you run Visual Studio on the Project Server computer, to use  **F5** debugging, uncomment the code after the line that initializes the **_pwa** global variable. To enable using the jQuery **ajax** method when debugging on the Project Server computer, you must set the **localhost** value for the PWA URL.If you run Visual Studio on a remote computer, the  **localhost** URL is not required. Before you deploy the add-in, comment out that code.</span></span>

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

3. <span data-ttu-id="1e277-p141">**retrieveOData** 関数を追加します。この関数は REST クエリ用に値を連結し、jQuery の **ajax** 関数を呼び出して、要求されたデータを **ProjectData** サービスから取得します。 **support.cors** 変数は、 **ajax** 関数でのクロス オリジン リソース共有 (CORS) を有効にします。 **support.cors** ステートメントがないか、 **false** に設定されていると、 **ajax** 関数は **No transport** エラーを返します。</span><span class="sxs-lookup"><span data-stu-id="1e277-p141">Add the  **retrieveOData** function, which concatenates values for the REST query and then calls the **ajax** function in jQuery to get the requested data from the **ProjectData** service. The **support.cors** variable enables cross-origin resource sharing (CORS) with the **ajax** function. If the **support.cors** statement is missing or is set to **false**, the  **ajax** function returns a **No transport** error.</span></span>

   > [!NOTE]
   > <span data-ttu-id="1e277-p142">次に示すコードは、Project Server 2013 のオンプレミスのインストールで動作します。Project Online の場合は、トークン ベースの認証に OAuth を使用できます。詳細については、「[Office アドインにおける同一生成元ポリシーの制限への対処](../develop/addressing-same-origin-policy-limitations.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="1e277-p142">The following code works with an on-premises installation of Project Server 2013. For Project Online, you can use OAuth for token-based authentication. For more information, see [Addressing same-origin policy limitations in Office Add-ins](../develop/addressing-same-origin-policy-limitations.md).</span></span>

   <span data-ttu-id="1e277-p143">**ajax** の呼び出しでは、_headers_ パラメーターまたは _beforeSend_ パラメーターを使用できます。_complete_ パラメーターは匿名関数であり、**retrieveOData** の変数と同じスコープ内にあります。_complete_ パラメーターの関数は、**odataText** コントロールに結果を表示すると共に、**parseODataResult** メソッドを呼び出すことにより JSON 応答を解析して表示します。_error_ パラメーターは名前付きの **getProjectDataErrorHandler** 関数を指定します。この関数は、エラー メッセージを **odataText** コントロールに書き込み、**throwError** メソッドを使用してポップアップ メッセージを表示します。</span><span class="sxs-lookup"><span data-stu-id="1e277-p143">In the **ajax** call, you can use either the _headers_ parameter or the _beforeSend_ parameter. The _complete_ parameter is an anonymous function so that it is in the same scope as the variables in **retrieveOData**. The function for the  _complete_ parameter displays results in the **odataText** control and also calls the **parseODataResult** method to parse and display the JSON response. The _error_ parameter specifies the named **getProjectDataErrorHandler** function, which writes an error message to the **odataText** control and also uses the **throwError** method to display a pop-up message.</span></span>

    ```js
    /****************************************************************
    * Functions to get and parse the Project Server reporting data.
    *****************************************************************/

    // Get data about all projects on Project Server,
    // by using a REST query with the ajax method in jQuery.
    function retrieveOData() {
        var restUrl = _odataUrl + PROJQUERY + QUERY_FILTER + QUERY_SELECT1 + QUERY_SELECT2;
        var accept = "application/json; odata=verbose";
        accept.toLocaleLowerCase();

        // Enable cross-origin scripting (required by jQuery 1.5 and later).
        // This does not work with Project Online.
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

4. <span data-ttu-id="1e277-p144">**parseODataResult** メソッドを追加します。このメソッドは OData サービスからの JSON 応答を逆シリアル化して、処理します。**parseODataResult** メソッドは、コストと作業のデータの平均値を小数点以下 1 桁または 2 桁の精度まで計算し、適切な色で値を書式設定し、単位 (**$**、**hrs**、または **%**) を追加し、指定されたテーブル セルに値を表示します。</span><span class="sxs-lookup"><span data-stu-id="1e277-p144">Add the **parseODataResult** method, which deserializes and processes the JSON response from the OData service. The **parseODataResult** method calculates average values of the cost and work data to an accuracy of one or two decimal places, formats values with the correct color and adds a unit ( **$**,  **hrs**, or  **%**), and then displays the values in specified table cells.</span></span>

   <span data-ttu-id="1e277-p145">アクティブなプロジェクトの GUID が **ProjectId** の値と一致する場合、**myProjectIndex** 変数はプロジェクトのインデックスに設定されます。アクティブなプロジェクトが Project Server で発行されていることを **myProjectIndex** が示している場合、**parseODataResult** メソッドはそのプロジェクトのコストと作業データを書式設定して表示します。アクティブなプロジェクトが発行されていない場合は、アクティブなプロジェクトの値は青い **NA** と表示されます。</span><span class="sxs-lookup"><span data-stu-id="1e277-p145">If the GUID of the active project matches the  **ProjectId** value, the **myProjectIndex** variable is set to the project index. If **myProjectIndex** indicates the active project is published on Project Server, the **parseODataResult** method formats and displays cost and work data for that project. If the active project is not published, values for the active project are displayed as a blue **NA**.</span></span>

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

## <a name="testing-the-helloprojectodata-add-in"></a><span data-ttu-id="1e277-248">HelloProjectOData アドインのテスト</span><span class="sxs-lookup"><span data-stu-id="1e277-248">Testing the HelloProjectOData add-in</span></span>

<span data-ttu-id="1e277-p146">**HelloProjectOData** アドインを Visual Studio 2015 でテストしてデバッグするには、開発用コンピューターに Project Professional 2013 がインストールされている必要があります。他のテスト シナリオを有効にするには、Project がローカル コンピューター上のファイルに対して開くか、Project Web App と接続するかを選択できることを確認してください。たとえば、次の手順を実行します。</span><span class="sxs-lookup"><span data-stu-id="1e277-p146">To test and debug the  **HelloProjectOData** add-in with Visual Studio 2015, Project Professional 2013 must be installed on the development computer. To enable different test scenarios, ensure that you can choose whether Project opens for files on the local computer or connects with Project Web App. For example, do the following steps:</span></span>

1. <span data-ttu-id="1e277-252">リボンの [ **ファイル**] タブで、Backstage ビューの [ **情報**] タブを選択し、[ **アカウントの管理**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="1e277-252">On the  **FILE** tab on the ribbon, choose the **Info** tab in the Backstage view, and then choose **Manage Accounts**.</span></span>

2. <span data-ttu-id="1e277-p147">[ **Project Web App アカウント**] ダイアログ ボックスの [ **利用可能なアカウント**] の一覧には、ローカル  **コンピューター** アカウントだけでなく、複数の Project Web App アカウントが表示されることがあります。[ **開始時**] セクションで、[ **アカウントを選択する**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="1e277-p147">In the  **Project web app Accounts** dialog box, the **Available accounts** list can have multiple Project Web App accounts in addition to the local **Computer** account. In the **When starting** section, select **Choose an account**.</span></span>

3. <span data-ttu-id="1e277-255">Project を閉じます。それにより、Visual Studio でアドインのデバッグ用に Project 起動できるようになります。</span><span class="sxs-lookup"><span data-stu-id="1e277-255">Close Project so that Visual Studio can start it for debugging the add-in.</span></span>

<span data-ttu-id="1e277-256">基本的なテストでは、次のことを行う必要があります。</span><span class="sxs-lookup"><span data-stu-id="1e277-256">Basic tests should include the following:</span></span>

- <span data-ttu-id="1e277-p148">アドインを Visual Studio から実行し、コストと作業のデータが含まれている Project Web App から発行済みのプロジェクトを開きます。アドインが  **ProjectData** エンドポイントを表示し、コストと作業のデータをテーブルに正しく表示することを確認します。 **odataText** コントロール内の出力により、REST クエリとその他の情報を確認できます。</span><span class="sxs-lookup"><span data-stu-id="1e277-p148">Run the add-in from Visual Studio, and then open a published project from Project Web App that contains cost and work data. Verify that the add-in displays the  **ProjectData** endpoint and correctly displays the cost and work data in the table. You can use the output in the **odataText** control to check the REST query and other information.</span></span>

- <span data-ttu-id="1e277-p149">アドインを再び実行し、Project の起動時に [ **ログイン**] ダイアログ ボックスでローカル コンピューターのプロファイルを選択します。ローカル .mpp ファイルを開き、アドインをテストします。 **ProjectData** エンドポイントを取得しようとしたときに、アドインがエラー メッセージを表示することを確認します。</span><span class="sxs-lookup"><span data-stu-id="1e277-p149">Run the add-in again, where you choose the local computer profile in the  **Login** dialog box when Project starts. Open a local .mpp file, and then test the add-in. Verify that the add-in displays an error message when you try to get the **ProjectData** endpoint.</span></span>

- <span data-ttu-id="1e277-p150">アドインを再び実行し、コストと作業のデータのタスクを持つプロジェクトを作成します。そのプロジェクトを Project Web App に保存することはできますが、発行はしないでください。アドインが Project Server からのデータを表示するものの、現在のプロジェクトについては  **NA** であることを確認します。</span><span class="sxs-lookup"><span data-stu-id="1e277-p150">Run the add-in again, where you create a project that has tasks with cost and work data. You can save the project to Project Web App, but don't publish it. Verify that the add-in displays data from Project Server, but  **NA** for the current project.</span></span>

### <a name="procedure-6-to-test-the-add-in"></a><span data-ttu-id="1e277-p151">手順 6. アドインをテストするには</span><span class="sxs-lookup"><span data-stu-id="1e277-p151">Procedure 6. To test the add-in</span></span>

1. <span data-ttu-id="1e277-p152">Project Professional 2013 を実行し、Project Web App と接続してから、テスト プロジェクトを作成します。ローカル リソースまたはエンタープライズ リソースにタスクを割り当て、いくつかのタスクに対して達成率のさまざまな値を設定してから、そのプロジェクトを発行します。Project を終了します。それにより、Visual Studio がアドインのデバッグ用に Project を起動できるようになります。</span><span class="sxs-lookup"><span data-stu-id="1e277-p152">Run Project Professional 2013, connect with Project Web App, and then create a test project. Assign tasks to local resources or to enterprise resources, set various values of percent complete on some tasks, and then publish the project. Quit Project, which enables Visual Studio to start Project for debugging the add-in.</span></span>

2. <span data-ttu-id="1e277-p153">Visual Studio で  **F5** キーを押します。Project Web App にログオンし、前のステップで作成したプロジェクトを開きます。このとき、読み取り専用モードで開くことも、編集モードで開くこともできます。</span><span class="sxs-lookup"><span data-stu-id="1e277-p153">In Visual Studio, press  **F5**. Log on to Project Web App, and then open the project that you created in the previous step. You can open the project in read-only mode or in edit mode.</span></span>

3. <span data-ttu-id="1e277-p154">リボンの **[プロジェクト]** タブにある **[Office アドイン]** ドロップダウン リストで、**[Hello ProjectData]** を選択します (図 5 を参照)。**[すべてのプロジェクトを比較]** ボタンが無効化されます。</span><span class="sxs-lookup"><span data-stu-id="1e277-p154">On the  **PROJECT** tab of the ribbon, in the **Office Add-ins** drop-down list, select **Hello ProjectData** (see Figure 5). The **Compare All Projects** button should be disabled.</span></span>

    <span data-ttu-id="1e277-276">*図 5. HelloProjectOData アドインの開始*</span><span class="sxs-lookup"><span data-stu-id="1e277-276">*Figure 5. Starting the HelloProjectOData add-in*</span></span>

    ![HelloProjectOData アプリのテスト](../images/pj15-hello-project-data-test-the-app.png)

4. <span data-ttu-id="1e277-p155">**[ProjectData 概要]** 作業ウィンドウで、**[ProjectData エンドポイントを取得]** を選択します。**projectDataEndPoint** 行に **ProjectData** サービスの URL が表示され、**[すべてのプロジェクトを比較]** ボタンが有効化されます (図 6 を参照)。</span><span class="sxs-lookup"><span data-stu-id="1e277-p155">In the  **Hello ProjectData** task pane, select **Get ProjectData Endpoint**. The  **projectDataEndPoint** line should show the URL of the **ProjectData** service, and the **Compare All Projects** button should be enabled (see Figure 6).</span></span>

5. <span data-ttu-id="1e277-p156">[ **すべてのプロジェクトを比較**] を選択します。アドインは、 **ProjectData** サービスからデータを取得している間、一時停止する可能性がありますが、その後、書式化された平均値と現在値をテーブルに表示します。</span><span class="sxs-lookup"><span data-stu-id="1e277-p156">Select  **Compare All Projects**. The add-in may pause while it retrieves data from the  **ProjectData** service, and then it should display the formatted average and current values in the table.</span></span>

    <span data-ttu-id="1e277-282">*図 6. REST クエリの結果の表示*</span><span class="sxs-lookup"><span data-stu-id="1e277-282">*Figure 6. Viewing results of the REST query*</span></span>

    ![REST クエリの結果の表示](../images/pj15-hello-project-data-rest-results.png)

6. <span data-ttu-id="1e277-p157">テキスト ボックス内の出力を調べます。ここには、ドキュメントのパス、REST クエリ、ステータス情報、および  **ajax** と **parseODataResult** の呼び出しからの JSON 結果が表示されるはずです。この出力は、 **parseODataResult** メソッドのコード ( `projCost += Number(res.d.results[i].ProjectCost);` など) の理解と作成とデバッグに役立ちます。</span><span class="sxs-lookup"><span data-stu-id="1e277-p157">Examine output in the text box. It should show the document path, REST query, status information, and JSON results from the calls to  **ajax** and **parseODataResult**. The output helps to understand, create, and debug code in the  **parseODataResult** method such as `projCost += Number(res.d.results[i].ProjectCost);`.</span></span>

    <span data-ttu-id="1e277-287">次に示すのは、Project Web App インスタンスの 3 つのプロジェクトの出力例です。わかりやすくするためにテキストに改行と空白を追加してあります。</span><span class="sxs-lookup"><span data-stu-id="1e277-287">Following is an example of the output with line breaks and spaces added to the text for clarity, for three projects in a Project Web App instance:</span></span>

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

7. <span data-ttu-id="1e277-p158">デバッグを停止し (**Shift + F5** キーを押す)、**F5** キーを再び押して、Project の新しいインスタンスを実行します。**[ログイン]** ダイアログ ボックスで、Project Web App ではなく、ローカル **コンピューター** のプロファイルを選択します。ローカル プロジェクト .mpp ファイルを作成するか開き、**[ProjectData 概要]** 作業ウィンドウを開き、**[ProjectData エンドポイントを取得]** を選択します。アドインが「**接続がありません!**」エラーを表示し (図 7 を参照)、**[すべてのプロジェクトを比較]** ボタンは無効化されたままになります。</span><span class="sxs-lookup"><span data-stu-id="1e277-p158">Stop debugging (press  **Shift + F5**), and then press  **F5** again to run a new instance of Project. In the **Login** dialog box, choose the local **Computer** profile, not Project Web App. Create or open a local project .mpp file, open the **Hello ProjectData** task pane, and then select **Get ProjectData Endpoint**. The add-in should show a  **No connection!** error (see Figure 7), and the **Compare All Projects** button should remain disabled.</span></span>

   <span data-ttu-id="1e277-293">*図 7. Project Web App 接続がない状態でのアドインの使用*</span><span class="sxs-lookup"><span data-stu-id="1e277-293">*Figure 7. Using the add-in without a Project web app connection*</span></span>

   ![Project Web App 接続がない状態でのアプリの使用](../images/pj15-hello-project-data-no-connection.png)

8. <span data-ttu-id="1e277-p159">デバッグを停止してから、 **F5** キーを再び押します。Project Web App にログオンし、コストと作業のデータが含まれるプロジェクトを作成します。このプロジェクトを保存することはできますが、発行しないでください。</span><span class="sxs-lookup"><span data-stu-id="1e277-p159">Stop debugging, and then press  **F5** again. Log on to Project Web App, and then create a project that contains cost and work data. You can save the project, but don't publish it.</span></span>

   <span data-ttu-id="1e277-298">**[ProjectData 概要]** 作業ウィンドウで **[すべてのプロジェクトを比較]** を選択すると、**[現在]** 列のフィールドには青い **[NA]** が表示されます (図 8 を参照)。</span><span class="sxs-lookup"><span data-stu-id="1e277-298">In the  **Hello ProjectData** task pane, when you select **Compare All Projects**, you should see a blue  **NA** for fields in the **Current** column (see Figure 8).</span></span>

   <span data-ttu-id="1e277-299">*図 8. 未発行のプロジェクトと他のプロジェクトの比較*</span><span class="sxs-lookup"><span data-stu-id="1e277-299">*Figure 8. Comparing an unpublished project with other projects*</span></span>

   ![未発行のプロジェクトと他のプロジェクトの比較](../images/pj15-hello-project-data-not-published.png)

<span data-ttu-id="1e277-p160">アドインがこれまでのテストで正常に動作しているとしても、実行する必要のあるテストはまだあります。たとえば、次のことを行います。</span><span class="sxs-lookup"><span data-stu-id="1e277-p160">Even if your add-in is working correctly in the previous tests, there are other tests that should be run. For example:</span></span>

- <span data-ttu-id="1e277-p161">Project Web App からタスクのコストと作業のデータがないプロジェクトを開きます。 [ **現在**] 列のフィールドには 0 の値が表示されるはずです。</span><span class="sxs-lookup"><span data-stu-id="1e277-p161">Open a project from Project Web App that has no cost or work data for the tasks. You should see values of zero in the fields in the  **Current** column.</span></span>

- <span data-ttu-id="1e277-305">タスクがないプロジェクトをテストします。</span><span class="sxs-lookup"><span data-stu-id="1e277-305">Test a project that has no tasks.</span></span>

- <span data-ttu-id="1e277-p162">アドインに変更を加えて発行した場合は、発行したアドインで再び同様のテストを実行する必要があります。その他の考慮事項については、「[次のステップ](#next-steps)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="1e277-p162">If you modify the add-in and publish it, you should run similar tests again with the published add-in. For other considerations, see [Next steps](#next-steps).</span></span>

> [!NOTE]
> <span data-ttu-id="1e277-p163">**ProjectData** サービスの 1 回のクエリで返すことのできるデータ量には制限があります。このデータ量はエンティティによって異なります。たとえば、**Projects** エンティティ セットでの既定の制限は 1 回のクエリで 100 プロジェクトですが、**Risks** エンティティ セットでの既定の制限は 200 プロジェクトです。運用インストールでは、**HelloProjectOData** のコード例に変更を加えて、100 プロジェクト以上のクエリを使用できるようにする必要があります。詳細については、「[次のステップ](#next-steps)」および「[Project レポート データの OData フィードにクエリを実行する](/previous-versions/office/project-odata/jj163048(v=office.15))」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="1e277-p163">There are limits to the amount of data that can be returned in one query of the  **ProjectData** service; the amount of data varies by entity. For example, the **Projects** entity set has a default limit of 100 projects per query, but the **Risks** entity set has a default limit of 200. For a production installation, the code in the **HelloProjectOData** example should be modified to enable queries of more than 100 projects. For more information, see [Next steps](#next-steps) and [Querying OData feeds for Project reporting data](/previous-versions/office/project-odata/jj163048(v=office.15)).</span></span>

## <a name="example-code-for-the-helloprojectodata-add-in"></a><span data-ttu-id="1e277-312">HelloProjectOData アドインのコード例</span><span class="sxs-lookup"><span data-stu-id="1e277-312">Example code for the HelloProjectOData add-in</span></span>

### <a name="helloprojectodatahtml-file"></a><span data-ttu-id="1e277-313">HelloProjectOData.html ファイル</span><span class="sxs-lookup"><span data-stu-id="1e277-313">HelloProjectOData.html file</span></span>

<span data-ttu-id="1e277-314">次のコードは、**HelloProjectODataWeb** プロジェクトの `Pages\HelloProjectOData.html` ファイルに収められています。</span><span class="sxs-lookup"><span data-stu-id="1e277-314">The following code is in the `Pages\HelloProjectOData.html` file of the **HelloProjectODataWeb** project.</span></span>

```HTML
<!DOCTYPE html>
<html>
    <head>
        <meta charset="UTF-8" />
        <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
        <title>Test ProjectData Service</title>

        <link rel="stylesheet" type="text/css" href="../Content/Office.css" />

        <!-- Add your CSS styles to the following file -->
        <link rel="stylesheet" type="text/css" href="../Content/App.css" />

        <!-- Use the CDN reference to the mini-version of jQuery when deploying your add-in. -->
        <!--<script src="http://ajax.aspnetcdn.com/ajax/jquery/jquery-1.9.0.min.js"></script> -->
        <script src="../Scripts/jquery-1.7.1.js"></script>

        <!-- Use the CDN reference to Office.js when deploying your add-in -->
        <!--<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>-->

        <!-- Use the local script references for Office.js to enable offline debugging -->
        <script src="../Scripts/Office/1.0/MicrosoftAjax.js"></script>
        <script src="../Scripts/Office/1.0/Office.js"></script>

        <!-- Add your JavaScript to the following files -->
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

### <a name="helloprojectodatajs-file"></a><span data-ttu-id="1e277-315">HelloProjectOData.js ファイル</span><span class="sxs-lookup"><span data-stu-id="1e277-315">HelloProjectOData.js file</span></span>

<span data-ttu-id="1e277-316">次のコードは、**HelloProjectODataWeb** プロジェクトの `Scripts\Office\HelloProjectOData.js` ファイルに収められています。</span><span class="sxs-lookup"><span data-stu-id="1e277-316">The following code is in the `Scripts\Office\HelloProjectOData.js` file of the **HelloProjectODataWeb** project.</span></span>

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

/****************************************************************
* Functions to get and parse the Project Server reporting data.
*****************************************************************/

// Get data about all projects on Project Server,
// by using a REST query with the ajax method in jQuery.
function retrieveOData() {
    var restUrl = _odataUrl + PROJQUERY + QUERY_FILTER + QUERY_SELECT1 + QUERY_SELECT2;
    var accept = "application/json; odata=verbose";
    accept.toLocaleLowerCase();

    // Enable cross-origin scripting (required by jQuery 1.5 and later).
    // This does not work with Project Online.
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

### <a name="appcss-file"></a><span data-ttu-id="1e277-317">App.css ファイル</span><span class="sxs-lookup"><span data-stu-id="1e277-317">App.css file</span></span>

<span data-ttu-id="1e277-318">次のコードは、**HelloProjectODataWeb** プロジェクトの `Content\App.css` ファイルに収められています。</span><span class="sxs-lookup"><span data-stu-id="1e277-318">The following code is in the `Content\App.css` file of the **HelloProjectODataWeb** project.</span></span>

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

### <a name="surfaceerrorsjs-file"></a><span data-ttu-id="1e277-319">SurfaceErrors.js ファイル</span><span class="sxs-lookup"><span data-stu-id="1e277-319">SurfaceErrors.js file</span></span>

<span data-ttu-id="1e277-320">SurfaceErrors.js ファイルのコードは、「[テキスト エディターを使用して Project 2013 用の作業ウィンドウ アドインを初めて作成する](../project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md)」の「_堅牢なプログラミング_」セクションからコピーできます。</span><span class="sxs-lookup"><span data-stu-id="1e277-320">You can copy code for the SurfaceErrors.js file from the _Robust Programming_ section of [Create your first task pane add-in for Project 2013 by using a text editor](../project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md).</span></span>

## <a name="next-steps"></a><span data-ttu-id="1e277-321">次の手順</span><span class="sxs-lookup"><span data-stu-id="1e277-321">Next steps</span></span>

<span data-ttu-id="1e277-p164">**HelloProjectOData** が、AppSource で販売したり、SharePoint アプリ カタログで配布したりする製品アドインであれば、その設計は異なるでしょう。たとえば、テキスト ボックスのデバッグ出力や、**ProjectData** エンドポイントを取得するためのボタンはおそらくありません。また、100 以上のプロジェクトを持つ Project Web App インスタンスを処理するには **retireveOData** 関数を書き直す必要もあるでしょう。</span><span class="sxs-lookup"><span data-stu-id="1e277-p164">If  **HelloProjectOData** were a production add-in to be sold in AppSource or distributed in a SharePoint add-in catalog, it would be designed differently. For example, there would be no debug output in a text box, and probably no button to get the **ProjectData** endpoint. You would also have to rewrite the **retireveOData** function to handle Project Web App instances that have more than 100 projects.</span></span>

<span data-ttu-id="1e277-p165">このアドインには、追加のエラー チェックと、エッジ ケースをキャッチして説明または表示するためのロジックを組み込む必要があります。たとえば、Project Web App インスタンスに、平均期間が 5 日で平均コストが $2400 になる 1000 個のプロジェクトがあって、期間が 20 日より長いのはアクティブ プロジェクトだけだとすると、コストと作業の比較は歪んだものになるでしょう。それは頻度グラフで示すことができます。期間を表示したり、同じような長さのプロジェクトを比較したり、同じ部門または異なる部門のプロジェクトを比較したりするオプションを追加するとよいでしょう。あるいは、表示するフィールドのリストからユーザーが選択できるような方法を追加することもできます。</span><span class="sxs-lookup"><span data-stu-id="1e277-p165">The add-in should contain additional error checks, plus logic to catch and explain or show edge cases. For example, if a Project Web App instance has 1000 projects with an average duration of five days and average cost of $2400, and the active project is the only one that has a duration longer than 20 days, the cost and work comparison would be skewed. That could be shown with a frequency graph. You could add options to display duration, compare similar length projects, or compare projects from the same or different departments. Or, add a way for the user to select from a list of fields to display.</span></span>

<span data-ttu-id="1e277-p166">**ProjectData** サービスの他のクエリについては、クエリ文字列の長さに制限があり、これは親コレクションから子コレクションのオブジェクトまでにクエリが取ることのできるステップ数に影響します。たとえば、**Projects** から **Tasks** へ、そしてタスク アイテムへという 2 ステップのクエリはうまく動作しますが、**Projects** から、**Tasks**、**Assignments** を経て、割り当てアイテムへという 3 ステップのクエリになると、URL の既定の最大長を超える可能性があります。詳しくは、「[Project レポート データの OData フィードにクエリを実行する](/previous-versions/office/project-odata/jj163048(v=office.15))」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="1e277-p166">For other queries of the  **ProjectData** service, there are limits to the length of the query string, which affects the number of steps that a query can take from a parent collection to an object in a child collection. For example, a two-step query of **Projects** to **Tasks** to task item works, but a three-step query such as **Projects** to **Tasks** to **Assignments** to assignment item may exceed the default maximum URL length. For more information, see [Querying OData feeds for Project reporting data](/previous-versions/office/project-odata/jj163048(v=office.15)).</span></span>

<span data-ttu-id="1e277-333">**HelloProjectOData** アドインを運用環境で使えるように変更する場合は、次の手順を実行してください。</span><span class="sxs-lookup"><span data-stu-id="1e277-333">If you modify the  **HelloProjectOData** add-in for production use, do the following steps:</span></span>

- <span data-ttu-id="1e277-334">HelloProjectOData.html ファイルで、パフォーマンスを向上させるために、office.js の参照をローカル プロジェクトから CDN の参照に変更します。</span><span class="sxs-lookup"><span data-stu-id="1e277-334">In the HelloProjectOData.html file, for better performance, change the office.js reference from the local project to the CDN reference:</span></span>

    ```HTML
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
    ```

- <span data-ttu-id="1e277-p167">**retrieveOData** 関数を書き換えて、100 以上のプロジェクトのクエリを使えるようにします。たとえば、`~/ProjectData/Projects()/$count` クエリでプロジェクトの数を取得し、プロジェクト データの REST クエリで _$skip_ 操作と _$top_ 操作を使用します。ループの中で複数のクエリを実行し、各クエリからのデータを平均化します。プロジェクト データの各クエリは、次のような形式になります。</span><span class="sxs-lookup"><span data-stu-id="1e277-p167">Rewrite the  **retrieveOData** function to enable queries of more than 100 projects. For example, you could get the number of projects with a `~/ProjectData/Projects()/$count` query, and use the _$skip_ operator and _$top_ operator in the REST query for project data. Run multiple queries in a loop, and then average the data from each query. Each query for project data would be of the form:</span></span> 

  `~/ProjectData/Projects()?skip= [numSkipped]&amp;$top=100&amp;$filter=[filter]&amp;$select=[field1,field2, ???????]`

  <span data-ttu-id="1e277-p168">For more information, see [OData System Query Options Using the REST Endpoint](/previous-versions/dynamicscrm-2015/developers-guide/gg309461(v=crm.7)). You can also use the [Set-SPProjectOdataConfiguration](/powershell/module/sharepoint-server/Set-SPProjectOdataConfiguration?view=sharepoint-ps) command in Windows PowerShell to override the default page size for a query of the **Projects** entity set (or any of the 33 entity sets). See [ProjectData - Project OData service reference](/previous-versions/office/project-odata/jj163015(v=office.15)).</span><span class="sxs-lookup"><span data-stu-id="1e277-p168">For more information, see [OData System Query Options Using the REST Endpoint](/previous-versions/dynamicscrm-2015/developers-guide/gg309461(v=crm.7)). You can also use the [Set-SPProjectOdataConfiguration](/powershell/module/sharepoint-server/Set-SPProjectOdataConfiguration?view=sharepoint-ps) command in Windows PowerShell to override the default page size for a query of the **Projects** entity set (or any of the 33 entity sets). See [ProjectData - Project OData service reference](/previous-versions/office/project-odata/jj163015(v=office.15)).</span></span>

- <span data-ttu-id="1e277-342">アドインを展開するには、「[Office アドインを発行する](../publish/publish.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="1e277-342">To deploy the add-in, see [Publish your Office Add-in](../publish/publish.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="1e277-343">関連項目</span><span class="sxs-lookup"><span data-stu-id="1e277-343">See also</span></span>

- [<span data-ttu-id="1e277-344">Project 用の作業ウィンドウ アドイン</span><span class="sxs-lookup"><span data-stu-id="1e277-344">Task pane add-ins for Project</span></span>](project-add-ins.md)
- [<span data-ttu-id="1e277-345">テキスト エディターを使用して Project 2013 用の作業ウィンドウ アドインを初めて作成する</span><span class="sxs-lookup"><span data-stu-id="1e277-345">Create your first task pane add-in for Project 2013 by using a text editor</span></span>](create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md)
- <span data-ttu-id="1e277-346">[ProjectData - Project OData サービス リファレンス](/previous-versions/office/project-odata/jj163015(v=office.15))</span><span class="sxs-lookup"><span data-stu-id="1e277-346">[ProjectData - Project OData service reference](/previous-versions/office/project-odata/jj163015(v=office.15))</span></span>
- [<span data-ttu-id="1e277-347">Office アドインの XML マニフェスト</span><span class="sxs-lookup"><span data-stu-id="1e277-347">Office Add-ins XML manifest</span></span>](../develop/add-in-manifests.md)
- [<span data-ttu-id="1e277-348">Office アドインを発行する</span><span class="sxs-lookup"><span data-stu-id="1e277-348">Publish your Office Add-in</span></span>](../publish/publish.md)
