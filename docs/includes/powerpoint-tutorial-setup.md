<span data-ttu-id="f9db6-101">開発プロジェクトを設定して、このチュートリアルを始めます。</span><span class="sxs-lookup"><span data-stu-id="f9db6-101">You'll begin this tutorial by setting up your development project.</span></span> 

> [!NOTE]
> <span data-ttu-id="f9db6-102">このページでは、PowerPoint アドインのチュートリアルの個々の手順について説明します。</span><span class="sxs-lookup"><span data-stu-id="f9db6-102">This page describes an individual step of the PowerPoint add-in tutorial.</span></span> <span data-ttu-id="f9db6-103">このページに検索エンジンの結果から、または直接リンクからアクセスした場合は、「[PowerPoint アドインのチュートリアル](../tutorials/powerpoint-tutorial.yml)」の紹介ページに移動し、チュートリアルを最初から始めてください。</span><span class="sxs-lookup"><span data-stu-id="f9db6-103">If you’ve arrived at this page via search engine results or other direct link, please go to the [PowerPoint add-in tutorial](../tutorials/powerpoint-tutorial.yml) introduction page to start the tutorial from the beginning.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="f9db6-104">前提条件</span><span class="sxs-lookup"><span data-stu-id="f9db6-104">Prerequisites</span></span>

[!include[Quick Start prerequisites](../includes/quickstart-vs-prerequisites.md)]

## <a name="setup"></a><span data-ttu-id="f9db6-105">セットアップ</span><span class="sxs-lookup"><span data-stu-id="f9db6-105">Setup</span></span>

<span data-ttu-id="f9db6-106">このチュートリアルでは、Visual Studio を使用してアドインを作成します。</span><span class="sxs-lookup"><span data-stu-id="f9db6-106">In this tutorial, you'll create an add-in using Visual Studio.</span></span>

### <a name="create-the-add-in-project"></a><span data-ttu-id="f9db6-107">アドイン プロジェクトの作成</span><span class="sxs-lookup"><span data-stu-id="f9db6-107">Create the add-in project</span></span>

1. <span data-ttu-id="f9db6-108">[Visual Studio] メニュー バーで、**[ファイル]** > **[新規作成]** > **[プロジェクト]** の順に選択します。</span><span class="sxs-lookup"><span data-stu-id="f9db6-108">On the Visual Studio menu bar, choose  **File** > **New** > **Project**.</span></span>
    
2. <span data-ttu-id="f9db6-109">**[Visual C#]** または **[Visual Basic]** の下にあるプロジェクトの種類の一覧で、**[Office/SharePoint]** を展開して、**[アドイン]** を選択し、プロジェクトの種類として **[PowerPoint Web アドイン]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="f9db6-109">In the list of project types under **Visual C#** or **Visual Basic**, expand  **Office/SharePoint**, choose **Add-ins**, and then choose **PowerPoint Web Add-in** as the project type.</span></span> 

3. <span data-ttu-id="f9db6-110">プロジェクトに **HelloWorld** という名前を付けて、**[OK]** ボタンをクリックします。</span><span class="sxs-lookup"><span data-stu-id="f9db6-110">Name the project **HelloWorld**, and then choose the **OK** button.</span></span>

4. <span data-ttu-id="f9db6-111">**[Office アドインの作成]** ダイアログ ウィンドウで、**[新機能を PowerPoint に追加する]** を選択してから、**[完了]** を選択してプロジェクトを作成します。</span><span class="sxs-lookup"><span data-stu-id="f9db6-111">In the **Create Office Add-in** dialog window, choose **Add new functionalities to PowerPoint**, and then choose **Finish** to create the project.</span></span>

5. <span data-ttu-id="f9db6-p102">Visual Studio によってソリューションとその 2 つのプロジェクトが作成され、**ソリューション エクスプローラー**に表示されます。**Home.html** ファイルが Visual Studio で開かれます。</span><span class="sxs-lookup"><span data-stu-id="f9db6-p102">Visual Studio creates a solution and its two projects appear in **Solution Explorer**. The **Home.html** file opens in Visual Studio.</span></span>

     ![PowerPoint チュートリアル - HelloWorld ソリューションで 2 つのプロジェクトを表示する Visual Studio ソリューション エクスプローラー ウィンドウ](../images/powerpoint-tutorial-solution-explorer.png)

### <a name="explore-the-visual-studio-solution"></a><span data-ttu-id="f9db6-115">Visual Studio ソリューションについて理解する</span><span class="sxs-lookup"><span data-stu-id="f9db6-115">Explore the Visual Studio solution</span></span>

[!include[Description of Visual Studio projects](../includes/quickstart-vs-solution.md)]

### <a name="update-code"></a><span data-ttu-id="f9db6-116">コードを更新する</span><span class="sxs-lookup"><span data-stu-id="f9db6-116">Update code</span></span> 

<span data-ttu-id="f9db6-117">アドイン コードを次のように編集し、このチュートリアルの後続の手順でアドイン機能を実装するために使用するフレームワークを作成します。</span><span class="sxs-lookup"><span data-stu-id="f9db6-117">Edit the add-in code as follows, to create the framework that you'll use to implement add-in functionality in subsequent steps of this tutorial.</span></span>

1. <span data-ttu-id="f9db6-118">**Home.html** では、アドインの作業ウィンドウにレンダリングされる HTML を指定します。</span><span class="sxs-lookup"><span data-stu-id="f9db6-118">**Home.html** specifies the HTML that will be rendered in the add-in's task pane.</span></span> <span data-ttu-id="f9db6-119">**Home.html** において、`id="content-main"` で **div** を検索し、**div** 全体を次のマークアップと置き換えてファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="f9db6-119">In **Home.html**, find the **div** with `id="content-main"`, replace that entire **div** with the following markup, and save the file.</span></span>

    ```html
    <!-- TODO2: Create the content-header div. -->
    <div id="content-main">
        <div class="padding">
            <!-- TODO1: Create the insert-image button. -->
            <!-- TODO3: Create the insert-text button. -->
            <!-- TODO4: Create the get-slide-metadata button. -->
            <!-- TODO5: Create the go-to-slide buttons. -->
        </div>
    </div>
    ```

2. <span data-ttu-id="f9db6-120">Web アプリケーション プロジェクトのルートにあるファイル **Home.js** を開きます。</span><span class="sxs-lookup"><span data-stu-id="f9db6-120">Open the file **Home.js** in the root of the web application project.</span></span> <span data-ttu-id="f9db6-121">このファイルは、アドイン用のスクリプトを指定します。</span><span class="sxs-lookup"><span data-stu-id="f9db6-121">This file specifies the script for the add-in.</span></span> <span data-ttu-id="f9db6-122">すべての内容を次のコードに置き換え、ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="f9db6-122">Replace the entire contents with the following code and save the file.</span></span>

    ```javascript
    (function () {
        "use strict";

        var messageBanner;

        Office.initialize = function (reason) {
            $(document).ready(function () {
                // Initialize the FabricUI notification mechanism and hide it
                var element = document.querySelector('.ms-MessageBanner');
                messageBanner = new fabric.MessageBanner(element);
                messageBanner.hideBanner();

                // TODO1: Assign event handler for insert-image button.
                // TODO4: Assign event handler for insert-text button.
                // TODO6: Assign event handler for get-slide-metadata button.
                // TODO8: Assign event handlers for the four navigation buttons.
            });
        };

        // TODO2: Define the insertImage function. 

        // TODO3: Define the insertImageFromBase64String function.

        // TODO5: Define the insertText function.

        // TODO7: Define the getSlideMetadata function.

        // TODO9: Define the navigation functions.

        // Helper function for displaying notifications
        function showNotification(header, content) {
            $("#notification-header").text(header);
            $("#notification-body").text(content);
            messageBanner.showBanner();
            messageBanner.toggleExpansion();
        }
    })();
    ```
