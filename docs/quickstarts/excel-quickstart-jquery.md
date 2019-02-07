---
title: jQuery を使用して初めての Excel アドインをビルドする
description: ''
ms.date: 01/17/2019
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: 7931e06899a94a0dcda2a5ab442d37ce21c119c0
ms.sourcegitcommit: 33dcf099c6b3d249811580d67ee9b790c0fdccfb
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/05/2019
ms.locfileid: "29742444"
---
# <a name="build-an-excel-add-in-using-jquery"></a><span data-ttu-id="b091d-102">jQuery を使用して Excel のアドインを作成する</span><span class="sxs-lookup"><span data-stu-id="b091d-102">Build an Excel add-in using jQuery</span></span>

<span data-ttu-id="b091d-103">この記事では、jQuery と Excel の JavaScript API を使用して Excel アドインを構築する手順について説明します。</span><span class="sxs-lookup"><span data-stu-id="b091d-103">In this article, you'll walk through the process of building an Excel add-in by using jQuery and the Excel JavaScript API.</span></span> 

## <a name="create-the-add-in"></a><span data-ttu-id="b091d-104">アドインを作成する</span><span class="sxs-lookup"><span data-stu-id="b091d-104">Create the add-in</span></span> 

[!include[Choose your editor](../includes/quickstart-choose-editor.md)]

# <a name="visual-studiotabvisual-studio"></a>[<span data-ttu-id="b091d-105">Visual Studio</span><span class="sxs-lookup"><span data-stu-id="b091d-105">Visual Studio</span></span>](#tab/visual-studio)

### <a name="prerequisites"></a><span data-ttu-id="b091d-106">前提条件</span><span class="sxs-lookup"><span data-stu-id="b091d-106">Prerequisites</span></span>

[!include[Quick Start prerequisites](../includes/quickstart-vs-prerequisites.md)]

### <a name="create-the-add-in-project"></a><span data-ttu-id="b091d-107">アドイン プロジェクトの作成</span><span class="sxs-lookup"><span data-stu-id="b091d-107">Create the add-in project</span></span>

1. <span data-ttu-id="b091d-108">[Visual Studio] メニュー バーで、**[ファイル]**  >  **[新規作成]**  >  **[プロジェクト]** の順に選択します。</span><span class="sxs-lookup"><span data-stu-id="b091d-108">On the Visual Studio menu bar, choose  **File** > **New** > **Project**.</span></span>
    
2. <span data-ttu-id="b091d-109">**[Visual C#]** または **[Visual Basic]** の下にあるプロジェクトの種類の一覧で、**[Office/SharePoint]** を展開して、**[アドイン]** を選択し、プロジェクトの種類として **[Excel Web アドイン]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="b091d-109">In the list of project types under **Visual C#** or **Visual Basic**, expand  **Office/SharePoint**, choose **Add-ins**, and then choose **Excel Web Add-in** as the project type.</span></span> 

3. <span data-ttu-id="b091d-110">プロジェクトに名前を付けて、**[OK]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="b091d-110">Name the project, and then choose **OK**.</span></span>

4. <span data-ttu-id="b091d-111">**[Office アドインの作成]** ダイアログ ウィンドウで、**[新機能を Excel に追加する]** を選択してから、**[完了]** を選択してプロジェクトを作成します。</span><span class="sxs-lookup"><span data-stu-id="b091d-111">In the **Create Office Add-in** dialog window, choose **Add new functionalities to Excel**, and then choose **Finish** to create the project.</span></span>

5. <span data-ttu-id="b091d-p101">Visual Studio によってソリューションとその 2 つのプロジェクトが作成され、**ソリューション エクスプローラー**に表示されます。**Home.html** ファイルが Visual Studio で開かれます。</span><span class="sxs-lookup"><span data-stu-id="b091d-p101">Visual Studio creates a solution and its two projects appear in **Solution Explorer**. The **Home.html** file opens in Visual Studio.</span></span>
    
### <a name="explore-the-visual-studio-solution"></a><span data-ttu-id="b091d-114">Visual Studio ソリューションについて理解する</span><span class="sxs-lookup"><span data-stu-id="b091d-114">Explore the Visual Studio solution</span></span>

[!include[Description of Visual Studio projects](../includes/quickstart-vs-solution.md)]

### <a name="update-the-code"></a><span data-ttu-id="b091d-115">コードを更新する</span><span class="sxs-lookup"><span data-stu-id="b091d-115">Update the code</span></span>

1. <span data-ttu-id="b091d-116">**Home.html** では、アドインの作業ウィンドウにレンダリングされる HTML を指定します。</span><span class="sxs-lookup"><span data-stu-id="b091d-116">**Home.html** specifies the HTML that will be rendered in the add-in's task pane.</span></span> <span data-ttu-id="b091d-117">**Home.html** で、`<body>` 要素を次のマークアップに置き換えて、ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="b091d-117">In **Home.html**, replace the `<body>` element with the following markup and save the file.</span></span>
 
    ```html
    <body class="ms-font-m ms-welcome">
        <div id="content-header">
            <div class="padding">
                <h1>Welcome</h1>
            </div>
        </div>
        <div id="content-main">
            <div class="padding">
                <p>Choose the button below to set the color of the selected range to green.</p>
                <br />
                <h3>Try it out</h3>
                <button class="ms-Button" id="set-color">Set color</button>
            </div>
        </div>
    </body>
    ```

2. <span data-ttu-id="b091d-118">Web アプリケーション プロジェクトのルートにあるファイル **Home.js** を開きます。</span><span class="sxs-lookup"><span data-stu-id="b091d-118">Open the file **Home.js** in the root of the web application project.</span></span> <span data-ttu-id="b091d-119">このファイルは、アドイン用のスクリプトを指定します。</span><span class="sxs-lookup"><span data-stu-id="b091d-119">This file specifies the script for the add-in.</span></span> <span data-ttu-id="b091d-120">すべての内容を次のコードに置き換え、ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="b091d-120">Replace the entire contents with the following code and save the file.</span></span> 

    ```js
    'use strict';

    (function () {

        Office.onReady(function() {
            // Office is ready
            $(document).ready(function () {
                // The document is ready
                $('#set-color').click(setColor);
            });
        });

        function setColor() {
            Excel.run(function (context) {
                var range = context.workbook.getSelectedRange();
                range.format.fill.color = 'green';

                return context.sync();
            }).catch(function (error) {
                console.log("Error: " + error);
                if (error instanceof OfficeExtension.Error) {
                    console.log("Debug info: " + JSON.stringify(error.debugInfo));
                }
            });
        }
    })();
    ```

3. <span data-ttu-id="b091d-121">Web アプリケーション プロジェクトのルートにあるファイル **Home.css** を開きます。</span><span class="sxs-lookup"><span data-stu-id="b091d-121">Open the file **Home.css** in the root of the web application project.</span></span> <span data-ttu-id="b091d-122">このファイルは、アドイン用のユーザー設定のスタイルを指定します。</span><span class="sxs-lookup"><span data-stu-id="b091d-122">This file specifies the custom styles for the add-in.</span></span> <span data-ttu-id="b091d-123">すべての内容を次のコードに置き換え、ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="b091d-123">Replace the entire contents with the following code and save the file.</span></span> 

    ```css
    #content-header {
        background: #2a8dd4;
        color: #fff;
        position: absolute;
        top: 0;
        left: 0;
        width: 100%;
        height: 80px; 
        overflow: hidden;
    }

    #content-main {
        background: #fff;
        position: fixed;
        top: 80px;
        left: 0;
        right: 0;
        bottom: 0;
        overflow: auto; 
    }

    .padding {
        padding: 15px;
    }
    ```

### <a name="update-the-manifest"></a><span data-ttu-id="b091d-124">マニフェストを更新する</span><span class="sxs-lookup"><span data-stu-id="b091d-124">Update the manifest</span></span>

1. <span data-ttu-id="b091d-125">アドイン プロジェクト内の XML マニフェスト ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="b091d-125">Open the XML manifest file in the add-in project.</span></span> <span data-ttu-id="b091d-126">このファイルは、アドインの設定と機能を定義します。</span><span class="sxs-lookup"><span data-stu-id="b091d-126">This file defines the add-in's settings and capabilities.</span></span>

2. <span data-ttu-id="b091d-127">`ProviderName` 要素にはプレースホルダー値が含まれています。</span><span class="sxs-lookup"><span data-stu-id="b091d-127">The `ProviderName` element has a placeholder value.</span></span> <span data-ttu-id="b091d-128">それを自分の名前に置き換えます。</span><span class="sxs-lookup"><span data-stu-id="b091d-128">Replace it with your name.</span></span>

3. <span data-ttu-id="b091d-129">`DisplayName` 要素の `DefaultValue` 属性にはプレースホルダー値が含まれています。</span><span class="sxs-lookup"><span data-stu-id="b091d-129">The `DefaultValue` attribute of the `DisplayName` element has a placeholder.</span></span> <span data-ttu-id="b091d-130">これは、**My Office Add-in** に置き換えてください。</span><span class="sxs-lookup"><span data-stu-id="b091d-130">Replace it with **My Office Add-in**.</span></span>

4. <span data-ttu-id="b091d-131">`Description` 要素の `DefaultValue` 属性にはプレースホルダー値が含まれています。</span><span class="sxs-lookup"><span data-stu-id="b091d-131">The `DefaultValue` attribute of the `Description` element has a placeholder.</span></span> <span data-ttu-id="b091d-132">これは、**A task pane add-in for Excel** に置き換えてください。</span><span class="sxs-lookup"><span data-stu-id="b091d-132">Replace it with **A task pane add-in for Excel**.</span></span>

5. <span data-ttu-id="b091d-133">ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="b091d-133">Save the file.</span></span>

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A task pane add-in for Excel"/>
    ...
    ```

### <a name="try-it-out"></a><span data-ttu-id="b091d-134">試してみる</span><span class="sxs-lookup"><span data-stu-id="b091d-134">Try it out</span></span>

1. <span data-ttu-id="b091d-p109">Visual Studio を使用して、新しく作成した Excel アドインをテストします。そのために、**F5** キーを押すか **[開始]** ボタンをクリックして、リボンに **[作業ウィンドウの表示]** アドイン ボタンが表示された Excel を起動します。アドインは IIS 上でローカルにホストされます。</span><span class="sxs-lookup"><span data-stu-id="b091d-p109">Using Visual Studio, test the newly created Excel add-in by pressing **F5** or choosing the **Start** button to launch Excel with the **Show Taskpane** add-in button displayed in the ribbon. The add-in will be hosted locally on IIS.</span></span>

2. <span data-ttu-id="b091d-137">Excel で、**[ホーム]** タブを選択し、リボンの **[作業ウィンドウの表示]** ボタンをクリックして、アドインの作業ウィンドウを開きます。</span><span class="sxs-lookup"><span data-stu-id="b091d-137">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Excel アドイン ボタン](../images/excel-quickstart-addin-2a.png)

3. <span data-ttu-id="b091d-139">ワークシート内で任意のセルの範囲を選択します。</span><span class="sxs-lookup"><span data-stu-id="b091d-139">Select any range of cells in the worksheet.</span></span>

4. <span data-ttu-id="b091d-140">作業ウィンドウで、**[色の設定]** ボタンをクリックして、選択範囲の色を緑に設定します。</span><span class="sxs-lookup"><span data-stu-id="b091d-140">In the task pane, choose the **Set color** button to set the color of the selected range to green.</span></span>

    ![Excel アドイン](../images/excel-quickstart-addin-2c.png)

# <a name="any-editortabvisual-studio-code"></a>[<span data-ttu-id="b091d-142">任意のエディター</span><span class="sxs-lookup"><span data-stu-id="b091d-142">Any editor</span></span>](#tab/visual-studio-code)

### <a name="prerequisites"></a><span data-ttu-id="b091d-143">前提条件</span><span class="sxs-lookup"><span data-stu-id="b091d-143">Prerequisites</span></span>

- [<span data-ttu-id="b091d-144">Node.js</span><span class="sxs-lookup"><span data-stu-id="b091d-144">Node.js</span></span>](https://nodejs.org)

- <span data-ttu-id="b091d-145">[Yeoman](https://github.com/yeoman/yo) の最新バージョンと [Office アドイン用の Yeoman ジェネレーター](https://github.com/OfficeDev/generator-office)をグローバルにインストールします。</span><span class="sxs-lookup"><span data-stu-id="b091d-145">Install the latest version of [Yeoman](https://github.com/yeoman/yo) and the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) globally.</span></span>
    ```bash
    npm install -g yo generator-office
    ```

### <a name="create-the-web-app"></a><span data-ttu-id="b091d-146">Web アプリを作成する</span><span class="sxs-lookup"><span data-stu-id="b091d-146">Create the web app</span></span>

1. <span data-ttu-id="b091d-147">Yeoman ジェネレーターを使用して、Excel アドイン プロジェクトを作成します。</span><span class="sxs-lookup"><span data-stu-id="b091d-147">Use the Yeoman generator to create an Excel add-in project.</span></span> <span data-ttu-id="b091d-148">次のコマンドを実行し、以下のプロンプトに応答します。</span><span class="sxs-lookup"><span data-stu-id="b091d-148">Run the following command and then answer the prompts as follows:</span></span>

    ```bash
    yo office
    ```

    - <span data-ttu-id="b091d-149">**Choose a project type: (プロジェクトの種類を選択)** `Office Add-in project using Jquery framework`</span><span class="sxs-lookup"><span data-stu-id="b091d-149">**Choose a project type:** `Office Add-in project using Jquery framework`</span></span>
    - <span data-ttu-id="b091d-150">**Choose a script type: (スクリプトの種類を選択)** `Javascript`</span><span class="sxs-lookup"><span data-stu-id="b091d-150">**Choose a script type:** `Javascript`</span></span>
    - <span data-ttu-id="b091d-151">**What would you want to name your add-in?: (アドインの名前を何にしますか)** `My Office Add-in`</span><span class="sxs-lookup"><span data-stu-id="b091d-151">**What do you want to name your add-in?:** `My Office Add-in`</span></span>
    - <span data-ttu-id="b091d-152">**Which Office client application would you like to support?: (どの Office クライアント アプリケーションをサポートしますか)** `Excel`</span><span class="sxs-lookup"><span data-stu-id="b091d-152">**Which Office client application would you like to support?:** `Excel`</span></span>

    ![Yeoman ジェネレーター](../images/yo-office-jquery.png)
    
    <span data-ttu-id="b091d-154">ウィザードを完了すると、ジェネレーターによってプロジェクトが作成されて、サポートしているノード コンポーネントがインストールされます。</span><span class="sxs-lookup"><span data-stu-id="b091d-154">After you complete the wizard, the generator will create the project and install supporting Node components.</span></span>

2. <span data-ttu-id="b091d-155">プロジェクトのルート フォルダーに移動します。</span><span class="sxs-lookup"><span data-stu-id="b091d-155">Navigate to the root folder of the project.</span></span>

    ```bash
    cd "My Office Add-in"
    ```

### <a name="update-the-code"></a><span data-ttu-id="b091d-156">コードを更新する</span><span class="sxs-lookup"><span data-stu-id="b091d-156">Update the code</span></span> 

1. <span data-ttu-id="b091d-157">コード エディターで、プロジェクトのルートにある **index.html** を開きます。</span><span class="sxs-lookup"><span data-stu-id="b091d-157">In your code editor, open **index.html** in the root of the project.</span></span> <span data-ttu-id="b091d-158">このファイルでは、アドインの作業ウィンドウにレンダリングされる HTML を指定します。</span><span class="sxs-lookup"><span data-stu-id="b091d-158">This file specifies the HTML that will be rendered in the add-in's task pane.</span></span> 
 
2. <span data-ttu-id="b091d-159">**index.html** 内で、`body` タグを次に示すマークアップに置き換え、ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="b091d-159">Within **index.html**, replace the `body` tag with the following markup and save the file.</span></span>
 
    ```html
    <body class="ms-font-m ms-welcome">
        <div id="content-header">
            <div class="padding">
                <h1>Welcome</h1>
            </div>
        </div>
        <div id="content-main">
            <div class="padding">
                <p>Choose the button below to set the color of the selected range to green.</p>
                <br />
                <h3>Try it out</h3>
                <button class="ms-Button" id="set-color">Set color</button>
            </div>
        </div>
        <script type="text/javascript" src="node_modules/jquery/dist/jquery.js"></script>
        <script type="text/javascript" src="node_modules/office-ui-fabric-js/dist/js/fabric.js"></script>
    </body>    
    ```

3. <span data-ttu-id="b091d-160">**src\index.js** ファイルを開いて、アドインのスクリプトを指定します。</span><span class="sxs-lookup"><span data-stu-id="b091d-160">Open the file **src\index.js** to specify the script for the add-in.</span></span> <span data-ttu-id="b091d-161">すべての内容を次のコードに置き換え、ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="b091d-161">Replace the entire contents with the following code and save the file.</span></span>

    ```js
    'use strict';
    
    (function () {
        Office.onReady(function() {
            // Office is ready
            $(document).ready(function () {
                // The document is ready
                $('#set-color').click(setColor);
            });
        });

        function setColor() {
            Excel.run(function (context) {
                var range = context.workbook.getSelectedRange();
                range.format.fill.color = 'green';

                return context.sync();
            }).catch(function (error) {
                console.log("Error: " + error);
                if (error instanceof OfficeExtension.Error) {
                    console.log("Debug info: " + JSON.stringify(error.debugInfo));
                }
            });
        }
    })();
    ```

4. <span data-ttu-id="b091d-162">**app.css** ファイルを開いて、アドインのカスタム スタイルを指定します。</span><span class="sxs-lookup"><span data-stu-id="b091d-162">Open the file **app.css** to specify the custom styles for the add-in.</span></span> <span data-ttu-id="b091d-163">すべての内容を次のコードに置き換え、ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="b091d-163">Replace the entire contents with the following code and save the file.</span></span>

    ```css
    #content-header {
        background: #2a8dd4;
        color: #fff;
        position: absolute;
        top: 0;
        left: 0;
        width: 100%;
        height: 80px; 
        overflow: hidden;
    }

    #content-main {
        background: #fff;
        position: fixed;
        top: 80px;
        left: 0;
        right: 0;
        bottom: 0;
        overflow: auto; 
    }

    .padding {
        padding: 15px;
    }
    ```

### <a name="update-the-manifest"></a><span data-ttu-id="b091d-164">マニフェストを更新する</span><span class="sxs-lookup"><span data-stu-id="b091d-164">Update the manifest</span></span>

1. <span data-ttu-id="b091d-165">**manifest.xml** ファイルを開いて、アドインの設定と機能を定義します。</span><span class="sxs-lookup"><span data-stu-id="b091d-165">Open the file **manifest.xml** to define the add-in's settings and capabilities.</span></span> 

2. <span data-ttu-id="b091d-166">`ProviderName` 要素にはプレースホルダー値が含まれています。</span><span class="sxs-lookup"><span data-stu-id="b091d-166">The `ProviderName` element has a placeholder value.</span></span> <span data-ttu-id="b091d-167">それを自分の名前に置き換えます。</span><span class="sxs-lookup"><span data-stu-id="b091d-167">Replace it with your name.</span></span>

3. <span data-ttu-id="b091d-168">`Description` 要素の `DefaultValue` 属性にはプレースホルダー値が含まれています。</span><span class="sxs-lookup"><span data-stu-id="b091d-168">The `DefaultValue` attribute of the `Description` element has a placeholder.</span></span> <span data-ttu-id="b091d-169">これは、**A task pane add-in for Excel** に置き換えてください。</span><span class="sxs-lookup"><span data-stu-id="b091d-169">Replace it with **A task pane add-in for Excel**.</span></span>

4. <span data-ttu-id="b091d-170">ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="b091d-170">Save the file.</span></span>

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A task pane add-in for Excel"/>
    ...
    ```

### <a name="start-the-dev-server"></a><span data-ttu-id="b091d-171">開発用サーバーの起動</span><span class="sxs-lookup"><span data-stu-id="b091d-171">Start the dev server</span></span>

[!include[Start server section](../includes/quickstart-yo-start-server.md)] 

### <a name="try-it-out"></a><span data-ttu-id="b091d-172">試してみる</span><span class="sxs-lookup"><span data-stu-id="b091d-172">Try it out</span></span>

1. <span data-ttu-id="b091d-173">アドインを実行して、Excel 内のアドインをサイドロードするのに使用するプラットフォームの手順に従います。</span><span class="sxs-lookup"><span data-stu-id="b091d-173">Follow the instructions for the platform you'll use to run your add-in to sideload the add-in within Excel.</span></span>

    - <span data-ttu-id="b091d-174">Windows: [Windows で Office アドインをサイドロードする](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="b091d-174">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>
    - <span data-ttu-id="b091d-175">Excel Online:[Office Online で Office アドインをサイドロードする](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)</span><span class="sxs-lookup"><span data-stu-id="b091d-175">Excel Online: [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)</span></span>
    - <span data-ttu-id="b091d-176">iPad および Mac: [iPad と Mac で Office アドインをサイドロードする](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="b091d-176">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>

2. <span data-ttu-id="b091d-177">Excel で、**[ホーム]** タブを選択し、リボンの **[作業ウィンドウの表示]** ボタンをクリックして、アドインの作業ウィンドウを開きます。</span><span class="sxs-lookup"><span data-stu-id="b091d-177">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Excel アドイン ボタン](../images/excel-quickstart-addin-2b.png)

3. <span data-ttu-id="b091d-179">ワークシート内で任意のセルの範囲を選択します。</span><span class="sxs-lookup"><span data-stu-id="b091d-179">Select any range of cells in the worksheet.</span></span>

4. <span data-ttu-id="b091d-180">作業ウィンドウで、**[色の設定]** ボタンをクリックして、選択範囲の色を緑に設定します。</span><span class="sxs-lookup"><span data-stu-id="b091d-180">In the task pane, choose the **Set color** button to set the color of the selected range to green.</span></span>

    ![Excel アドイン](../images/excel-quickstart-addin-2c.png)

---

## <a name="next-steps"></a><span data-ttu-id="b091d-182">次の手順</span><span class="sxs-lookup"><span data-stu-id="b091d-182">Next steps</span></span>

<span data-ttu-id="b091d-p116">これで完了です。jQuery を使用して Excel アドインが正常に作成されました。次に、Excel アドインの機能の詳細について説明します。Excel アドインのチュートリアルに従って、より複雑なアドインをビルドします。</span><span class="sxs-lookup"><span data-stu-id="b091d-p116">Congratulations, you've successfully created an Excel add-in using jQuery! Next, learn more about the capabilities of an Excel add-in and build a more complex add-in by following along with the Excel add-in tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="b091d-185">Excel アドインのチュートリアル</span><span class="sxs-lookup"><span data-stu-id="b091d-185">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial.md)

## <a name="see-also"></a><span data-ttu-id="b091d-186">関連項目</span><span class="sxs-lookup"><span data-stu-id="b091d-186">See also</span></span>

* [<span data-ttu-id="b091d-187">Excel アドインのチュートリアル</span><span class="sxs-lookup"><span data-stu-id="b091d-187">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial-create-table.md)
* [<span data-ttu-id="b091d-188">Excel JavaScript API を使用した基本的なプログラミングの概念</span><span class="sxs-lookup"><span data-stu-id="b091d-188">Fundamental programming concepts with the Excel JavaScript API</span></span>](../excel/excel-add-ins-core-concepts.md)
* [<span data-ttu-id="b091d-189">Excel アドインのコード サンプル</span><span class="sxs-lookup"><span data-stu-id="b091d-189">Excel add-in code samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [<span data-ttu-id="b091d-190">Excel JavaScript API リファレンス</span><span class="sxs-lookup"><span data-stu-id="b091d-190">Excel JavaScript API reference</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview)

